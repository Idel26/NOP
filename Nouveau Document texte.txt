
from __future__ import annotations

import io
import re
import zipfile
import html
import mimetypes
import imaplib
from copy import copy
from email.message import EmailMessage
from email.utils import formatdate

import pandas as pd
import streamlit as st
from openpyxl import load_workbook


# ============================================================
# CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="NOP Automation",
    page_icon="⚙️",
    layout="wide"
)

st.title("NOP AUTOMATION")
st.caption("Extra Cost generation and Gmail draft creation")

TEMPLATE_ROW = 3

# Colonnes nécessaires pour le filtre et les Extra Costs.
EXTRA_COLUMNS = [
    "OTP",
    "PROJET",
    "DATE ETOILE",
    "LIEU",
    "ARTICLE",
    "DESIGNATION",
    "QTE CN52N",
    "SOURCING",
]

# Colonnes supplémentaires nécessaires pour les brouillons.
MAIL_COLUMNS = [
    "DEFINITION STATUT",
    "LEADTIME",
    "QTE COMMANDE",
    "NO DE COMMANDE",
    "CONTACT SOURCING",
]

REQUIRED_COLUMNS = EXTRA_COLUMNS + MAIL_COLUMNS

# Correspondance avec le modèle Extra Cost.
# A = GFE : volontairement vide.
CELL_MAPPING = {
    "B": "ARTICLE",
    "C": "DESIGNATION",
    "D": "PROJET",
    "E": "QTE CN52N",
    "F": "DATE ETOILE",
    "G": "LIEU",
    "J": "OTP",
}


# ============================================================
# OUTILS DE NETTOYAGE
# ============================================================

def is_empty(value):
    if value is None:
        return True
    try:
        return bool(pd.isna(value))
    except (TypeError, ValueError):
        return False


def normalize_cell(value):
    if is_empty(value):
        return None
    if isinstance(value, str):
        value = value.strip()
        return value if value else None
    return value


def clean_text(value):
    if is_empty(value):
        return ""
    value = str(value).strip()
    if value.endswith(".0"):
        value = value[:-2]
    return value


def normalize_text(value):
    return re.sub(
        r"[^A-Za-z0-9]",
        "",
        clean_text(value)
    ).lower()


def filename_part(value):
    value = clean_text(value) or "SANS_VALEUR"
    value = re.sub(r'[\\/:*?"<>|]+', "_", value)
    value = re.sub(r"\s+", " ", value)
    return value.strip(" ._") or "SANS_VALEUR"


# ============================================================
# IDENTIFICATION DE LA NATURE
# ============================================================

def detect_nature(article):
    article = clean_text(article).upper()

    if article.endswith("XXXXS"):
        return "STANDARDS"

    if article.endswith("VXXQC"):
        return "VES"

    if article.endswith("GXXXC"):
        return "GEOM"

    return None


# ============================================================
# FILTRAGE DES LIGNES
# ============================================================

def filter_eligible_rows(source):
    source = source.copy()

    source.columns = [
        str(column).strip()
        for column in source.columns
    ]

    missing = [
        col for col in REQUIRED_COLUMNS
        if col not in source.columns
    ]

    if missing:
        raise ValueError(
            "Colonnes manquantes dans le fichier source : "
            + ", ".join(missing)
        )

    # Normalisation des valeurs utilisées dans le filtre.
    status = (
        source["DEFINITION STATUT"]
        .apply(clean_text)
        .str.strip()
    )

    leadtime = (
        source["LEADTIME"]
        .apply(clean_text)
        .str.upper()
        .str.strip()
    )

    # Règle NOP : statut 2 ET leadtime W/O CA.
    eligible = source[
        (status == "2") &
        (leadtime == "W/O CA")
    ].copy()

    # Nettoyage des champs utiles.
    for col in REQUIRED_COLUMNS:
        eligible[col] = eligible[col].apply(normalize_cell)

    return eligible


# ============================================================
# COPIE DU STYLE DU MODELE EXCEL
# ============================================================

def copy_cell_style(source, target):
    if source.has_style:
        target._style = copy(source._style)

    if source.number_format:
        target.number_format = source.number_format

    if source.alignment:
        target.alignment = copy(source.alignment)

    if source.protection:
        target.protection = copy(source.protection)


def fill_template(template_bytes, rows):
    workbook = load_workbook(
        io.BytesIO(template_bytes)
    )

    worksheet = workbook[workbook.sheetnames[0]]

    if worksheet.max_row < TEMPLATE_ROW:
        raise ValueError(
            "Le modèle Extra Cost ne contient pas la ligne 3."
        )

    max_col = max(worksheet.max_column, 10)

    # Mémoriser le style de la ligne 3.
    style_reference = {}

    for col in range(1, max_col + 1):
        style_reference[col] = copy(
            worksheet.cell(TEMPLATE_ROW, col)
        )

    # Supprimer les anciennes lignes de données sous la ligne 3.
    if worksheet.max_row > TEMPLATE_ROW:
        worksheet.delete_rows(
            TEMPLATE_ROW + 1,
            worksheet.max_row - TEMPLATE_ROW
        )

    # Remplir une ligne par article.
    for index, (_, record) in enumerate(rows.iterrows()):

        target_row = TEMPLATE_ROW + index

        if target_row > worksheet.max_row:
            worksheet.insert_rows(target_row)

        # Reproduire le style du modèle.
        for col in range(1, max_col + 1):
            source_style = style_reference[col]
            target_cell = worksheet.cell(target_row, col)
            copy_cell_style(source_style, target_cell)

        # GFE ignoré : colonne A vide.
        worksheet.cell(target_row, 1).value = None

        # Remplissage des cellules prévues.
        for column_letter, source_column in CELL_MAPPING.items():
            worksheet[f"{column_letter}{target_row}"] = (
                normalize_cell(record[source_column])
            )

    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)

    return output.getvalue()


# ============================================================
# GENERATION DES EXTRA COST
# ============================================================

def generate_extra_cost_files(eligible, template_bytes):

    data = eligible.copy()

    data["NATURE"] = data["ARTICLE"].apply(detect_nature)

    unknown_articles = data[
        data["NATURE"].isna()
    ].copy()

    recognized = data[
        data["NATURE"].notna()
    ].copy()

    grouping_columns = [
        "OTP",
        "PROJET",
        "SOURCING",
        "NATURE",
    ]

    generated_files = []

    for (otp, project, sourcing, nature), group in recognized.groupby(
        grouping_columns,
        dropna=False,
        sort=False
    ):

        suffix = ""

        if nature == "VES":
            suffix = " - VES"

        elif nature == "GEOM":
            suffix = " - GEOM"

        filename = (
            "EXTRA COST - "
            f"{filename_part(otp)} - "
            f"{filename_part(project)} - "
            f"{filename_part(sourcing)}"
            f"{suffix}.xlsx"
        )

        content = fill_template(
            template_bytes,
            group
        )

        generated_files.append({
            "filename": filename,
            "content": content,
            "otp": clean_text(otp),
            "project": clean_text(project),
            "sourcing": clean_text(sourcing),
            "nature": nature,
            "rows": len(group),
        })

    return generated_files, unknown_articles


# ============================================================
# RECHERCHE DES DOCUMENTS PO ET CMD
# ============================================================

def find_po_file(po_number, uploaded_documents):

    po_clean = normalize_text(po_number)

    if not po_clean:
        return None

    for document in uploaded_documents:

        name = document.name.lower()
        name_clean = normalize_text(document.name)

        if po_clean not in name_clean:
            continue

        # Ne pas confondre le PO et la CMD.
        if "cmd" in name or "commande" in name:
            continue

        return document

    return None


def find_cmd_file(po_number, uploaded_documents):

    po_clean = normalize_text(po_number)

    if not po_clean:
        return None

    for document in uploaded_documents:

        name = document.name.lower()
        name_clean = normalize_text(document.name)

        if po_clean not in name_clean:
            continue

        if "cmd" in name or "commande" in name:
            return document

    return None


# ============================================================
# CONSTRUCTION DU MAIL
# ============================================================

TABLE_COLUMNS = [
    "OTP",
    "PROJET",
    "DATE ETOILE",
    "LIEU",
    "ARTICLE",
    "DESIGNATION",
    "QTE COMMANDE",
    "NO DE COMMANDE",
]


def build_email(
    sender,
    contact,
    po_number,
    otp,
    project,
    sourcing,
    group
):

    message = EmailMessage()

    message["From"] = sender
    message["To"] = contact

    message["Subject"] = (
        f"CMD FULL NOP PO N°{po_number} - "
        f"{project} - {otp} - {sourcing}"
    )

    # Version texte.
    text_body = f"""Hello All,

As part of the NOP process for the {project}, please find attached Order No. {po_number} for VC {otp}, based on the prices agreed upon during your nomination.

Could you please confirm the following:

- Compliance with the Supplier Code of Conduct;
- Compliance with the required part number and index;
- Start of Production (SOP) date and expected delivery date.
- Please find the order attached with the extra cost if the price is not ok.

Thanks in advance.
"""

    message.set_content(text_body)

    # Tableau HTML.
    headers = "".join(
        f'<th style="border:1px solid black;'
        f'padding:7px;background:#4472C4;color:white;">'
        f'{html.escape(col)}</th>'
        for col in TABLE_COLUMNS
    )

    table_rows = ""

    for _, row in group.iterrows():

        cells = ""

        for col in TABLE_COLUMNS:
            value = html.escape(clean_text(row[col]))

            cells += (
                '<td style="border:1px solid black;'
                f'padding:6px;">{value}</td>'
            )

        table_rows += f"<tr>{cells}</tr>"

    project_html = html.escape(project)
    po_html = html.escape(po_number)
    otp_html = html.escape(otp)

    html_body = f"""
    <html>
    <body style="font-family:Arial,sans-serif;font-size:13px;">

    <p>Hello All,</p>

    <p>
    As part of the NOP process for
    <strong><span style="background-color:yellow;">
    {project_html}
    </span></strong>,
    please find attached
    <strong><span style="background-color:yellow;">
    Order No. {po_html}
    </span></strong>
    for VC
    <strong><span style="background-color:yellow;">
    {otp_html}
    </span></strong>,
    based on the prices agreed upon during your nomination.
    </p>

    <p>Could you please confirm the following:</p>

    <p>
    <strong>
    - Compliance with the Supplier Code of Conduct;<br>
    - Compliance with the required part number and index;<br>
    - Start of Production (SOP) date and expected delivery date.<br>
    - Please find the order attached with the extra cost if the price is not ok.
    </strong>
    </p>

    <table style="border-collapse:collapse;width:100%;
    font-family:Arial,sans-serif;font-size:11px;">
    <thead><tr>{headers}</tr></thead>
    <tbody>{table_rows}</tbody>
    </table>

    <br>
    <p>Thanks in advance.</p>

    </body>
    </html>
    """

    message.add_alternative(
        html_body,
        subtype="html"
    )

    return message


# ============================================================
# AJOUTER UNE PIECE JOINTE
# ============================================================

def attach_uploaded_file(message, uploaded_file):

    data = uploaded_file.getvalue()
    filename = uploaded_file.name

    mime_type, _ = mimetypes.guess_type(filename)

    if mime_type:
        maintype, subtype = mime_type.split("/", 1)
    else:
        maintype, subtype = "application", "octet-stream"

    message.add_attachment(
        data,
        maintype=maintype,
        subtype=subtype,
        filename=filename
    )


def attach_extra_cost(message, extra_file):

    filename = extra_file["filename"]
    content = extra_file["content"]

    message.add_attachment(
        content,
        maintype="application",
        subtype=(
            "vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        filename=filename
    )


# ============================================================
# CREATION DU BROUILLON GMAIL - IMAP
# ============================================================

def create_gmail_drafts(
    sender,
    app_password,
    eligible,
    generated_files,
    uploaded_documents
):

    mail = None
    created = 0
    errors = []
    warnings = []

    try:

        mail = imaplib.IMAP4_SSL("imap.gmail.com")

        mail.login(
            sender,
            app_password
        )

        # Dossier Brouillons Gmail.
        status, folders = mail.list()

        folder_names = [
            folder.decode(errors="ignore")
            for folder in (folders or [])
        ]

        draft_folder = None

        for folder in folder_names:

            if "[Gmail]/Drafts" in folder:
                draft_folder = '"[Gmail]/Drafts"'
                break

            if "[Gmail]/Brouillons" in folder:
                draft_folder = '"[Gmail]/Brouillons"'
                break

        if not draft_folder:
            raise RuntimeError(
                "Dossier Gmail Brouillons introuvable. "
                "Vérifie la langue et les dossiers de ton compte Gmail."
            )

        # Ne garder que les lignes avec un numéro de commande.
        mail_data = eligible.copy()

        mail_data["NO DE COMMANDE"] = (
            mail_data["NO DE COMMANDE"].apply(clean_text)
        )

        mail_data = mail_data[
            mail_data["NO DE COMMANDE"] != ""
        ].copy()

        group_columns = [
            "NO DE COMMANDE",
            "OTP",
            "PROJET",
            "SOURCING",
            "CONTACT SOURCING",
        ]

        groups = mail_data.groupby(
            group_columns,
            dropna=False,
            sort=False
        )

        for group_key, group in groups:

            po_number = clean_text(group_key[0])
            otp = clean_text(group_key[1])
            project = clean_text(group_key[2])
            sourcing = clean_text(group_key[3])
            contact = clean_text(group_key[4])

            if not contact:
                errors.append(
                    f"Commande {po_number} : CONTACT SOURCING vide."
                )
                continue

            po_file = find_po_file(
                po_number,
                uploaded_documents
            )

            cmd_file = find_cmd_file(
                po_number,
                uploaded_documents
            )

            # Extra Costs correspondant à OTP + projet + sourcing.
            matching_extra_costs = []

            for item in generated_files:

                if (
                    normalize_text(item["otp"]) == normalize_text(otp)
                    and normalize_text(item["project"]) == normalize_text(project)
                    and normalize_text(item["sourcing"]) == normalize_text(sourcing)
                ):
                    matching_extra_costs.append(item)

            message = build_email(
                sender,
                contact,
                po_number,
                otp,
                project,
                sourcing,
                group
            )

            # Pièce jointe PO.
            if po_file:
                attach_uploaded_file(message, po_file)
            else:
                warnings.append(
                    f"Commande {po_number} : PO non trouvé."
                )

            # Pièce jointe CMD.
            if cmd_file:
                attach_uploaded_file(message, cmd_file)
            else:
                warnings.append(
                    f"Commande {po_number} : CMD non trouvée."
                )

            # Tous les Extra Costs correspondants.
            for extra_file in matching_extra_costs:
                attach_extra_cost(message, extra_file)

            if not matching_extra_costs:
                warnings.append(
                    f"Commande {po_number} : Extra Cost correspondant non trouvé."
                )

            # Création du brouillon uniquement.
            # Aucun mail.send() n'est utilisé.
            try:

                result = mail.append(
                    draft_folder,
                    "\\Draft",
                    None,
                    message.as_bytes()
                )

                if result[0] == "OK":
                    created += 1
                else:
                    errors.append(
                        f"Commande {po_number} : "
                        f"échec de création du brouillon."
                    )

            except Exception as exc:
                errors.append(
                    f"Commande {po_number} : {str(exc)}"
                )

    finally:

        if mail is not None:
            try:
                mail.logout()
            except Exception:
                pass

    return created, errors, warnings


# ============================================================
# INTERFACE STREAMLIT
# ============================================================

st.sidebar.title("NOP Automation")

page = st.sidebar.radio(
    "Navigation",
    [
        "Accueil",
        "Traitement NOP",
        "Informations"
    ]
)


if page == "Accueil":

    st.header("Bienvenue dans NOP Automation")

    st.write(
        "Cette application automatise le traitement des lignes "
        "NOP, la génération des Extra Costs et la création "
        "des brouillons Gmail."
    )

    st.info(
        "Règle de sélection : DEFINITION STATUT = 2 "
        "ET LEADTIME = W/O CA."
    )

    st.markdown(
        """
        **Étapes du traitement :**

        1. Importer le fichier source Excel.
        2. Importer le modèle Extra Cost.
        3. Importer les documents PO et CMD.
        4. Générer les fichiers Extra Cost.
        5. Créer les brouillons Gmail.
        """
    )


elif page == "Informations":

    st.header("Règles de traitement")

    st.markdown(
        """
        | Condition | Traitement |
        |---|---|
        | DEFINITION STATUT = 2 | Obligatoire |
        | LEADTIME = W/O CA | Obligatoire |
        | ARTICLE se termine par XXXXS | STANDARDS |
        | ARTICLE se termine par VXXQC | VES |
        | ARTICLE se termine par GXXXC | GEOM |

        **Nommage des fichiers :**

        - STANDARDS : `EXTRA COST - OTP - PROJET - SOURCING.xlsx`
        - VES : `EXTRA COST - OTP - PROJET - SOURCING - VES.xlsx`
        - GEOM : `EXTRA COST - OTP - PROJET - SOURCING - GEOM.xlsx`

        La cellule GFE reste vide.
        Les quantités vides restent vides.
        Aucun e-mail n'est envoyé automatiquement.
        """
    )


elif page == "Traitement NOP":

    st.header("Traitement automatique NOP")

    st.subheader("1. Importer les fichiers")

    source_upload = st.file_uploader(
        "Fichier source Excel (.xlsx ou .xlsm)",
        type=["xlsx", "xlsm"],
        key="source"
    )

    template_upload = st.file_uploader(
        "Modèle Extra Cost (.xlsx)",
        type=["xlsx"],
        key="template"
    )

    documents_upload = st.file_uploader(
        "Documents PO et CMD (PDF, Excel, Word)",
        type=[
            "pdf",
            "xlsx",
            "xls",
            "xlsm",
            "doc",
            "docx"
        ],
        accept_multiple_files=True,
        key="documents"
    )

    st.divider()

    if source_upload and template_upload:

        if st.button(
            "1. Analyser le fichier et générer les Extra Costs",
            type="primary"
        ):

            try:

                with st.spinner("Analyse du fichier source..."):

                    source = pd.read_excel(
                        io.BytesIO(source_upload.getvalue()),
                        engine="openpyxl"
                    )

                    source.columns = [
                        str(col).strip()
                        for col in source.columns
                    ]

                    eligible = filter_eligible_rows(source)

                    generated_files, unknown_articles = (
                        generate_extra_cost_files(
                            eligible,
                            template_upload.getvalue()
                        )
                    )

                    # Sauvegarder les résultats dans la session.
                    st.session_state["eligible"] = eligible
                    st.session_state["generated_files"] = generated_files
                    st.session_state["unknown_articles"] = unknown_articles
                    st.session_state["source_analyzed"] = True

                st.success("Analyse terminée.")

                col1, col2, col3 = st.columns(3)

                col1.metric(
                    "Lignes source",
                    len(source)
                )

                col2.metric(
                    "Lignes éligibles",
                    len(eligible)
                )

                col3.metric(
                    "Extra Costs générés",
                    len(generated_files)
                )

                st.subheader("Résultat du filtrage")

                st.write(
                    "Lignes retenues : STATUT 2 et LEADTIME W/O CA."
                )

                st.dataframe(
                    eligible,
                    use_container_width=True
                )

                if generated_files:

                    st.subheader("Fichiers Extra Cost générés")

                    zip_buffer = io.BytesIO()

                    with zipfile.ZipFile(
                        zip_buffer,
                        "w",
                        zipfile.ZIP_DEFLATED
                    ) as archive:

                        for item in generated_files:

                            archive.writestr(
                                item["filename"],
                                item["content"]
                            )

                            st.write(
                                f"- **{item['filename']}** "
                                f"({item['rows']} article(s))"
                            )

                    st.download_button(
                        "Télécharger tous les Extra Costs (.zip)",
                        data=zip_buffer.getvalue(),
                        file_name="EXTRA_COST_GENERES.zip",
                        mime="application/zip"
                    )

                else:

                    st.warning(
                        "Aucun Extra Cost généré. "
                        "Vérifie les terminaisons de la colonne ARTICLE."
                    )

                if not unknown_articles.empty:

                    st.warning(
                        "Certaines références ARTICLE ne sont pas reconnues."
                    )

                    st.dataframe(
                        unknown_articles[
                            ["OTP", "PROJET", "ARTICLE", "SOURCING"]
                        ],
                        use_container_width=True
                    )

            except Exception as exc:

                st.error(
                    f"Erreur pendant le traitement : {str(exc)}"
                )

    else:

        st.info(
            "Importe le fichier source et le modèle Extra Cost "
            "pour commencer."
        )

    # --------------------------------------------------------
    # CREATION DES BROUILLONS
    # --------------------------------------------------------

    if st.session_state.get("source_analyzed", False):

        st.divider()
        st.subheader("2. Création des brouillons Gmail")

        st.warning(
            "Les identifiants Gmail sont utilisés uniquement "
            "pour cette session. Ne les enregistre pas dans GitHub."
        )

        sender = st.text_input(
            "Adresse Gmail professionnelle",
            placeholder="nom@entreprise.com"
        )

        app_password = st.text_input(
            "Mot de passe d'application Gmail",
            type="password",
            help=(
                "Utilise un mot de passe d'application autorisé "
                "par ton organisation. Ne partage jamais ce mot de passe."
            )
        )

        if st.button(
            "2. Créer les brouillons Gmail",
            type="primary"
        ):

            if not sender or not app_password:
                st.error(
                    "Renseigne l'adresse Gmail et le mot de passe d'application."
                )

            else:

                try:

                    with st.spinner(
                        "Connexion Gmail et création des brouillons..."
                    ):

                        created, errors, warnings = create_gmail_drafts(
                            sender,
                            app_password,
                            st.session_state["eligible"],
                            st.session_state["generated_files"],
                            documents_upload or []
                        )

                    st.success(
                        f"{created} brouillon(s) créé(s) dans Gmail."
                    )

                    st.info(
                        "Aucun e-mail n'a été envoyé. "
                        "Vérifie Gmail > Brouillons."
                    )

                    if warnings:

                        st.warning("Documents manquants ou avertissements :")

                        for warning in warnings:
                            st.write("- " + warning)

                    if errors:

                        st.error("Erreurs de traitement :")

                        for error in errors:
                            st.write("- " + error)

                except Exception as exc:

                    st.error(
                        f"Erreur Gmail : {str(exc)}"
                    )

    else:

        st.caption(
            "Génère d'abord les Extra Costs pour activer "
            "la création des brouillons Gmail."
        )
