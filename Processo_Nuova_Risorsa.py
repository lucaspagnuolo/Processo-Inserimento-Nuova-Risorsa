import streamlit as st
import csv
import pandas as pd
from datetime import datetime, timedelta
import io
import unicodedata
import zipfile

# ------------------------------------------------------------
# Caricamento configurazione da Excel caricato dall'utente
# ------------------------------------------------------------
def load_config_from_bytes(data: bytes):
    cfg = pd.read_excel(io.BytesIO(data), sheet_name=None, engine="openpyxl")

    # Estrai configurazione Risorsa Interna se presente
    if "Risorsa Interna" in cfg:
        ris = cfg["Risorsa Interna"]
    else:
        ris = pd.DataFrame()

    ou_df = ris[ris["Section"] == "OU"][["Key/App", "Label/Gruppi/Value"]].rename(
        columns={"Key/App": "key", "Label/Gruppi/Value": "label"})
    ou_options = dict(zip(ou_df["key"], ou_df["label"]))

    grp_df = ris[ris["Section"] == "InserimentoGruppi"][["Key/App", "Label/Gruppi/Value"]].rename(
        columns={"Key/App": "app", "Label/Gruppi/Value": "gruppi"})
    gruppi = dict(zip(grp_df["app"], grp_df["gruppi"]))

    def_df = ris[ris["Section"] == "Defaults"][["Key/App", "Label/Gruppi/Value"]].rename(
        columns={"Key/App": "key", "Label/Gruppi/Value": "value"})
    defaults = dict(zip(def_df["key"], def_df["value"]))

    # Estrai organigramma se presente
    organigramma = {}
    if "organigramma" in cfg:
        org = cfg["organigramma"].iloc[:, :2].dropna(how="all")
        org.columns = ["label", "value"]
        organigramma = dict(zip(org["label"], org["value"]))

    return ou_options, gruppi, defaults, organigramma

# ------------------------------------------------------------
# App 1.1: Nuova Risorsa Interna
# ------------------------------------------------------------
st.set_page_config(page_title="1.1 Nuova Risorsa Interna")
st.title("1.1 Nuova Risorsa Interna")

config_file = st.file_uploader(
    "Carica il file di configurazione (config.xlsx)", type=["xlsx"],
    help="Deve contenere il foglio 'Risorsa Interna' con campo Section"
)
if not config_file:
    st.warning("Per favore carica il file di configurazione per continuare.")
    st.stop()

ou_options, gruppi, defaults, organigramma = load_config_from_bytes(config_file.read())

# --- Estrazione valori defaults e gruppi O365 ---
dl_standard = defaults.get("dl_standard", "").split(";") if defaults.get("dl_standard") else []
dl_vip = defaults.get("dl_vip", "").split(";") if defaults.get("dl_vip") else []

o365_groups = []
for k, v in defaults.items():
    if str(k).startswith("grp_o365_") and v:
        parts = [p.strip() for p in str(v).split(";") if p.strip()]
        for p in parts:
            token = p
            if token.startswith("365 "):
                token = "O" + token
            o365_groups.append(token)

grp_foorban = defaults.get("grp_foorban", "")
grp_salesforce = defaults.get("grp_salesforce", "")
pillole = defaults.get("pillole", "")

# Utility functions
def auto_quote(fields, quotechar='"', predicate=lambda s: ' ' in s):
    out = []
    for f in fields:
        s = str(f)
        if predicate(s):
            out.append(f'{quotechar}{s}{quotechar}')
        else:
            out.append(s)
    return out

def normalize_name(s: str) -> str:
    nfkd = unicodedata.normalize('NFKD', s)
    ascii_str = nfkd.encode('ASCII', 'ignore').decode()
    return ascii_str.replace(' ', '').replace("'", '').lower()

def genera_samaccountname(nome, cognome, secondo_nome="", secondo_cognome="", esterno=False):
    n = normalize_name(nome)
    sn = normalize_name(secondo_nome)
    c = normalize_name(cognome)
    sc = normalize_name(secondo_cognome)
    suffix = ".ext" if esterno else ""
    limit = 16 if esterno else 20
    cand1 = f"{n}{sn}.{c}{sc}"
    if len(cand1) <= limit:
        return cand1 + suffix
    cand2 = f"{n[:1]}{sn[:1]}.{c}{sc}"
    if len(cand2) <= limit:
        return cand2 + suffix
    base = f"{n[:1]}{sn[:1]}.{c}"
    return base[:limit] + suffix

def build_full_name(cognome, secondo_cognome, nome, secondo_nome, esterno=False):
    parts = [p for p in [cognome, secondo_cognome, nome, secondo_nome] if p]
    return " ".join(parts) + (" (esterno)" if esterno else "")

HEADER_UTENTE = [
    "sAMAccountName","Creation","OU","Name","DisplayName","cn","GivenName",
    "Surname","employeeNumber","employeeID","department","Description",
    "passwordNeverExpired","ExpireDate","userprincipalname","mail","mobile",
    "RimozioneGruppo","InserimentoGruppo","disable","moveToOU","telephoneNumber","company"
]
HEADER_COMPUTER = [
    "Computer","OU","add_mail","remove_mail","add_mobile","remove_mobile",
    "add_userprincipalname","remove_userprincipalname","disable","moveToOU"
]

# Input Module
st.subheader("Modulo Inserimento Nuova Risorsa Interna")
employee_id      = st.text_input("Matricola", defaults.get("employee_id_default", "")).strip()
cognome          = st.text_input("Cognome").strip().capitalize()
secondo_cognome  = st.text_input("Secondo Cognome").strip().capitalize()
nome             = st.text_input("Nome").strip().capitalize()
secondo_nome     = st.text_input("Secondo Nome").strip().capitalize()
codice_fiscale   = st.text_input("Codice Fiscale", "").strip()

# Dropdown per Sigla Divisione-Area
if organigramma:
    dept_label = st.selectbox("Sigla Divisione-Area", ["-- Seleziona --"] + list(organigramma.keys()))
    department = organigramma.get(dept_label, "") if dept_label != "-- Seleziona --" else defaults.get("department_default", "")
else:
    department = st.text_input("Sigla Divisione-Area", defaults.get("department_default", "")).strip()

numero_telefono  = st.text_input("Mobile (+39 già inserito)", "").replace(" ", "")
description      = st.text_input("PC (lascia vuoto per <PC>)", "<PC>").strip()
resident_flag    = st.checkbox("È Resident?")
numero_fisso     = st.text_input("Numero fisso Resident (+39 già inserito)", "").strip() if resident_flag else ""
telephone_number = f"+39 {numero_fisso}" if resident_flag and numero_fisso else defaults.get("telephone_interna", "")

ou_vals = list(ou_options.values()) if ou_options else []
def_o = defaults.get("ou_default", ou_vals[0] if ou_vals else "")
index_default = 0
if def_o in ou_vals:
    index_default = ou_vals.index(def_o)

label_ou = st.selectbox("Tipologia Utente", ou_vals, index=index_default) if ou_vals else st.text_input("Tipologia Utente", "")
selected_key = list(ou_options.keys())[ou_vals.index(label_ou)] if ou_vals else ""
ou_value = ou_options[selected_key] if ou_vals else ""
inserimento_gruppo = ""
company = defaults.get("company_interna", "")

data_operativa = st.text_input("Data operatività (gg/mm/aaaa)", "").strip()
profilazione = st.checkbox("Profilazione SM?")
sm_lines = st.text_area("SM (una per riga)", "").splitlines() if profilazione else []

dl_list = dl_standard if selected_key == "utenti_standard" else dl_vip if selected_key == "utenti_vip" else []

# Generazione CSV Utente + Computer + Profilazione
if st.button("Genera CSV"):
    sAM = genera_samaccountname(nome, cognome, secondo_nome, secondo_cognome, False)
    cn = build_full_name(cognome, secondo_cognome, nome, secondo_nome, False)
    name_parts = [cognome] + ([secondo_cognome] if secondo_cognome else []) + [nome[:1] if nome else ""]
    basename = "_".join([p for p in name_parts if p])
    given = f"{nome} {secondo_nome}".strip()
    surn = f"{cognome} {secondo_cognome}".strip()
    mobile = f"+39 {numero_telefono}" if numero_telefono else ""

    row_ut = [
        sAM, "SI", ou_value, cn, cn, cn, given, surn,
        codice_fiscale, employee_id, department, description or "<PC>",
        "No", "", f"{sAM}@consip.it", f"{sAM}@consip.it", mobile,
        "", "", "", "", telephone_number, company
    ]

    row_cp = [
        description or "", "", f"{sAM}@consip.it", "", mobile, "", cn, "", "", ""
    ]

    existing_o365 = list(o365_groups)
    inserimento_gruppo_val = gruppi.get("interna", "") or ""
    inser_list = []
    if inserimento_gruppo_val:
        for g in str(inserimento_gruppo_val).split(";"):
            gg = g.strip()
            if not gg:
                continue
            if gg.startswith("365 "):
                gg = "O" + gg
            inser_list.append(gg)

    merged_profilazione = []
    for g in existing_o365 + inser_list:
        if g and g not in merged_profilazione:
            merged_profilazione.append(g)

    gruppi_profilazione_str = ";".join(merged_profilazione)

    row_profilazione = [""] * len(HEADER_UTENTE)
    row_profilazione[0] = sAM
    try:
        idx_inserimento = HEADER_UTENTE.index("InserimentoGruppo")
        row_profilazione[idx_inserimento] = gruppi_profilazione_str
    except ValueError:
        row_profilazione.append(gruppi_profilazione_str)

    msg_utente = (
        "Salve.\n"
        "Vi richiediamo la definizione della utenza nell’AD Consip come dettagliato nei file:\n"
        f"//srv_dati/AreaCondivisa/DEPSI/IC/Utenze/Interni/{basename}_utente.csv\n"
        "Restiamo in attesa di un vostro riscontro ad attività completata.\n"
        "Saluti"
    )
    msg_computer = (
        "Salve.\n"
        "Si richiede modifiche come da file:\n"
        f"//srv_dati/AreaCondivisa/DEPSI/IC/PC/{basename}_computer.csv\n"
        "Restiamo in attesa di un vostro riscontro ad attività completata.\n"
        "Saluti"
    )
    msg_profilazione = (
        "Salve.\n"
        "Si richiede modifiche come da file:\n"
        f"//srv_dati/AreaCondivisa/DEPSI/IC/Profilazione/{basename}_profilazione.csv\n"
        "Restiamo in attesa di un vostro riscontro ad attività completata.\n"
        "Saluti"
    )

    # --- Preparazione CSV ---
    def prepara_csv(row, header):
        buf = io.StringIO()
        w = csv.writer(buf, quoting=csv.QUOTE_NONE, escapechar="\\")
        quoted = auto_quote(row, quotechar='"', predicate=lambda s: ' ' in s)
        w.writerow(header)
        w.writerow(quoted)
        buf.seek(0)
        return buf

    buf_user = prepara_csv(row_ut, HEADER_UTENTE)
    buf_comp = prepara_csv(row_cp, HEADER_COMPUTER)
    buf_prof = prepara_csv(row_profilazione, HEADER_UTENTE)

    # --- Template HTML per la posta elettronica ---
    table_rows_html = f"""
        <tr><td><strong>Tipo Utenza</strong></td><td>Remota</td></tr>
        <tr><td><strong>Utenza</strong></td><td>{sAM}</td></tr>
        <tr><td><strong>Alias</strong></td><td>{sAM}</td></tr>
        <tr><td><strong>Display name</strong></td><td>{cn}</td></tr>
        <tr><td><strong>Common name</strong></td><td>{cn}</td></tr>
        <tr><td><strong>e-mail</strong></td><td>{sAM}@consip.it</td></tr>
        <tr><td><strong>e-mail secondaria</strong></td><td>{sAM}@consipspa.mail.onmicrosoft.com</td></tr>
    """

    dl_html = ""
    if dl_list:
        dl_html = "<h4>DL da aggiungere il giorno {}</h4><ul>{}</ul>".format(
            data_operativa or "",
            "".join(f"<li>{dl}</li>" for dl in dl_list)
        )

    sm_html = ""
    if profilazione and sm_lines:
        sm_html = "<h4>Profilare su SM</h4><ul>{}</ul>".format("".join(f"<li>{s}</li>" for s in sm_lines))

    azure_items = []
    if grp_foorban:
        azure_items.append(grp_foorban)
    if grp_salesforce:
        azure_items.append(grp_salesforce)
    if pillole:
        azure_items.append(f"canale {pillole}")
    azure_html = ""
    if azure_items:
        azure_html = "<h4>Aggiungere utenza al gruppo Azure</h4><ul>{}</ul>".format("".join(f"<li>{a}</li>" for a in azure_items))

    template_preview_html = f"""
    <!doctype html>
    <html lang="it">
    <head>
      <meta charset="utf-8">
      <title>Anteprima Template - {basename}</title>
      <style>
        body {{ font-family: Arial, Helvetica, sans-serif; font-size:14px; }}
        table {{ border-collapse: collapse; width: 100%; max-width: 800px; }}
        td, th {{ border: 1px solid #ddd; padding: 8px; vertical-align: top; }}
        th {{ background-color: #f4f4f4; text-align: left; }}
        h2, h4 {{ margin: 12px 0 6px; }}
      </style>
    </head>
    <body>
      <h2>Richiesta definizione casella - anteprima template</h2>
      <table>
        <thead><tr><th>Campo</th><th>Valore</th></tr></thead>
        <tbody>
          {table_rows_html}
        </tbody>
      </table>
      {dl_html}
      {sm_html}
      {azure_html}
      <p>Grazie<br/>Saluti</p>
    </body>
    </html>
    """

    # File .eml (client di posta)
    eml_headers = [
        f"Subject: Richiesta definizione casella - {sAM}",
        f"From: your.name@consip.it",
        f"To: destinatario@consip.it",
        "MIME-Version: 1.0",
        'Content-Type: multipart/alternative; boundary="BOUNDARY"',
        "",
        "--BOUNDARY",
        "Content-Type: text/html; charset=utf-8",
        "",
    ]
    eml_body = template_preview_html
    eml_footer = ["", "--BOUNDARY--", ""]
    template_preview_eml = "\r\n".join(eml_headers + [eml_body] + eml_footer)

    # --- ZIP unico ---
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
        zipf.writestr(f"{basename}_utente.csv", buf_user.getvalue())
        zipf.writestr(f"{basename}_computer.csv", buf_comp.getvalue())
        zipf.writestr(f"{basename}_profilazione.csv", buf_prof.getvalue())
        zipf.writestr(f"{basename}_template_preview.html", template_preview_html)
        zipf.writestr(f"{basename}_template_preview.eml", template_preview_eml)
        zipf.writestr(f"{basename}_msg_utente.txt", msg_utente)
        zipf.writestr(f"{basename}_msg_computer.txt", msg_computer)
        zipf.writestr(f"{basename}_msg_profilazione.txt", msg_profilazione)

    zip_buffer.seek(0)

    st.download_button(
        "📦 Scarica Tutti i CSV + Template (ZIP)",
        data=zip_buffer,
        file_name=f"{basename}_bundle.zip",
        mime="application/zip"
    )

    st.success(f"✅ CSV e template generati per '{sAM}'")
