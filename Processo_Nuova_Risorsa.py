import streamlit as st
import csv
import pandas as pd
from datetime import datetime, timedelta
import io
import unicodedata

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

# --- Estrazione valori defaults e gruppi O365 (gestione robusta di ; e più chiavi grp_o365_) ---
dl_standard = defaults.get("dl_standard", "").split(";") if defaults.get("dl_standard") else []
dl_vip = defaults.get("dl_vip", "").split(";") if defaults.get("dl_vip") else []

# raccogliamo tutti i valori di default con chiave che inizia per grp_o365_
o365_groups = []
for k, v in defaults.items():
    if str(k).startswith("grp_o365_") and v:
        parts = [p.strip() for p in str(v).split(";") if p.strip()]
        for p in parts:
            token = p
            # correzione automatica: se per errore manca la 'O' iniziale (es. "365 ...")
            if token.startswith("365 "):
                token = "O" + token
            o365_groups.append(token)

grp_foorban = defaults.get("grp_foorban", "")
grp_salesforce = defaults.get("grp_salesforce", "")  # <-- lettura grp_salesforce
pillole = defaults.get("pillole", "")

# Percorso di archivio (raw string per evitare escape warnings)
ARCHIVE_PATH = r"\\\\srv_dati.consip.tesoro.it\AreaCondivisa\DEPSI\IC\AD_Modifiche"

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

# Dropdown per Sigla Divisione-Area da organigramma
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
# inserimento_gruppo lasciato vuoto per il CSV utente (lo useremo solo per costruire il CSV profilazione)
inserimento_gruppo = ""
company = defaults.get("company_interna", "")

data_operativa = st.text_input("Data operatività (gg/mm/aaaa)", "").strip()
profilazione = st.checkbox("Profilazione SM?")
sm_lines = st.text_area("SM (una per riga)", "").splitlines() if profilazione else []

dl_list = dl_standard if selected_key == "utenti_standard" else dl_vip if selected_key == "utenti_vip" else []

# Preview Messaggio
if st.button("Template per Posta Elettronica"):
    sAM = genera_samaccountname(nome, cognome, secondo_nome, secondo_cognome, False)
    cn = build_full_name(cognome, secondo_cognome, nome, secondo_nome, False)
    table_md = (
        "| Campo             | Valore                                      |\n"
        "|-------------------|---------------------------------------------|\n"
        f"| Tipo Utenza       | Remota                                      |\n"
        f"| Utenza            | {sAM}                                       |\n"
        f"| Alias             | {sAM}                                       |\n"
        f"| Display name      | {cn}                                        |\n"
        f"| Common name       | {cn}                                        |\n"
        f"| e-mail            | {sAM}@consip.it                              |\n"
        f"| e-mail secondaria | {sAM}@consipspa.mail.onmicrosoft.com        |\n"
    )
    st.markdown("Ciao.  \nRichiedo cortesemente la definizione di una casella di posta come sottoindicato.")
    st.markdown(table_md)
    if dl_list:
        st.markdown(f"Il giorno **{data_operativa}** occorre inserire la casella nelle DL:")
        for dl in dl_list:
            st.markdown(f"- {dl}")
    if profilazione:
        st.markdown("Profilare su SM:")
        for sm in sm_lines:
            st.markdown(f"- {sm}")

    st.markdown("**Aggiungere utenza al gruppo Azure:**")
    if grp_foorban:
        st.markdown(f"- {grp_foorban}")
    if grp_salesforce:
        st.markdown(f"- {grp_salesforce}")
    if pillole:
        st.markdown(f"- canale {pillole}")

    st.markdown("Grazie  \nSaluti")

# Generazione CSV Utente + Computer + Profilazione (ex O365)
if st.button("Genera CSV"):
    sAM = genera_samaccountname(nome, cognome, secondo_nome, secondo_cognome, False)
    cn = build_full_name(cognome, secondo_cognome, nome, secondo_nome, False)
    norm_cognome = normalize_name(cognome)
    norm_secondo = normalize_name(secondo_cognome) if secondo_cognome else ''
    name_parts = [cognome] + ([secondo_cognome] if secondo_cognome else []) + [nome[:1] if nome else ""]
    basename = "_".join([p for p in name_parts if p])
    given = f"{nome} {secondo_nome}".strip()
    surn = f"{cognome} {secondo_cognome}".strip()
    mobile = f"+39 {numero_telefono}" if numero_telefono else ""

    # ---> costruisco row_ut assicurandomi che InserimentoGruppo sia vuoto
    row_ut = [
        sAM, "SI", ou_value, cn, cn, cn, given, surn,
        codice_fiscale, employee_id, department, description or "<PC>",
        "No", "", f"{sAM}@consip.it", f"{sAM}@consip.it", mobile,
        "", "", "", "", telephone_number, company
    ]

    row_cp = [
        description or "", "", f"{sAM}@consip.it", "", mobile, "", cn, "", "", ""
    ]

    # Costruisco la lista dei gruppi di profilazione: o365_groups + inserimento (key "interna")
    existing_o365 = list(o365_groups)
    inserimento_gruppo_val = gruppi.get("interna", "") or ""
    inser_gr_raw = inserimento_gruppo_val
    inser_list = []
    if inser_gr_raw:
        for g in str(inser_gr_raw).split(";"):
            gg = g.strip()
            if not gg:
                continue
            if gg.startswith("365 "):
                gg = "O" + gg
            inser_list.append(gg)

    # merged_profilazione evitando duplicati mantenendo ordine
    merged_profilazione = []
    for g in existing_o365 + inser_list:
        if g and g not in merged_profilazione:
            merged_profilazione.append(g)

    # join senza spazi dopo ';' come richiesto dall'utente
    gruppi_profilazione_str = ";".join(merged_profilazione)

    row_profilazione = [""] * len(HEADER_UTENTE)
    row_profilazione[0] = sAM
    try:
        idx_inserimento = HEADER_UTENTE.index("InserimentoGruppo")
        row_profilazione[idx_inserimento] = gruppi_profilazione_str
    except ValueError:
        # se per qualche motivo l'header non c'è, appendiamo comunque il valore
        row_profilazione.append(gruppi_profilazione_str)

    msg_utente = (
    "Salve.\n"
    "Vi richiediamo la definizione della utenza nell’AD Consip come dettagliato nei file:\n"
    fr"\\srv_dati\AreaCondivisa\DEPSI\IC\Utenze\Interni\{basename}_utente.csv \n"
    "Restiamo in attesa di un vostro riscontro ad attività completata.\n"
    "Saluti"
    )

    msg_computer = (
    "Salve.\n"
    "Si richiede modifiche come da file:\n"
    f"\\srv_dati\AreaCondivisa\DEPSI\IC\PC\{basename}_computer.csv\n"
    "Restiamo in attesa di un vostro riscontro ad attività completata.\n"
    "Saluti"
    )

    st.subheader(f"Nuova Utenza AD [{cognome}]")
    st.markdown(msg_utente)
    st.subheader("Anteprima CSV Utente")
    st.dataframe(pd.DataFrame([row_ut], columns=HEADER_UTENTE))

    st.subheader(f"Modifica AD Computer [{cognome}]")
    st.markdown(msg_computer)
    st.subheader("Anteprima CSV Computer")
    st.dataframe(pd.DataFrame([row_cp], columns=HEADER_COMPUTER))

    st.subheader("Anteprima CSV Profilazione")
    st.dataframe(pd.DataFrame([row_profilazione], columns=HEADER_UTENTE))

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

    st.download_button(
        "📥 Scarica CSV Utente",
        data=buf_user.getvalue(),
        file_name=f"{basename}_utente.csv",
        mime="text/csv"
    )
    st.download_button(
        "📥 Scarica CSV Computer",
        data=buf_comp.getvalue(),
        file_name=f"{basename}_computer.csv",
        mime="text/csv"
    )
    st.download_button(
        "📥 Scarica CSV Profilazione",
        data=buf_prof.getvalue(),
        file_name=f"{basename}_profilazione.csv",
        mime="text/csv"
    )
    st.success(f"✅ CSV generati per '{sAM}'")
