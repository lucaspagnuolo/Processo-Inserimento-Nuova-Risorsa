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

    ou_df = ris[ris["Section"] == "OU"][["Key/App","Label/Gruppi/Value"]].rename(
        columns={"Key/App": "key", "Label/Gruppi/Value": "label"})
    ou_options = dict(zip(ou_df["key"], ou_df["label"]))

    grp_df = ris[ris["Section"] == "InserimentoGruppi"][["Key/App","Label/Gruppi/Value"]].rename(
        columns={"Key/App": "app", "Label/Gruppi/Value": "gruppi"})
    gruppi = dict(zip(grp_df["app"], grp_df["gruppi"]))

    def_df = ris[ris["Section"] == "Defaults"][["Key/App","Label/Gruppi/Value"]].rename(
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

dl_standard = defaults.get("dl_standard", "").split(";") if defaults.get("dl_standard") else []
dl_vip = defaults.get("dl_vip", "").split(";") if defaults.get("dl_vip") else []
o365_groups = [v for k, v in defaults.items() if k.startswith("grp_o365_")]
grp_foorban = defaults.get("grp_foorban", "")
pillole = defaults.get("pillole", "")

# Utility functions
def normalize_name(s: str) -> str:
    """Rimuove spazi, apostrofi e accenti, restituisce in minuscolo."""
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

ou_vals           = list(ou_options.values())
def_o             = defaults.get("ou_default", ou_vals[0] if ou_vals else "")
label_ou          = st.selectbox("Tipologia Utente", ou_vals, index=ou_vals.index(def_o))
selected_key      = list(ou_options.keys())[ou_vals.index(label_ou)]
ou_value          = ou_options[selected_key]
inserimento_gruppo = gruppi.get("interna", "")
company           = defaults.get("company_interna", "")

data_operativa = st.text_input("Data operatività (gg/mm/aaaa)", "").strip()
profilazione    = st.checkbox("Profilazione SM?")
sm_lines        = st.text_area("SM (una per riga)", "").splitlines() if profilazione else []

dl_list = dl_standard if selected_key == "utenti_standard" else dl_vip if selected_key == "utenti_vip" else []

# Preview Messaggio (Lasciata invariata)
if st.button("Template per Posta Elettronica"):
    sAM = genera_samaccountname(nome, cognome, secondo_nome, secondo_cognome, False)
    cn = build_full_name(cognome, secondo_cognome, nome, secondo_nome, False)
    groups_md = "\n".join(f"- {g}" for g in o365_groups)
    table_md = f"""
| Campo             | Valore                                     |
|-------------------|--------------------------------------------|
| Tipo Utenza       | Remota                                     |
| Utenza            | {sAM}                                      |
| Alias             | {sAM}                                      |
| Display name      | {cn}                                       |
| Common name       | {cn}                                       |
| e-mail            | {sAM}@consip.it                            |
| e-mail secondaria | {sAM}@consipspa.mail.onmicrosoft.com       |
"""
    st.markdown("Ciao.  \nRichiedo cortesemente la definizione di una casella di posta come sottoindicato.")
    st.markdown(table_md)
    st.markdown(f"Inviare batch di notifica migrazione mail a: imac@consip.it  \nAggiungere utenza di dominio ai gruppi:\n{groups_md}")
    if dl_list:
        st.markdown(f"Il giorno **{data_operativa}** occorre inserire la casella nelle DL:")
        for dl in dl_list:
            st.markdown(f"- {dl}")
    if profilazione:
        st.markdown("Profilare su SM:")
        for sm in sm_lines:
            st.markdown(f"- {sm}")
    st.markdown(f"Aggiungere utenza al:\n- gruppo Azure: {grp_foorban}\n- canale {pillole}")
    st.markdown("Grazie  \nSaluti")

# Unica Generazione CSV Utente + Computer
if st.button("Genera CSV"):    
    sAM = genera_samaccountname(nome, cognome, secondo_nome, secondo_cognome, False)
    cn = build_full_name(cognome, secondo_cognome, nome, secondo_nome, False)
    norm_cognome = normalize_name(cognome)
    norm_secondo = normalize_name(secondo_cognome) if secondo_cognome else ''
    name_parts = [norm_cognome] + ([norm_secondo] if norm_secondo else []) + [nome[:1].lower()]
    basename = "_".join(name_parts)
    name_parts = [cognome] + ([secondo_cognome] if secondo_cognome else []) + [nome[:1]]
    basename = "_".join(name_parts)
    given = f"{nome} {secondo_nome}".strip()
    surn = f"{cognome} {secondo_cognome}".strip()
    mobile = f"+39 {numero_telefono}" if numero_telefono else ""

    # Riga utente
    row_ut = [
        sAM, "SI", ou_value, cn, cn, cn, given, surn,
        codice_fiscale, employee_id, department, description or "<PC>",
        "No", "", f"{sAM}@consip.it", f"{sAM}@consip.it", mobile,
        "", inserimento_gruppo, "", "", telephone_number, company
    ]
    # Riga computer
    row_cp = [
        description or "", "", f"{sAM}@consip.it", "", f"\"{mobile}\"", "", f"\"{cn}\"", "", "", ""
    ]

    st.markdown(f"""
Ciao.  
Si richiede modifiche come da file:  
- `{basename}_computer.csv`  (oggetti di tipo computer)  
- `{basename}_utente.csv`  (oggetti di tipo utenze)  
Archiviati al percorso:  
`\\\\\srv_dati.consip.tesoro.it\AreaCondivisa\DEPSI\IC\AD_Modifiche`  
Grazie
"""
    )
    st.subheader("Anteprima CSV Utente")
    st.dataframe(pd.DataFrame([row_ut], columns=HEADER_UTENTE))
    st.subheader("Anteprima CSV Computer")
    st.dataframe(pd.DataFrame([row_cp], columns=HEADER_COMPUTER))

    # Download
    buf_ut = io.StringIO()
    w_ut = csv.writer(buf_ut, quoting=csv.QUOTE_NONE, escapechar="\\")
    w_ut.writerow(HEADER_UTENTE)
    w_ut.writerow(row_ut)
    buf_ut.seek(0)
    buf_cp = io.StringIO()
    w_cp = csv.writer(buf_cp, quoting=csv.QUOTE_NONE, escapechar="\\")
    w_cp.writerow(HEADER_COMPUTER)
    w_cp.writerow(row_cp)
    buf_cp.seek(0)

    st.download_button(
        "📥 Scarica CSV Utente",
        data=buf_ut.getvalue(),
        file_name=f"{basename}_utente.csv",
        mime="text/csv"
    )
    st.download_button(
        "📥 Scarica CSV Computer",
        data=buf_cp.getvalue(),
        file_name=f"{basename}_computer.csv",
        mime="text/csv"
    )
    st.success(f"✅ CSV generati per '{sAM}'")
