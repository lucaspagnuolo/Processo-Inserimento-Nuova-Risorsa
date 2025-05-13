import streamlit as st
import csv
import pandas as pd
from datetime import datetime, timedelta
import io

# ------------------------------------------------------------
# Caricamento configurazione da Excel caricato dall'utente
# ------------------------------------------------------------
def load_config_from_bytes(data: bytes):
    cfg = pd.read_excel(io.BytesIO(data), sheet_name="Risorsa Interna")
    # Sezione OU
    ou_df = cfg[cfg["Section"] == "OU"][["Key/App", "Label/Gruppi/Value"]].rename(
        columns={"Key/App": "key", "Label/Gruppi/Value": "label"}
    )
    ou_options = dict(zip(ou_df["key"], ou_df["label"]))
    # Sezione InserimentoGruppi
    grp_df = cfg[cfg["Section"] == "InserimentoGruppi"][["Key/App", "Label/Gruppi/Value"]].rename(
        columns={"Key/App": "app", "Label/Gruppi/Value": "gruppi"}
    )
    gruppi = dict(zip(grp_df["app"], grp_df["gruppi"]))
    # Sezione Defaults
    def_df = cfg[cfg["Section"] == "Defaults"][["Key/App", "Label/Gruppi/Value"]].rename(
        columns={"Key/App": "key", "Label/Gruppi/Value": "value"}
    )
    defaults = dict(zip(def_df["key"], def_df["value"]))
    return ou_options, gruppi, defaults

# ------------------------------------------------------------
# App 1.1: Nuova Risorsa Interna
# ------------------------------------------------------------
st.set_page_config(page_title="1.1 Nuova Risorsa Interna")
st.title("1.1 Nuova Risorsa Interna")

config_file = st.file_uploader(
    "Carica il file di configurazione (config.xlsx)",
    type=["xlsx"],
    help="Deve contenere il foglio “Risorsa Interna” con campo Section"
)
if not config_file:
    st.warning("Per favore carica il file di configurazione per continuare.")
    st.stop()

ou_options, gruppi, defaults = load_config_from_bytes(config_file.read())

# Lettura DL default da Defaults
dl_standard = defaults.get("dl_standard", "").split(";")  # utenti.consip@...;...
dl_vip = defaults.get("dl_vip", "").split(";")      # utenti.consip@...;...

# Estrazione gruppi O365 da Defaults
o365_groups = [
    defaults.get("grp_o365_standard", "O365 Utenti Standard"),
    defaults.get("grp_o365_teams", "O365 Teams Premium"),
    defaults.get("grp_o365_copilot", "O365 Copilot Plus")
]

# ------------------------------------------------------------
# Utility functions
# ------------------------------------------------------------
def formatta_data(data: str) -> str:
    for sep in ["-", "/"]:
        try:
            g, m, a = map(int, data.split(sep))
            dt = datetime(a, m, g) + timedelta(days=1)
            return dt.strftime("%m/%d/%Y 00:00")
        except:
            continue
    return data

def genera_samaccountname(nome: str, cognome: str,
                          secondo_nome: str = "", secondo_cognome: str = "",
                          esterno: bool = False) -> str:
    n, sn = nome.strip().lower(), secondo_nome.strip().lower()
    c, sc = cognome.strip().lower(), secondo_cognome.strip().lower()
    suffix = ".ext" if esterno else ""
    limit  = 16 if esterno else 20
    cand   = f"{n}{sn}.{c}{sc}"
    if len(cand) <= limit:
        return cand + suffix
    cand = f"{n[:1]}{sn[:1]}.{c}{sc}"
    if len(cand) <= limit:
        return cand + suffix
    return (f"{n[:1]}{sn[:1]}.{c}")[:limit] + suffix

def build_full_name(cognome: str, secondo_cognome: str,
                    nome: str, secondo_nome: str,
                    esterno: bool = False) -> str:
    parts = [p for p in [cognome, secondo_cognome, nome, secondo_nome] if p]
    full  = " ".join(parts)
    return full + (" (esterno)" if esterno else "")

HEADER = [
    "sAMAccountName","Creation","OU","Name","DisplayName","cn","GivenName","Surname",
    "employeeNumber","employeeID","department","Description","passwordNeverExpired",
    "ExpireDate","userprincipalname","mail","mobile","RimozioneGruppo","InserimentoGruppo",
    "disable","moveToOU","telephoneNumber","company"
]

# ------------------------------------------------------------
# Form di input
# ------------------------------------------------------------
st.subheader("Modulo Inserimento Nuova Risorsa Interna")

employee_id        = st.text_input("Matricola", defaults.get("employee_id_default", "")).strip()
cognome            = st.text_input("Cognome").strip().capitalize()
secondo_cognome    = st.text_input("Secondo Cognome").strip().capitalize()
nome               = st.text_input("Nome").strip().capitalize()
secondo_nome       = st.text_input("Secondo Nome").strip().capitalize()
codice_fiscale     = st.text_input("Codice Fiscale", "").strip()
department         = st.text_input("Sigla Divisione-Area", defaults.get("department_default", "")).strip()
numero_telefono    = st.text_input("Mobile (+39 già inserito)", "").replace(" ", "")
description        = st.text_input("PC (lascia vuoto per <PC>)", "<PC>").strip()

# Flag Resident e Numero Fisso
resident_flag      = st.checkbox("È Resident?")
numero_fisso_input = ""
if resident_flag:
    numero_fisso_input = st.text_input("Numero fisso Resident (+39 già inserito)", "").strip()
telephone_default  = defaults.get("telephone_interna", "")
telephone_number   = f"+39 {numero_fisso_input}" if resident_flag and numero_fisso_input else telephone_default

# Tipologia Utente
ou_labels    = list(ou_options.values())
default_ou   = defaults.get("ou_default", ou_labels[0] if ou_labels else "")
label_ou_key = list(ou_options.keys())[ou_labels.index(default_ou)]
label_ou     = st.selectbox("Tipologia Utente", options=ou_labels,
                             index=ou_labels.index(default_ou) if default_ou in ou_labels else 0)
ou_value     = ou_options[label_ou_key]

inserimento_gruppo = gruppi.get("interna", "")
company            = defaults.get("company_interna", "")

# ------------------------------------------------------------
# Parametri in base alla Tipologia UT
# ------------------------------------------------------------
selected_tipologia = label_ou_key  # e.g. 'utenti_standard' o 'utenti_vip'
if selected_tipologia == 'utenti_standard':
    dl_list = dl_standard
elif selected_tipologia == 'utenti_vip':
    dl_list = dl_vip
else:
    dl_list = []

# ------------------------------------------------------------
# Preview Messaggio
# ------------------------------------------------------------
if st.button("Template per Posta Elettronica"):
    sAM = genera_samaccountname(nome, cognome, secondo_nome, secondo_cognome, False)
    cn  = build_full_name(cognome, secondo_cognome, nome, secondo_nome, False)
    groups_md = "\n".join([f"- {g}" for g in o365_groups])

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
| cell              | +39 {numero_telefono}                      |
"""
    st.markdown("Ciao.  \nRichiedo cortesemente la definizione di una casella di posta come sottoindicato.")
    st.markdown(table_md)
    st.markdown(f"Inviare batch di notifica migrazione mail a: imac@consip.it  \n" + 
                f"Aggiungere utenza di dominio ai gruppi:\n{groups_md}")
    # DL default
    if dl_list:
        st.markdown("Case da inserire nelle DL (default):")
        for dl in dl_list:
            if dl.strip(): st.markdown(f"- {dl}")
    st.markdown("Grazie  \nSaluti")

# ------------------------------------------------------------
# Generazione CSV
# ------------------------------------------------------------
if st.button("Genera CSV Interna"):
    sAM = genera_samaccountname(nome, cognome, secondo_nome, secondo_cognome, False)
    cn  = build_full_name(cognome, secondo_cognome, nome, secondo_nome, False)
    given = f"{nome} {secondo_nome}".strip()
    surn  = f"{cognome} {secondo_cognome}".strip()
    mobile = f"+39 {numero_telefono}" if numero_telefono else ""

    row = [
        sAM, "SI", ou_value,
        cn.replace(" (esterno)", ""),
        cn, cn, given, surn,
        codice_fiscale, employee_id, department,
        description or "<PC>", "No", "",
        f"{sAM}@consip.it", f"{sAM}@consip.it",
        mobile,
        "", inserimento_gruppo, "", "",
        telephone_number,
        company
    ]
    for i in (2,3,4,5): row[i] = f"\"{row[i]}\""
    if secondo_nome: row[6] = f"\"{row[6]}\""
    if secondo_cognome: row[7] = f"\"{row[7]}\""
    for i in (16,21): row[i] = f"\"{row[i]}\""

    buf = io.StringIO()
    writer = csv.writer(buf, quoting=csv.QUOTE_NONE, escapechar="\\")
    writer.writerow(HEADER)
    writer.writerow(row)
    buf.seek(0)

    st.dataframe(pd.DataFrame([row], columns=HEADER))
    st.download_button(
        label="📥 Scarica CSV",
        data=buf.getvalue(),
        file_name=f"{cognome}_{nome[:1]}_interno.csv",
        mime="text/csv"
    )
    st.success(f"✅ File CSV generato per '{sAM}'")
