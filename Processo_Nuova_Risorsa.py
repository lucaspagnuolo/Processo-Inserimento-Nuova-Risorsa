import streamlit as st
import csv
import pandas as pd
from datetime import datetime, timedelta
import io

# ------------------------------------------------------------
# Caricamento configurazione da Excel caricato dall'utente
# ------------------------------------------------------------
def load_config_from_bytes(data: bytes):
    cfg = pd.read_excel(io.BytesIO(data), sheet_name=None)
    # Mappa key→label per OU
    ou_df = cfg.get("OU", pd.DataFrame(columns=["key", "label"]))
    ou_options = dict(zip(ou_df["key"], ou_df["label"]))
    # Mappa app→gruppi
    grp_df = cfg.get("InserimentoGruppi", pd.DataFrame(columns=["app", "gruppi"]))
    gruppi = dict(zip(grp_df["app"], grp_df["gruppi"]))
    # Mappa key→value per defaults
    def_df = cfg.get("Defaults", pd.DataFrame(columns=["key", "value"]))
    defaults = dict(zip(def_df["key"], def_df["value"]))
    return ou_options, gruppi, defaults

# File uploader per configurazione
st.title("1.1 Nuova Risorsa Interna - Configurable")
config_file = st.file_uploader(
    "Carica il file di configurazione (config.xlsx)",
    type=["xlsx"],
    help="Deve contenere i fogli OU, InserimentoGruppi e Defaults"
)
if not config_file:
    st.warning("Per favore carica il file di configurazione per continuare.")
    st.stop()

# Leggi configurazione
ou_options, gruppi, defaults = load_config_from_bytes(config_file.read())

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
    limit = 16 if esterno else 20
    cand = f"{n}{sn}.{c}{sc}"
    if len(cand) <= limit: return cand + suffix
    cand = f"{(n[:1])}{(sn[:1])}.{c}{sc}"
    if len(cand) <= limit: return cand + suffix
    return (f"{n[:1]}{sn[:1]}.{c}")[:limit] + suffix

def build_full_name(cognome: str, secondo_cognome: str,
                    nome: str, secondo_nome: str,
                    esterno: bool = False) -> str:
    parts = [p for p in [cognome, secondo_cognome, nome, secondo_nome] if p]
    full = " ".join(parts)
    return full + (" (esterno)" if esterno else "")

HEADER = [
    "sAMAccountName", "Creation", "OU", "Name", "DisplayName", "cn", "GivenName", "Surname",
    "employeeNumber", "employeeID", "department", "Description", "passwordNeverExpired",
    "ExpireDate", "userprincipalname", "mail", "mobile", "RimozioneGruppo", "InserimentoGruppo",
    "disable", "moveToOU", "telephoneNumber", "company"
]

# ------------------------------------------------------------
# App 1.1: Nuova Risorsa Interna
# ------------------------------------------------------------
st.subheader("Modulo Inserimento Nuova Risorsa Interna")

# Input ordinati e rinominati
employee_id        = st.text_input("Matricola", defaults.get("employee_id_default", "")).strip()
cognome            = st.text_input("Cognome").strip().capitalize()
secondo_cognome    = st.text_input("Secondo Cognome").strip().capitalize()
nome               = st.text_input("Nome").strip().capitalize()
secondo_nome       = st.text_input("Secondo Nome").strip().capitalize()
codice_fiscale     = st.text_input("Codice Fiscale", "").strip()
department         = st.text_input("Sigla Divisione-Area", defaults.get("department_default", "")).strip()
numero_telefono    = st.text_input("Mobile", "").replace(" ", "")
description        = st.text_input("PC (lascia vuoto per <PC>)", "<PC>").strip()

# OU a tendina da config (rinominato come Tipologia Utente)
ou_labels    = list(ou_options.values())
default_ou   = defaults.get("ou_default", ou_labels[0] if ou_labels else "")
label_ou     = st.selectbox("Tipologia Utente", options=ou_labels,
                             index=ou_labels.index(default_ou) if default_ou in ou_labels else 0)
selected_ou_key = list(ou_options.keys())[ou_labels.index(label_ou)]
ou_value     = ou_options[selected_ou_key]

# Altri valori fissi da config
inserimento_gruppo = gruppi.get("interna", "")
telephone_number   = defaults.get("telephone_interna", "")
company            = defaults.get("company_interna", "")

if st.button("Genera CSV Interna"):
    sAM = genera_samaccountname(nome, cognome, secondo_nome, secondo_cognome, False)
    cn  = build_full_name(cognome, secondo_cognome, nome, secondo_nome, False)
    row = [
        sAM, "SI", ou_value, cn.replace(" (esterno)", ""), cn, cn,
        " ".join([nome, secondo_nome]).strip(),
        " ".join([cognome, secondo_cognome]).strip(),
        codice_fiscale, employee_id, department,
        description or "<PC>", "No", "",
        f"{sAM}@consip.it",
        f"{sAM}@consip.it",
        f"+39 {numero_telefono}" if numero_telefono else "",
        "", inserimento_gruppo, "", "",
        telephone_number, company
    ]
    # Scrittura CSV in memoria
    buf = io.StringIO()
    writer = csv.writer(buf, quoting=csv.QUOTE_MINIMAL)
    writer.writerow(HEADER)
    writer.writerow(row)
    buf.seek(0)
    # Anteprima e download
    df = pd.DataFrame([row], columns=HEADER)
    st.dataframe(df)
    st.download_button(
        label="📥 Scarica CSV",
        data=buf.getvalue(),
        file_name=f"{cognome}_{nome[:1]}_interno.csv",
        mime="text/csv"
    )
    st.success(f"✅ File CSV generato per '{sAM}'")
