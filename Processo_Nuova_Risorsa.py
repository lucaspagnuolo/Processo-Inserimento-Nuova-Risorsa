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
    # Sezioni
    ou_df = cfg[cfg["Section"] == "OU"][["Key/App","Label/Gruppi/Value"]].rename(
        columns={"Key/App": "key", "Label/Gruppi/Value": "label"})
    ou_options = dict(zip(ou_df["key"], ou_df["label"]))

    grp_df = cfg[cfg["Section"] == "InserimentoGruppi"][["Key/App","Label/Gruppi/Value"]].rename(
        columns={"Key/App": "app", "Label/Gruppi/Value": "gruppi"})
    gruppi = dict(zip(grp_df["app"], grp_df["gruppi"]))

    def_df = cfg[cfg["Section"] == "Defaults"][["Key/App","Label/Gruppi/Value"]].rename(
        columns={"Key/App": "key", "Label/Gruppi/Value": "value"})
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
    help="Deve contenere il foglio 'Risorsa Interna' con campo Section"
)
if not config_file:
    st.warning("Per favore carica il file di configurazione per continuare.")
    st.stop()

ou_options, gruppi, defaults = load_config_from_bytes(config_file.read())

# ------------------------------------------------------------
# Lettura Defaults
# ------------------------------------------------------------
dl_standard = defaults.get("dl_standard", "").split(";")
dl_vip = defaults.get("dl_vip", "").split(";")
o365_groups = [
    defaults.get("grp_o365_standard", "O365 Utenti Standard"),
    defaults.get("grp_o365_teams", "O365 Teams Premium"),
    defaults.get("grp_o365_copilot", "O365 Copilot Plus")
    defaults.get("grp_o365_viva", "O365 VivaEngage")
]
grp_foorban = defaults.get("grp_foorban", "Foorban_Users")
pillole = defaults.get("pillole", "Pillole formative Teams Premium")

# ------------------------------------------------------------
# Utility
# ------------------------------------------------------------
def formatta_data(data: str) -> str:
    for sep in ["-", "/"]:
        try:
            g, m, a = map(int, data.split(sep))
            dt = datetime(a, m, g) + timedelta(days=1)
            return dt.strftime("%m/%d/%Y 00:00")
        except Exception:
            continue
    return data

# Algoritmo genera_samaccountname
# ------------------------------------------------------------
def genera_samaccountname(
    nome: str,
    cognome: str,
    secondo_nome: str = "",
    secondo_cognome: str = "",
    esterno: bool = False
) -> str:
    n, sn = nome.strip().lower(), secondo_nome.strip().lower()
    c, sc = cognome.strip().lower(), secondo_cognome.strip().lower()
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

# Funzione per nome completo
# ------------------------------------------------------------
def build_full_name(
    cognome: str,
    secondo_cognome: str,
    nome: str,
    secondo_nome: str,
    esterno: bool = False
) -> str:
    parts = [p for p in [cognome, secondo_cognome, nome, secondo_nome] if p]
    return " ".join(parts) + (" (esterno)" if esterno else "")

# Header CSV Interna
HEADER = [
    "sAMAccountName",
    "Creation",
    "OU",
    "Name",
    "DisplayName",
    "cn",
    "GivenName",
    "Surname",
    "employeeNumber",
    "employeeID",
    "department",
    "Description",
    "passwordNeverExpired",
    "ExpireDate",
    "userprincipalname",
    "mail",
    "mobile",
    "RimozioneGruppo",
    "InserimentoGruppo",
    "disable",
    "moveToOU",
    "telephoneNumber",
    "company"
]

# ------------------------------------------------------------
# Input
# ------------------------------------------------------------
st.subheader("Modulo Inserimento Nuova Risorsa Interna")
employee_id = st.text_input(
    "Matricola",
    defaults.get("employee_id_default", "")
).strip()
cognome = st.text_input("Cognome").strip().capitalize()
secondo_cognome = st.text_input("Secondo Cognome").strip().capitalize()
nome = st.text_input("Nome").strip().capitalize()
secondo_nome = st.text_input("Secondo Nome").strip().capitalize()
codice_fiscale = st.text_input("Codice Fiscale", "").strip()
department = st.text_input(
    "Sigla Divisione-Area",
    defaults.get("department_default", "")
).strip()
numero_telefono = st.text_input("Mobile (+39 già inserito)", "").replace(" ", "")
description = st.text_input(
    "PC (lascia vuoto per <PC>)",
    "<PC>"
).strip()

resident_flag = st.checkbox("È Resident?")
numero_fisso_input = ""
if resident_flag:
    numero_fisso_input = st.text_input(
        "Numero fisso Resident (+39 già inserito)",
        ""
    ).strip()
telephone_default = defaults.get("telephone_interna", "")
telephone_number = (
    f"+39 {numero_fisso_input}" if resident_flag and numero_fisso_input else telephone_default
)

# Tipologia Utente
ou_keys = list(ou_options.keys())
ou_vals = list(ou_options.values())
def_o = defaults.get("ou_default", ou_vals[0] if ou_vals else "")
label_ou = st.selectbox("Tipologia Utente", ou_vals, index=ou_vals.index(def_o))
selected_key = ou_keys[ou_vals.index(label_ou)]
ou_value = ou_options[selected_key]

inserimento_gruppo = gruppi.get("interna", "")
company = defaults.get("company_interna", "")

# Configurazione Data e SM
st.subheader("Configurazione Data Operatività e Profilazione SM")
data_operativa = st.text_input(
    "In che giorno prende operatività? (gg/mm/aaaa)",
    ""
).strip()
profilazione_flag = st.checkbox("Deve essere profilato su qualche SM?")
sm_lines = []
if profilazione_flag:
    sm_lines = st.text_area(
        "SM su quali va profilato",
        "",
        placeholder="Inserisci una SM per riga"
    ).splitlines()

# Scegli DL
dl_list = dl_standard if selected_key == "utenti_standard" else dl_vip if selected_key == "utenti_vip" else []

# ------------------------------------------------------------
# Preview Messaggio
# ------------------------------------------------------------
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
| cell              | +39 {numero_telefono}                      |
"""
    st.markdown("Ciao.  \nRichiedo cortesemente la definizione di una casella di posta come sottoindicato.")
    st.markdown(table_md)
    st.markdown(f"Inviare batch di notifica migrazione mail a: imac@consip.it  \nAggiungere utenza di dominio ai gruppi:\n{groups_md}")
    if dl_list:
        st.markdown(f"Il giorno **{data_operativa}** occorre inserire la casella nelle DL:")
        for dl in dl_list:
            st.markdown(f"- {dl}")
    if profilazione_flag:
        st.markdown("Profilare su SM:")
        for sm in sm_lines:
            st.markdown(f"- {sm}")
    st.markdown(f"Aggiungere utenza al:\n- gruppo Azure: {grp_foorban}\n- canale {pillole}")
    st.markdown("Grazie  \nSaluti")

# ------------------------------------------------------------
# Generazione CSV Interna
# ------------------------------------------------------------
if st.button("Genera CSV Interna"):
    sAM = genera_samaccountname(nome, cognome, secondo_nome, secondo_cognome, False)
    cn = build_full_name(cognome, secondo_cognome, nome, secondo_nome, False)
    given = f"{nome} {secondo_nome}".strip()
    surn = f"{cognome} {secondo_cognome}".strip()
    mobile = f"+39 {numero_telefono}" if numero_telefono else ""
    row = [sAM, "SI", ou_value, cn.replace(" (esterno)", ""), cn, cn, given, surn,
           codice_fiscale, employee_id, department, description or "<PC>", "No", "",
           f"{sAM}@consip.it", f"{sAM}@consip.it", mobile, "", inserimento_gruppo, "", "",
           telephone_number, company]
    for i in (2, 3, 4, 5): row[i] = f"\"{row[i]}\""
    if secondo_nome: row[6] = f"\"{row[6]}\""
    if secondo_cognome: row[7] = f"\"{row[7]}\""
    for i in (16, 21): row[i] = f"\"{row[i]}\""
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

# ------------------------------------------------------------
# Generazione CSV Computer (versione corretta)
# ------------------------------------------------------------
if st.button("Genera CSV Computer"):
    # Generazione valori
    sAM = genera_samaccountname(nome, cognome, secondo_nome, secondo_cognome, False)
    cn = build_full_name(cognome, secondo_cognome, nome, secondo_nome, False)
    mobile = f"+39 {numero_telefono}" if numero_telefono else ""
    comp = description or ""

    # Definisci header
    comp_header = [
        "Computer", "OU", "add_mail", "remove_mail",
        "add_mobile", "remove_mobile",
        "add_userprincipalname", "remove_userprincipalname",
        "disable", "moveToOU"
    ]

    # Costruisci la riga: avvolgi manualmente in " " solo mobile e add_userprincipalname
    comp_row = [
        comp,
        "",  # OU
        f"{sAM}@consip.it",
        "",  # remove_mail
        f"\"{mobile}\"",  # add_mobile con doppio apice
        "",  # remove_mobile
        f"\"{cn}\"",  # add_userprincipalname con doppio apice
        "",  # remove_userprincipalname
        "",  # disable
        ""   # moveToOU
    ]

    # Anteprima
    df_comp = pd.DataFrame([comp_row], columns=comp_header)
    st.dataframe(df_comp)

    # Generazione CSV con QUOTE_NONE
    buf2 = io.StringIO()
    writer2 = csv.writer(buf2, quoting=csv.QUOTE_NONE, escapechar="\\")
    writer2.writerow(comp_header)
    writer2.writerow(comp_row)
    buf2.seek(0)

    st.download_button(
        label="📥 Scarica CSV Computer",
        data=buf2.getvalue(),
        file_name=f"{cognome}_{nome[:1]}_computer.csv",
        mime="text/csv"
    )
    st.success("✅ File CSV Computer generato correttamente")
