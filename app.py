"""
H2 FAST - Progettazione di impianti di produzione e stoccaggio di idrogeno verde (toolkit H2READY).

Interfaccia Streamlit in tre step sul motore H2FAsT (progetto AMETHYST):
    1. Parametri      -> stesse sezioni del file INPUT.xlsx originale + fonti energetiche
    2. Elaborazione   -> riepilogo, controlli, calcolo di tutte le configurazioni
    3. Risultati      -> migliori progetti, flussi energetici, business plan, esplorazione, Excel di output
"""
import copy
import hashlib
import json
import traceback

import numpy as np
import pandas as pd
import requests
import streamlit as st

import fonti_energia as F
import grafici as G
import motore_h2fast as M

st.set_page_config(page_title="H2 FAST - Impianti H2", page_icon="⚡", layout="wide")

# =============================================================================================
# PARAMETRI PREDEFINITI (default della classe originale Analisi_finanziaria, salvo dove indicato)
# =============================================================================================
DEFAULT = {
    # --- fonti energetiche
    "modo_fonti": "API", "lat": 46.0637, "lon": 13.2358, "luoghi_separati": False, "lat_w": 46.0637,
    "lon_w": 13.2358, "anno": 2019, "perdite": 14.0, "angoli_ottimali": True, "inclinazione": 35.0,
    "azimut": 0.0, "p_PV": 1000.0, "p_Wind": 0.0, "v_cut_in": 3.0, "v_rated": 12.0, "v_cut_out": 25.0,
    "tipo_file": "SI", "usa_extra": False, "p_Extra": 0.0,
    # --- dati tecnici
    "batteria": "SI", "min_batt": 0.2, "max_batteria": 0.5, "eff_batt": 0.95, "min_elet": 0.2,
    "tassoDEN": 0.005, "bar": 300.0, "dP_el": 1.0, "dP_bat": 2.0,
    # --- investimenti
    "Terr": 0.0, "OpeE": 0.0, "StazzRif": 500000.0, "SpeTOpere": 0.0, "BombSto": 0.0, "LavoImp": 0.0,
    "CarrEll": 0.0, "ImpPV1eurokW": 800.0, "ImpWindeurokW": 1000.0, "ImpExtraeurokW": 2000.0,
    "EletteuroKW": 1650.0, "CompreuroKW": 4000.0, "AccuEeurokW": 200.0,
    "costounitariostoccaggio": 1200.0,   # originale 0: con 0 lo stoccaggio risultava gratuito
    "idrogstocperc": 0.1,
    # --- costi operativi
    "costlitroacqua": 0.035, "PercEserImp": 0.005, "Percentimpianti": 0.0025, "PercentOpeEd": 0.0005,
    "SpesAmmGen": 7000.0, "Affitto": 0.0, "CostiPersonal": 0.0, "AltriCost": 0.0, "IVAsualtriCost": "SI",
    # --- dati economico-finanziari
    "DurPianEcon": 20, "inflazione": 0.02, "inflazionePrezzoElet": 0.02, "tassoVAN": 0.1, "incentpubb": 0.0,
    "duratincentpubb": 0, "prezzoindrogeno": 10.0, "inflazioneIdrog": 0.01,
    "prezzoElett": 0.10,                 # originale 1 €/kWh (valore segnaposto)
    "ContrPubb": 0.0, "DebitoSenior": 0.8, "DurDebitoSenior": 20, "tassoDebito": 0.05, "FreqPagamenti": 1,
    "DurataPonte": 0, "tassoPonte": 0.0, "aliquoMedia": 0.275, "MaxInterssDed": 0.3, "Perciva": 0.22,
    # --- opzioni di output
    "attributo": "VAN", "n_progetti": 5, "lingua": "ITA", "relazione": "NO",
    "relazioni": [{"X": "PotEle", "Y": "VAN"}, {"X": "AccuE", "Y": "LCOH"}],
    "si_fa_simulazione": "NO", "attributo_simulazione": "prezzo", "si_fa_grafico_SA": "NO",
    "SA_variable_list": ["VAN", "LCOH"],
}
INTERI = {"anno", "DurPianEcon", "duratincentpubb", "DurDebitoSenior", "FreqPagamenti", "DurataPonte", "n_progetti"}
CHIAVI_PROFILO = ["modo_fonti", "lat", "lon", "luoghi_separati", "lat_w", "lon_w", "anno", "perdite",
                  "angoli_ottimali", "inclinazione", "azimut", "p_PV", "p_Wind", "v_cut_in", "v_rated",
                  "v_cut_out", "tipo_file", "usa_extra", "p_Extra"]
SI_NO = {"SI": "Sì", "NO": "No"}
CRITERI = ["VAN", "TIR", "LCOH", "PAYBACK", "costo_full_cost", "costo_medio_operativo", "prezzoindrogeno",
           "ProdAnnuaIdrogkg", "CapFac", "EnergiaAutocons", "ProdElettVend", "spegn_giorn", "PotEle", "AccuE",
           "potenza_compressore", "impianto_stocc", "costo_full_cost_levelized"]
VARIABILI_GRAFICI = [c for c in M.VARIABILI.keys()]


# =============================================================================================
# UTILITÀ
# =============================================================================================
def stato():
    s = st.session_state
    if "par" not in s:
        s.par = copy.deepcopy(DEFAULT)
    s.setdefault("step", 1)
    s.setdefault("file_profilo", None)      # (nome, bytes) del CSV principale
    s.setdefault("file_extra", None)        # (nome, bytes) del CSV della fonte extra
    s.setdefault("profilo", None)           # {"componenti":..., "totale":..., "firma":...}
    s.setdefault("risultato", None)
    s.setdefault("excel", None)
    s.setdefault("config_sel", None)
    return s


def eur(v, dec=0):
    if v is None or (isinstance(v, float) and not np.isfinite(v)):
        return "n.d."
    return f"€ {v:,.{dec}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def num(v, dec=0, um=""):
    if v is None or (isinstance(v, float) and not np.isfinite(v)):
        return "n.d."
    s = f"{v:,.{dec}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{s} {um}".strip()


def nome_var(codice, unita=True):
    return M.nome_variabile(codice, "ITA", unita)


def pulisci_widget():
    for k in list(st.session_state.keys()):
        if k.startswith("w_"):
            del st.session_state[k]


# ---- campi persistenti tra gli step: il valore vive in st.session_state.par -------------------
def _init(nome, valore):
    k = "w_" + nome
    if k not in st.session_state:
        st.session_state[k] = valore
    return k


def campo(nome, etichetta, perc=False, minimo=0.0, massimo=None, passo=None, help=None, fmt=None, disabled=False):
    par = st.session_state.par
    intero = nome in INTERI
    v0 = par[nome] * 100 if perc else par[nome]
    v0 = int(round(v0)) if intero else float(v0)
    k = _init(nome, v0)
    conv = int if intero else float
    v = st.number_input(etichetta, key=k, min_value=conv(minimo) if minimo is not None else None,
                        max_value=conv(massimo) if massimo is not None else None,
                        step=conv(passo) if passo is not None else (1 if intero else None),
                        format=fmt or ("%d" if intero else ("%.2f" if perc else None)),
                        help=help, disabled=disabled)
    par[nome] = v / 100 if perc else v
    return par[nome]


def scelta(nome, etichetta, opzioni, fmt=None, help=None, orizzontale=False):
    par = st.session_state.par
    if par[nome] not in opzioni:
        par[nome] = opzioni[0]
    k = _init(nome, par[nome])
    if orizzontale:
        v = st.radio(etichetta, opzioni, key=k, format_func=fmt or str, help=help, horizontal=True)
    else:
        v = st.selectbox(etichetta, opzioni, key=k, format_func=fmt or str, help=help)
    par[nome] = v
    return v


def interruttore(nome, etichetta, help=None):
    par = st.session_state.par
    k = _init(nome, bool(par[nome]))
    par[nome] = st.toggle(etichetta, key=k, help=help)
    return par[nome]


def si_no(nome, etichetta, help=None):
    return scelta(nome, etichetta, ["SI", "NO"], fmt=lambda x: SI_NO[x], help=help, orizzontale=True)


# ---- servizi esterni (con cache) -------------------------------------------------------------
@st.cache_data(show_spinner=False, ttl=7 * 86400)
def pvgis(lat, lon, kwp, anno, perdite, ottimali, incl, azim):
    return F.scarica_pvgis(lat, lon, kwp, anno, perdite, ottimali, incl, azim)


@st.cache_data(show_spinner=False, ttl=7 * 86400)
def vento(lat, lon, anno):
    return F.scarica_vento(lat, lon, anno)


@st.cache_data(show_spinner=False, ttl=86400)
def cerca_localita(testo):
    r = requests.get("https://nominatim.openstreetmap.org/search",
                     params={"q": testo, "format": "json", "limit": 1, "countrycodes": "it,si,at,fr,de,ch"},
                     headers={"User-Agent": "H2FAST-H2READY/1.0"}, timeout=20)
    r.raise_for_status()
    d = r.json()
    return (float(d[0]["lat"]), float(d[0]["lon"]), d[0]["display_name"]) if d else None


def firma_profilo(par, s):
    base = {k: par[k] for k in CHIAVI_PROFILO}
    for nome in ("file_profilo", "file_extra"):
        f = s.get(nome)
        base[nome] = hashlib.md5(f[1]).hexdigest() if f else None
    return json.dumps(base, sort_keys=True, default=str)


def costruisci_profilo(par, s):
    comp = {}
    if par["modo_fonti"] == "API":
        comp["Fotovoltaico"] = pvgis(round(par["lat"], 4), round(par["lon"], 4), par["p_PV"], par["anno"],
                                     par["perdite"], par["angoli_ottimali"], par["inclinazione"], par["azimut"])
        if par["p_Wind"] > 0:
            lat_w, lon_w = (par["lat_w"], par["lon_w"]) if par["luoghi_separati"] else (par["lat"], par["lon"])
            v = vento(round(lat_w, 4), round(lon_w, 4), par["anno"])
            comp["Eolico"] = F.potenza_eolica(v, par["p_Wind"], par["v_cut_in"], par["v_rated"], par["v_cut_out"])
    else:
        if not s.file_profilo:
            raise F.ErroreProfilo("Carica il file CSV con il profilo orario.")
        comp["Fotovoltaico"] = F.leggi_profilo(s.file_profilo[1], par["tipo_file"] == "SI")
    if par["usa_extra"]:
        if not s.file_extra:
            raise F.ErroreProfilo("Hai attivato la fonte extra: carica il suo CSV (8760 valori in kW).")
        comp["Fonte extra"] = F.leggi_profilo(s.file_extra[1], False)
    totale = np.sum(list(comp.values()), axis=0)
    return {"componenti": comp, "totale": totale, "firma": firma_profilo(par, s)}


def parametri_motore(par):
    """Dizionario per Analisi_combinata (stessi nomi del file INPUT.xlsx)."""
    p = {k: v for k, v in par.items() if k not in ("relazioni",)}
    if par["modo_fonti"] != "API":
        p["p_Wind"] = 0.0
    if not par["usa_extra"]:
        p["p_Extra"] = 0.0
    p["variable_1_list"] = [r["X"] for r in par["relazioni"] if r.get("X") and r.get("Y")]
    p["variable_2_list"] = [r["Y"] for r in par["relazioni"] if r.get("X") and r.get("Y")]
    p["file_csv"], p["tipo_file"] = None, "NO"
    return p


def potenza_totale(par):
    pw = par["p_Wind"] if par["modo_fonti"] == "API" else 0.0
    pe = par["p_Extra"] if par["usa_extra"] else 0.0
    return par["p_PV"] + pw + pe


# =============================================================================================
# INTESTAZIONE, ACCESSO H2READY, STEPPER
# =============================================================================================
def accesso_h2ready():
    """Se l'app ha i Secrets del foglio master H2READY usa il blocco di accesso del toolkit (categoria 'fast').
    Senza Secrets l'app funziona da sola (utile per sviluppo e prove)."""
    try:
        configurato = "gsheets" in st.secrets.get("connections", {})
        attivo = st.secrets.get("h2fast", {}).get("usa_accesso_h2ready", True)
    except Exception:
        return None
    if not (configurato and attivo):
        return None
    import h2ready as H
    comune = H.blocco_accesso("H2 FAST - Progettazione impianti di produzione e stoccaggio H2",
                              categoria="fast", consenti_manuale=True)
    if comune is None:
        st.stop()
    return comune


PASSI = ["Parametri", "Elaborazione", "Risultati"]


def vai(n):
    st.session_state.step = n


def stepper(s):
    st.markdown("""
    <style>
    .h2-step {border-radius: 10px; padding: 10px 14px; border: 1px solid rgba(128,128,128,.3);}
    .h2-step.attivo {border: 2px solid #2a78d6; background: rgba(42,120,214,.08);}
    .h2-step.fatto {opacity: .75;}
    .h2-step b {font-size: 1.05rem;}
    .h2-step span {display:block; font-size: .8rem; opacity:.75;}
    </style>""", unsafe_allow_html=True)
    cols = st.columns(3)
    descr = ["dati come nell'INPUT.xlsx", "controlli e calcolo", "progetti, grafici, Excel"]
    for i, (c, nome) in enumerate(zip(cols, PASSI), start=1):
        classe = "attivo" if s.step == i else ("fatto" if s.step > i else "")
        segno = "✓ " if s.step > i else f"{i}. "
        c.markdown(f"<div class='h2-step {classe}'><b>{segno}{nome}</b><span>{descr[i - 1]}</span></div>",
                   unsafe_allow_html=True)
    st.write("")


# =============================================================================================
# STEP 1 - PARAMETRI
# =============================================================================================
def sezione_import(s):
    with st.expander("📂 Importa o salva i parametri (INPUT.xlsx originale o file .json)"):
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("**Importa**")
            f = st.file_uploader("INPUT.xlsx di H2FAST oppure parametri .json salvati da questa app",
                                 type=["xlsx", "json"], key="up_parametri")
            if f is not None and st.button("Importa questi parametri", type="primary"):
                try:
                    if f.name.lower().endswith(".json"):
                        nuovi = json.loads(f.getvalue().decode("utf-8"))
                    else:
                        nuovi = M.leggi_input_excel(f)
                        if nuovi.get("variable_1_list"):
                            nuovi["relazioni"] = [{"X": x, "Y": y} for x, y in
                                                  zip(nuovi["variable_1_list"], nuovi["variable_2_list"])]
                    ignorati = []
                    for k, v in nuovi.items():
                        if k in DEFAULT:
                            if k in INTERI:
                                v = int(v)
                            elif isinstance(DEFAULT[k], float):
                                v = float(v)
                            s.par[k] = v
                        elif k not in ("file_csv", "variable_1_list", "variable_2_list"):
                            ignorati.append(k)
                    if f.name.lower().endswith(".xlsx"):
                        s.par["modo_fonti"] = "CSV"
                    pulisci_widget()
                    s.pop("_rel_base", None)
                    st.success("Parametri importati." + (f" Ignorati: {', '.join(ignorati)}" if ignorati else ""))
                    if f.name.lower().endswith(".xlsx"):
                        st.info("Dall'Excel arriva solo il nome del file orario: carica il CSV nella scheda "
                                "'Fonti energetiche'.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Non riesco a leggere il file: {e}")
        with c2:
            st.markdown("**Salva**")
            st.caption("Scarica i parametri attuali per riprendere il lavoro in un secondo momento.")
            st.download_button("Scarica parametri (.json)", json.dumps(s.par, indent=2, default=str),
                               file_name="h2fast_parametri.json", mime="application/json")
            if st.button("Ripristina i valori predefiniti"):
                s.par = copy.deepcopy(DEFAULT)
                pulisci_widget()
                s.pop("_rel_base", None)
                st.rerun()


def tab_fonti(s):
    par = s.par
    scelta("modo_fonti", "Da dove arrivano i dati orari di produzione?", ["API", "CSV"],
           fmt=lambda x: {"API": "🌍 Scarica da PVGIS e Open-Meteo (mappa)",
                          "CSV": "📄 Carico io un file CSV orario"}[x], orizzontale=True)

    if par["modo_fonti"] == "API":
        try:
            import folium
            from streamlit_folium import st_folium
            mappa_ok = True
        except ImportError:
            mappa_ok = False

        c_cerca, c_bott = st.columns([4, 1])
        testo = c_cerca.text_input("Cerca una località", placeholder="es. Udine, oppure un indirizzo",
                                   key="cerca_loc")
        c_bott.write("")
        c_bott.write("")
        if c_bott.button("Cerca", width="stretch") and testo:
            try:
                ris = cerca_localita(testo)
                if ris:
                    par["lat"], par["lon"] = round(ris[0], 4), round(ris[1], 4)
                    if not par["luoghi_separati"]:
                        par["lat_w"], par["lon_w"] = par["lat"], par["lon"]
                    for k in ("w_lat", "w_lon", "w_lat_w", "w_lon_w"):
                        s.pop(k, None)
                    st.toast(ris[2][:120])
                    st.rerun()
                else:
                    st.warning("Località non trovata.")
            except Exception as e:
                st.warning(f"Ricerca non disponibile: {e}")

        interruttore("luoghi_separati", "📍 Fotovoltaico ed eolico in due luoghi diversi")

        def mappa(chiave_lat, chiave_lon, titolo, key):
            st.markdown(f"**{titolo}**")
            c1, c2 = st.columns(2)
            with c1:
                campo(chiave_lat, "Latitudine", minimo=-90, massimo=90, passo=0.01, fmt="%.4f")
            with c2:
                campo(chiave_lon, "Longitudine", minimo=-180, massimo=180, passo=0.01, fmt="%.4f")
            if not mappa_ok:
                return
            m = folium.Map(location=[par[chiave_lat], par[chiave_lon]], zoom_start=8)
            folium.Marker([par[chiave_lat], par[chiave_lon]]).add_to(m)
            out = st_folium(m, height=330, use_container_width=True, key=key, returned_objects=["last_clicked"])
            click = (out or {}).get("last_clicked")
            if click:
                pos = (round(click["lat"], 4), round(click["lng"], 4))
                if pos != s.get(f"_click_{key}"):
                    s[f"_click_{key}"] = pos
                    par[chiave_lat], par[chiave_lon] = pos
                    s.pop("w_" + chiave_lat, None)
                    s.pop("w_" + chiave_lon, None)
                    st.rerun()
            st.caption("Clicca sulla mappa per spostare il punto.")

        if par["luoghi_separati"]:
            cm1, cm2 = st.columns(2)
            with cm1:
                mappa("lat", "lon", "☀️ Fotovoltaico", "mappa_pv")
            with cm2:
                mappa("lat_w", "lon_w", "💨 Eolico", "mappa_vento")
        else:
            mappa("lat", "lon", "☀️💨 Sito dell'impianto", "mappa_unica")
            par["lat_w"], par["lon_w"] = par["lat"], par["lon"]

        st.markdown("##### Impianti")
        c1, c2, c3 = st.columns(3)
        with c1:
            campo("p_PV", "Fotovoltaico [kWp]", passo=50.0, help="Excel: PV system peak power")
            campo("anno", "Anno meteo di riferimento", minimo=2005, massimo=2023,
                  help="PVGIS 5.3 copre 2005-2023; se non disponibile si passa a PVGIS 5.2 (2005-2020).")
        with c2:
            campo("perdite", "Perdite di sistema FV [%]", minimo=0, massimo=50, passo=1.0)
            interruttore("angoli_ottimali", "Inclinazione e orientamento ottimali (PVGIS)")
            if not par["angoli_ottimali"]:
                campo("inclinazione", "Inclinazione [°]", minimo=0, massimo=90, passo=1.0)
                campo("azimut", "Orientamento [°] (0 = sud, -90 = est, 90 = ovest)", minimo=-180, massimo=180,
                      passo=5.0)
        with c3:
            campo("p_Wind", "Eolico [kW]", passo=50.0, help="0 = nessun impianto eolico")
            if par["p_Wind"] > 0:
                with st.expander("Curva di potenza della turbina"):
                    campo("v_cut_in", "Velocità di avvio [m/s]", passo=0.5)
                    campo("v_rated", "Velocità nominale [m/s]", passo=0.5)
                    campo("v_cut_out", "Velocità di arresto [m/s]", passo=0.5)
                    st.caption("Vento a 100 m da Open-Meteo (ERA5); curva cubica tra avvio e nominale.")
    else:
        c1, c2 = st.columns([3, 2])
        with c1:
            f = st.file_uploader("Profilo orario (CSV, 8760 valori)", type=["csv", "txt"], key="up_profilo")
            if f is not None:
                s.file_profilo = (f.name, f.getvalue())
            if s.file_profilo:
                st.caption(f"File in uso: **{s.file_profilo[0]}**")
        with c2:
            si_no("tipo_file", "Il file proviene da PVGIS?",
                  help="Sì: CSV orario scaricato da PVGIS (colonna P in W). No: una colonna di 8760 valori "
                       "in kW, separatore ';' e decimale ',' (come nell'originale) oppure ',' e '.'.")
            campo("p_PV", "Potenza nominale a cui si riferisce il profilo [kWp]", passo=50.0,
                  help="Deve essere la stessa potenza usata per generare il profilo: serve per la ricerca "
                       "delle taglie e per il costo dell'impianto.")

    with st.expander("➕ Fonte extra (idroelettrico, biomassa, misure reali...)", expanded=par["usa_extra"]):
        interruttore("usa_extra", "Aggiungi un profilo orario extra")
        if par["usa_extra"]:
            c1, c2 = st.columns([3, 2])
            with c1:
                f = st.file_uploader("CSV della fonte extra (8760 valori in kW)", type=["csv", "txt"],
                                     key="up_extra")
                if f is not None:
                    s.file_extra = (f.name, f.getvalue())
                if s.file_extra:
                    st.caption(f"File in uso: **{s.file_extra[0]}**")
            with c2:
                campo("p_Extra", "Potenza nominale fonte extra [kW]", passo=10.0)

    st.divider()
    firma = firma_profilo(par, s)
    aggiornato = s.profilo is not None and s.profilo["firma"] == firma
    c1, c2 = st.columns([1, 3])
    with c1:
        if st.button("⬇️ Carica / aggiorna il profilo" if not aggiornato else "✅ Profilo aggiornato",
                     type="primary" if not aggiornato else "secondary", width="stretch"):
            with st.spinner("Recupero dei dati orari..."):
                try:
                    s.profilo = costruisci_profilo(par, s)
                    st.rerun()
                except F.ErroreProfilo as e:
                    st.error(str(e))
                except Exception as e:
                    st.error(f"Errore inatteso: {e}")
    with c2:
        if s.profilo is None:
            st.info("Carica il profilo orario per vedere la produzione e passare all'elaborazione.")
        elif not aggiornato:
            st.warning("Hai cambiato località, potenze o file: aggiorna il profilo prima di proseguire.")
    if s.profilo is not None:
        tot = s.profilo["totale"]
        p_tot = potenza_totale(par)
        resa = F.resa_specifica(tot, p_tot)
        m1, m2, m3 = st.columns(3)
        m1.metric("Produzione annua", num(tot.sum() / 1000, 0, "MWh"))
        m2.metric("Potenza nominale totale", num(p_tot, 0, "kW"))
        m3.metric("Ore equivalenti", num(resa, 0, "h/anno"))
        if not 500 <= resa <= 4500:
            st.warning("Le ore equivalenti sembrano fuori scala: controlla che la potenza indicata corrisponda "
                       "al profilo caricato (unità in kW, non W).")
        st.plotly_chart(G.fig_profilo_fonti(s.profilo["componenti"]), width="stretch")


def tab_tecnici(s):
    st.caption("Corrisponde alla sezione *Dati tecnici* dell'INPUT.xlsx.")
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("**Elettrolizzatore**")
        campo("min_elet", "Minimo tecnico [% della potenza nominale]", perc=True, massimo=100, passo=5.0,
              help="Excel: % Electrolyzer power cutoff")
        campo("tassoDEN", "Degrado annuo della produzione [%]", perc=True, massimo=100, passo=0.1,
              help="Excel: tasso di decrescita dell'efficienza nominale annua")
        campo("bar", "Pressione di stoccaggio [bar]", minimo=30, passo=10.0,
              help="Determina la potenza del compressore")
    with c2:
        st.markdown("**Batteria**")
        si_no("batteria", "Batteria nel sistema", help="Excel: Presence of battery in the system")
        dis = s.par["batteria"] != "SI"
        campo("min_batt", "SoC minimo [%]", perc=True, massimo=100, passo=5.0, disabled=dis,
              help="Excel: Cutoff SoC")
        campo("max_batteria", "Energia massima erogabile in 1 h [% della capacità]", perc=True, massimo=100,
              passo=5.0, disabled=dis, help="Excel: Maximum energy % deliverable in 1h")
        campo("eff_batt", "Efficienza di carica e scarica [%]", perc=True, minimo=1, massimo=100, passo=1.0,
              disabled=dis)
    with c3:
        st.markdown("**Ricerca delle taglie** (parametri facoltativi)")
        campo("dP_el", "Elettrolizzatore massimo [× potenza rinnovabile]", minimo=0.05, passo=0.1, fmt="%.2f",
              help="Excel: Electrolyzer range limit. Si esplorano 10-100 taglie da potenza/granularità "
                   "fino a potenza × questo valore.")
        campo("dP_bat", "Batteria massima [× taglia elettrolizzatore]", minimo=0.0, passo=0.1, fmt="%.2f",
              disabled=s.par["batteria"] != "SI", help="Excel: Battery range limit")
        n = M.conta_configurazioni(potenza_totale(s.par), s.par["dP_el"], s.par["batteria"], s.par["dP_bat"])
        st.metric("Configurazioni da esplorare", num(n))


def tab_investimenti(s):
    st.caption("Corrisponde alla sezione *Investimenti* dell'INPUT.xlsx. Valori netti IVA.")
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("**Costi unitari**")
        campo("ImpPV1eurokW", "Fotovoltaico [€/kWp]", passo=50.0)
        if s.par["modo_fonti"] == "API" and s.par["p_Wind"] > 0:
            campo("ImpWindeurokW", "Eolico [€/kW]", passo=50.0)
        if s.par["usa_extra"]:
            campo("ImpExtraeurokW", "Fonte extra [€/kW]", passo=50.0)
        campo("EletteuroKW", "Elettrolizzatore [€/kW]", passo=50.0)
        campo("CompreuroKW", "Compressore [€/kW]", passo=100.0)
        campo("AccuEeurokW", "Batterie [€/kWh]", passo=10.0, disabled=s.par["batteria"] != "SI")
    with c2:
        st.markdown("**Stoccaggio idrogeno**")
        campo("costounitariostoccaggio", "Stoccaggio [€/kg di capacità]", passo=50.0,
              help="Nell'originale il default era 0 (stoccaggio gratuito): qui 1.200 €/kg.")
        campo("idrogstocperc", "Capacità di stoccaggio [% della produzione annua]", perc=True, massimo=100,
              passo=1.0)
        campo("BombSto", "Bombole di stoccaggio [€]", passo=1000.0)
        campo("StazzRif", "Stazione di rifornimento [€]", passo=10000.0,
              help="Default originale 500.000 €: azzeralo se l'impianto non prevede la stazione.")
    with c3:
        st.markdown("**Opere e altri costi**")
        campo("Terr", "Terreno [€]", passo=1000.0, help="0 se il terreno è di proprietà")
        campo("OpeE", "Opere edili [€]", passo=1000.0)
        campo("SpeTOpere", "Spese tecniche [€]", passo=1000.0)
        campo("LavoImp", "Lavori impiantistici [€]", passo=1000.0)
        campo("CarrEll", "Carrello elevatore [€]", passo=1000.0)


def tab_costi_operativi(s):
    st.caption("Corrisponde alla sezione *Costi operativi* (dati economici 1) dell'INPUT.xlsx.")
    c1, c2 = st.columns(2)
    with c1:
        campo("costlitroacqua", "Costo dell'acqua [€/litro]", passo=0.005, fmt="%.3f")
        campo("PercEserImp", "Esercizio impianti [% dell'investimento impianti]", perc=True, passo=0.1)
        campo("Percentimpianti", "Manutenzione impianti [% dell'investimento impianti]", perc=True, passo=0.05)
        campo("PercentOpeEd", "Manutenzione opere edili [% dell'investimento opere]", perc=True, passo=0.01)
    with c2:
        campo("SpesAmmGen", "Spese amministrative e generali [€/anno]", passo=500.0)
        campo("Affitto", "Affitti passivi [€/anno]", passo=500.0)
        campo("CostiPersonal", "Costo del personale [€/anno]", passo=1000.0)
        campo("AltriCost", "Altri costi [€/anno]", passo=500.0)
        si_no("IVAsualtriCost", "IVA sugli altri costi")


def tab_finanziari(s):
    st.caption("Corrisponde alla sezione *Dati economico-finanziari* (dati economici 2) dell'INPUT.xlsx.")
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("**Ricavi e piano**")
        campo("prezzoindrogeno", "Prezzo di vendita dell'idrogeno [€/kg]", passo=0.5)
        campo("prezzoElett", "Prezzo dell'energia immessa in rete [€/kWh]", passo=0.01, fmt="%.3f",
              help="Nell'originale il default era 1 €/kWh (valore segnaposto).")
        campo("incentpubb", "Incentivo pubblico [€/kg di H2]", passo=0.5)
        campo("duratincentpubb", "Durata incentivo [anni]", massimo=50)
        campo("DurPianEcon", "Durata del piano economico [anni]", minimo=3, massimo=40)
        campo("tassoVAN", "Tasso di attualizzazione (VAN, LCOH) [%]", perc=True, passo=0.5)
    with c2:
        st.markdown("**Inflazione e imposte**")
        campo("inflazione", "Inflazione dei costi [%]", perc=True, minimo=-10, passo=0.5)
        campo("inflazionePrezzoElet", "Inflazione prezzo energia [%]", perc=True, minimo=-10, passo=0.5)
        campo("inflazioneIdrog", "Inflazione prezzo idrogeno [%]", perc=True, minimo=-10, passo=0.5)
        campo("aliquoMedia", "Aliquota media sugli utili [%]", perc=True, passo=0.5)
        campo("MaxInterssDed", "Interessi deducibili [% dell'EBITDA]", perc=True, passo=5.0)
        campo("Perciva", "Aliquota IVA [%]", perc=True, passo=1.0)
    with c3:
        st.markdown("**Fonti di finanziamento**")
        campo("ContrPubb", "Contributo pubblico in conto capitale [%]", perc=True, massimo=100, passo=5.0)
        campo("DebitoSenior", "Debito senior [% dell'investimento]", perc=True, massimo=100, passo=5.0)
        cp = 1 - s.par["ContrPubb"] - s.par["DebitoSenior"]
        st.caption(f"Capitale proprio: **{cp * 100:.0f}%**")
        if cp < 0:
            st.error("Contributo + debito superano il 100% dell'investimento.")
        campo("DurDebitoSenior", "Durata debito [anni]", massimo=40)
        campo("tassoDebito", "Tasso del debito [%]", perc=True, passo=0.25)
        campo("FreqPagamenti", "Rate per anno", minimo=1, massimo=12)
        campo("DurataPonte", "Durata prestito ponte IVA [anni] (0 = nessuno)", massimo=10)
        campo("tassoPonte", "Tasso prestito ponte [%]", perc=True, passo=0.25)


def tab_output(s):
    par = s.par
    st.caption("Corrisponde alle opzioni in fondo all'INPUT.xlsx: criterio di ottimizzazione, grafici, simulazioni.")
    c1, c2, c3 = st.columns(3)
    with c1:
        scelta("attributo", "Criterio con cui scegliere i migliori progetti", CRITERI, fmt=nome_var,
               help="Come nell'originale: per costi, prezzo, payback e spegnimenti vince il valore più basso, "
                    "per le altre grandezze il più alto. 'Prezzo idrogeno' calcola per ogni configurazione "
                    "il prezzo minimo che azzera il VAN (più lento).")
        campo("n_progetti", "Quanti progetti mostrare", minimo=1, massimo=20)
        scelta("lingua", "Lingua del file Excel di output", ["ITA", "ENG"])
    with c2:
        si_no("si_fa_simulazione", "Calcola il prezzo o l'incentivo di equilibrio")
        if par["si_fa_simulazione"] == "SI":
            scelta("attributo_simulazione", "Cosa cercare", ["prezzo", "incentivo"],
                   fmt=lambda x: {"prezzo": "Prezzo di vendita di equilibrio (VAN = 0)",
                                  "incentivo": "Incentivo pubblico €/kg di equilibrio (VAN = 0)"}[x])
        si_no("si_fa_grafico_SA", "Grafici di sensitivity (batteria × elettrolizzatore)")
        if par["si_fa_grafico_SA"] == "SI":
            k = _init("SA_variable_list", list(par["SA_variable_list"]))
            par["SA_variable_list"] = st.multiselect("Grandezze per la sensitivity", VARIABILI_GRAFICI,
                                                     key=k, format_func=nome_var)
            if par["batteria"] != "SI":
                st.caption("La sensitivity richiede la batteria nel sistema.")
    with c3:
        si_no("relazione", "Grafici di relazione tra variabili")
        if par["relazione"] == "SI":
            if "w_relazioni" not in s or "_rel_base" not in s:
                s["_rel_base"] = pd.DataFrame(par["relazioni"] or [{"X": "PotEle", "Y": "VAN"}])
            df = st.data_editor(
                s["_rel_base"], key="w_relazioni", num_rows="dynamic", width="stretch", hide_index=True,
                column_config={
                    "X": st.column_config.SelectboxColumn("Variabile X", options=VARIABILI_GRAFICI, required=True),
                    "Y": st.column_config.SelectboxColumn("Variabile Y", options=VARIABILI_GRAFICI, required=True)})
            par["relazioni"] = df.dropna().to_dict("records")
            st.caption("Codici: " + ", ".join(f"{k} = {nome_var(k, False)}" for k in ("VAN", "TIR", "LCOH",
                                                                                         "PotEle", "AccuE")))


def step1(s):
    sezione_import(s)
    schede = st.tabs(["① Fonti energetiche", "② Dati tecnici", "③ Investimenti", "④ Costi operativi",
                      "⑤ Dati economico-finanziari", "⑥ Opzioni di output"])
    with schede[0]:
        tab_fonti(s)
    with schede[1]:
        tab_tecnici(s)
    with schede[2]:
        tab_investimenti(s)
    with schede[3]:
        tab_costi_operativi(s)
    with schede[4]:
        tab_finanziari(s)
    with schede[5]:
        tab_output(s)

    st.divider()
    pronto = s.profilo is not None and s.profilo["firma"] == firma_profilo(s.par, s)
    c1, c2 = st.columns([3, 1])
    with c1:
        if not pronto:
            st.caption("Per proseguire carica (o aggiorna) il profilo orario nella scheda *Fonti energetiche*.")
    with c2:
        st.button("Avanti: elaborazione →", type="primary", width="stretch", disabled=not pronto,
                  on_click=vai, args=(2,))


# =============================================================================================
# STEP 2 - ELABORAZIONE
# =============================================================================================
def controlli(par, profilo):
    errori, avvisi = [], []
    if potenza_totale(par) <= 0:
        errori.append("La potenza rinnovabile totale è zero.")
    if 1 - par["ContrPubb"] - par["DebitoSenior"] < -1e-9:
        errori.append("Contributo pubblico + debito superano il 100% dell'investimento.")
    if par["batteria"] == "SI" and par["eff_batt"] <= 0:
        errori.append("L'efficienza della batteria deve essere maggiore di zero.")
    if par["DurDebitoSenior"] >= par["DurPianEcon"]:
        avvisi.append("Il debito dura quanto o più del piano economico: le rate oltre l'ultimo anno del piano "
                      "non vengono considerate (l'originale in questo caso si bloccava).")
    if par["StazzRif"] > 0:
        avvisi.append(f"È inclusa una stazione di rifornimento da {eur(par['StazzRif'])} (default originale).")
    if par["ContrPubb"] > 0 or par["incentpubb"] > 0:
        avvisi.append(f"Sono attivi contributi pubblici: {par['ContrPubb'] * 100:.0f}% sull'investimento, "
                      f"{par['incentpubb']:.2f} €/kg per {par['duratincentpubb']} anni.")
    if par["si_fa_grafico_SA"] == "SI" and par["batteria"] != "SI":
        avvisi.append("I grafici di sensitivity richiedono la batteria: verranno saltati.")
    if par["attributo"] == "prezzoindrogeno" and (par["relazione"] == "SI" or par["si_fa_simulazione"] == "SI"):
        avvisi.append("Con il criterio 'Prezzo idrogeno' l'originale non produce grafici di relazione né "
                      "simulazioni di equilibrio: verranno saltati.")
    return errori, avvisi


def step2(s):
    par = s.par
    n = M.conta_configurazioni(potenza_totale(par), par["dP_el"], par["batteria"], par["dP_bat"])
    per_conf = 0.00012 + 0.0006
    if par["attributo"] == "prezzoindrogeno":
        per_conf *= 10
    stima = n * per_conf + (2 if M.NUMBA_DISPONIBILE else 0)
    if not M.NUMBA_DISPONIBILE and par["batteria"] == "SI":
        stima += n * 0.013

    st.subheader("Riepilogo")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Potenza rinnovabile", num(potenza_totale(par), 0, "kW"))
    c2.metric("Produzione annua", num(s.profilo["totale"].sum() / 1000, 0, "MWh"))
    c3.metric("Configurazioni", num(n))
    c4.metric("Tempo stimato", f"~{max(1, stima):.0f} s")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Batteria", SI_NO[par["batteria"]])
    c2.metric("Criterio", nome_var(par["attributo"], False))
    c3.metric("Prezzo H2", num(par["prezzoindrogeno"], 2, "€/kg"))
    c4.metric("Capitale proprio / debito / contributo",
              f"{(1 - par['ContrPubb'] - par['DebitoSenior']) * 100:.0f} / {par['DebitoSenior'] * 100:.0f} / "
              f"{par['ContrPubb'] * 100:.0f} %")

    with st.expander("Tutti i parametri che verranno usati"):
        righe = [{"Parametro": k, "Valore": str(v)} for k, v in parametri_motore(par).items()
                 if k not in ("file_csv", "tipo_file")]
        st.dataframe(pd.DataFrame(righe), hide_index=True, width="stretch", height=300)

    errori, avvisi = controlli(par, s.profilo)
    for e in errori:
        st.error(e)
    for a in avvisi:
        st.warning(a)

    st.divider()
    c1, c2, c3 = st.columns([1, 2, 1])
    c1.button("← Torna ai parametri", on_click=vai, args=(1,), width="stretch")
    avvia = c3.button("🚀 Avvia elaborazione", type="primary", width="stretch", disabled=bool(errori))
    if s.risultato is not None:
        c2.button("Vai agli ultimi risultati →", on_click=vai, args=(3,), width="stretch")
    if avvia:
        barra = st.progress(0.0, text="Avvio...")

        def progresso(f, testo):
            barra.progress(f, text=f"{testo} - {f * 100:.0f}%")

        try:
            c = M.Analisi_combinata(parametri_motore(par), s.profilo["totale"], progresso=progresso)
            c.componenti_profilo = s.profilo["componenti"]
            s.risultato = c
            s.par_usati = copy.deepcopy(par)
            s.excel = None
            s.config_sel = None
            s.pop("w_progetto", None)
            s.step = 3
            st.rerun()
        except Exception as e:
            barra.empty()
            st.error(f"L'elaborazione si è interrotta: {e}")
            with st.expander("Dettagli tecnici"):
                st.code(traceback.format_exc())


# =============================================================================================
# STEP 3 - RISULTATI
# =============================================================================================
def tabella_numeri(df, colonne_anni=False):
    out = df.copy()
    out = out.apply(pd.to_numeric, errors="coerce")
    if colonne_anni:
        out.columns = [f"Anno {c}" for c in out.columns]
    return out


@st.cache_data(show_spinner=False, max_entries=8)
def _dati_config(_c, idx, chiave):
    a = _c._analisi_finanziaria(idx)
    a.RUN(tabelle=True)
    return a, _c.analisi1.andamenti_progetto(idx)


def progetto_scelto(s, c):
    """Elenco dei progetti consultabili: i migliori N + l'eventuale configurazione cliccata nel grafico."""
    opzioni = list(range(len(c.top_progetti)))
    etichette = {i: f"Progetto {i + 1}" + (" (migliore)" if i == 0 else "") for i in opzioni}
    if s.config_sel is not None and s.config_sel not in c.top_idx:
        opzioni.append("sel")
        etichette["sel"] = f"Configurazione {s.config_sel} (scelta dal grafico)"
    if "_progetto_richiesto" in s:          # richiesta arrivata dal grafico (prima che il widget esista)
        s["w_progetto"] = s.pop("_progetto_richiesto")
    if s.get("w_progetto") not in opzioni:
        s.pop("w_progetto", None)
    scelta_p = st.selectbox("Progetto da esaminare", opzioni, format_func=lambda x: etichette[x], key="w_progetto")
    if scelta_p == "sel":
        a, andam = _dati_config(c, s.config_sel, id(c))
        return a, andam, s.config_sel
    return c.top_progetti[scelta_p], c.top_andamenti[scelta_p], c.top_idx[scelta_p]


def kpi(a, c):
    r1 = st.columns(4)
    r1[0].metric("VAN", eur(a.VAN), help="Valore attuale netto dei flussi di cassa al tasso indicato")
    r1[1].metric("TIR", num(a.TIR * 100, 1, "%") if a.TIR is not None else "n.d.")
    r1[2].metric("Payback", f"{a.PAYBACK} anni" if a.PAYBACK is not None else "oltre il piano")
    r1[3].metric("LCOH attualizzato", num(a.LCOH, 2, "€/kg"),
                 help="(Investimento + costi operativi attualizzati) / produzione attualizzata")
    r2 = st.columns(4)
    r2[0].metric("Elettrolizzatore", num(a.PotEle, 0, "kW"))
    r2[1].metric("Batteria", num(a.AccuE, 0, "kWh"))
    r2[2].metric("Idrogeno (1° anno)", num(a.ProdAnnuaIdrogkg, 0, "kg"))
    r2[3].metric("Capacity factor", num(a.CapFac, 1, "%"))
    r3 = st.columns(4)
    r3[0].metric("Investimento (netto IVA)", eur(a.investimento))
    r3[1].metric("Compressore", num(a.potenza_compressore, 1, "kW"))
    r3[2].metric("Stoccaggio", num(a.idrogeno_stocc, 0, "kg"))
    r3[3].metric("Spegnimenti / anno", num(a.spegn_giorn, 0))


def scheda_migliori(c):
    st.markdown(f"Configurazioni esplorate: **{c.analisi1.qt_progetti}** · criterio: "
                f"**{nome_var(c.attributo, False)}**")
    tab = c.TabellaMax.apply(pd.to_numeric, errors="coerce").astype(float)
    riga_tir = [i for i in tab.index if str(i).startswith(("TIR", "Project IRR"))]
    if riga_tir:
        tab.loc[riga_tir[0]] = tab.loc[riga_tir[0]] * 100
        tab = tab.rename(index={riga_tir[0]: riga_tir[0] + " [%]"})
    testo = tab.map(lambda v: "n.d." if pd.isna(v) else (num(v, 2) if abs(v) < 100 else num(v, 0)))
    st.dataframe(testo, width="stretch")
    t = c.top_righe.copy()
    t["Progetto"] = [f"P{i + 1}" for i in range(len(t))]
    col1, col2 = st.columns(2)
    import plotly.graph_objects as go
    for col, var in ((col1, "VAN"), (col2, "LCOH")):
        fig = go.Figure(go.Bar(x=t["Progetto"], y=t[var], marker_color=G.BLU,
                               hovertemplate="%{x}: %{y:,.2f}<extra></extra>"))
        col.plotly_chart(G._stile(fig, nome_var(var), 280, None, legenda=False), width="stretch")


def scheda_energia(a, andam, batteria):
    c1, c2 = st.columns([3, 2])
    c1.plotly_chart(G.fig_energia_mensile(andam, batteria), width="stretch")
    c2.plotly_chart(G.fig_idrogeno_mensile(andam), width="stretch")
    c1, c2 = st.columns([1, 3])
    with c1:
        inizio = st.slider("Primo giorno del periodo", 0, 364, 172, key="g_inizio",
                           help="0 = 1° gennaio; 172 = fine giugno")
        giorni = st.select_slider("Durata", [3, 7, 14, 31, 92, 365], value=7, key="g_durata",
                                  format_func=lambda d: f"{d} giorni")
    with c2:
        st.plotly_chart(G.fig_periodo(andam, batteria, inizio, giorni), width="stretch")
    c1, c2 = st.columns(2)
    c1.plotly_chart(G.fig_heatmap(andam[0], "Potenza all'elettrolizzatore (giorno × ora)", "kW"),
                    width="stretch")
    if batteria and a.AccuE > 0:
        c2.plotly_chart(G.fig_heatmap(andam[6], "Energia in batteria (giorno × ora)", "kWh", "Greens"),
                        width="stretch")
    else:
        c2.plotly_chart(G.fig_heatmap(andam[2], "Energia immessa / non utilizzata (giorno × ora)", "kW", "Oranges"),
                        width="stretch")
    st.plotly_chart(G.fig_durata(andam), width="stretch")
    with st.expander("Dati orari del progetto (CSV)"):
        nomi = ["E_elettrolizzatore_kWh", "H2_kg", "E_immessa_kWh", "Produzione_FER_kWh", "P_min_elett_kW",
                "P_nom_elett_kW", "E_batteria_kWh", "P_batteria_kW", "E_batteria_disponibile_kWh",
                "Energia_disponibile_kWh", "Max_erogabile_kW"]
        df = pd.DataFrame(andam.T, columns=nomi[:andam.shape[0]], index=G.ORE.strftime("%d/%m %H:00"))
        st.download_button("Scarica le 8760 ore", df.to_csv(sep=";", decimal=","), "flussi_orari.csv", "text/csv")


def scheda_economia(a):
    c1, c2 = st.columns([3, 2])
    c1.plotly_chart(G.fig_flussi_cassa(a), width="stretch")
    c2.plotly_chart(G.fig_costo_kg(a), width="stretch")
    st.plotly_chart(G.fig_capex(a), width="stretch")
    fmt = lambda v: num(v, 0)
    st.markdown("##### Conto economico")
    st.dataframe(tabella_numeri(a.dfContoEconomico, True).style.format(fmt), width="stretch")
    st.markdown("##### Flussi di cassa")
    st.dataframe(tabella_numeri(a.dfFlussiMonetari, True).style.format(fmt), width="stretch")
    with st.expander("Indici finanziari (tabella dell'originale)"):
        st.dataframe(a.dfIndiciFinanziari.T.rename(columns={0: "Valore"}), width="stretch")


def scheda_esplora(s, c):
    df = c.risultati.replace([np.inf, -np.inf], np.nan)
    disponibili = [v for v in M.VARIABILI if v in df.columns and df[v].notna().any() and df[v].nunique() > 1]
    st.caption("Ogni punto è una configurazione elettrolizzatore/batteria. La frontiera di Pareto unisce le "
               "configurazioni per cui non ne esiste un'altra migliore su entrambi gli assi. "
               "Clicca un punto per esaminarlo in dettaglio.")
    c1, c2, c3 = st.columns(3)
    x = c1.selectbox("Asse X", disponibili, index=disponibili.index("investimento") if "investimento" in disponibili
                     else 0, format_func=nome_var, key="es_x")
    y = c2.selectbox("Asse Y", disponibili, index=disponibili.index("VAN") if "VAN" in disponibili else 1,
                     format_func=nome_var, key="es_y")
    colore = c3.selectbox("Colore", disponibili, index=disponibili.index("PotEle") if "PotEle" in disponibili
                          else 0, format_func=nome_var, key="es_col")
    c1, c2, c3 = st.columns(3)
    max_x = c1.radio(f"Per {nome_var(x, False)} è meglio", ["alto", "basso"], horizontal=True,
                     index=0 if M.VARIABILI[x][3] == "max" else 1, key=f"es_vx_{x}") == "alto"
    max_y = c2.radio(f"Per {nome_var(y, False)} è meglio", ["alto", "basso"], horizontal=True,
                     index=0 if M.VARIABILI[y][3] == "max" else 1, key=f"es_vy_{y}") == "alto"
    pareto = G.fronte_pareto(df, x, y, max_x, max_y)
    c3.metric("Configurazioni sulla frontiera", len(pareto))
    fig = G.fig_esplora(df, x, y, colore, nome_var(x), nome_var(y), nome_var(colore), pareto, c.top_idx,
                        s.config_sel)
    ev = st.plotly_chart(fig, width="stretch", on_select="rerun", selection_mode="points", key="es_graf")
    try:
        punti = ev.selection.points if ev else []
    except AttributeError:
        punti = []
    for p in punti:
        cd = p.get("customdata")
        if isinstance(cd, (list, tuple)):
            cd = cd[0] if cd else None
        if cd is not None and int(cd) != s.config_sel:
            s.config_sel = int(cd)
            s["_progetto_richiesto"] = "sel" if s.config_sel not in c.top_idx else c.top_idx.index(s.config_sel)
            st.rerun()
    if s.config_sel is not None:
        r = df[df["ID"] == s.config_sel].iloc[0]
        st.info(f"Configurazione {s.config_sel}: elettrolizzatore {num(r['PotEle'], 0, 'kW')}, batteria "
                f"{num(r['AccuE'], 0, 'kWh')}, VAN {eur(r['VAN'])}, LCOH {num(r['LCOH'], 2, '€/kg')}. "
                "È selezionata in alto in *Progetto da esaminare*.")
    with st.expander("Tabella di tutte le configurazioni"):
        st.dataframe(df, width="stretch", hide_index=True, height=350)
        st.download_button("Scarica CSV", df.to_csv(index=False, sep=";", decimal=","), "configurazioni.csv",
                           "text/csv")


def scheda_relazioni(c):
    import plotly.graph_objects as go
    if c.relazione in ("SI", "YES") and c.attributo not in M.CRITERI_PREZZO:
        st.markdown("##### Relazioni tra variabili")
        coppie = list(zip(c.variable_1_list, c.variable_2_list))
        for i in range(0, len(coppie), 2):
            cols = st.columns(2)
            for col, (vx, vy) in zip(cols, coppie[i:i + 2]):
                xs, ys = c.crea_grafico_relazione(vx, vy)
                fig = go.Figure(go.Scatter(x=xs, y=ys, mode="markers", marker=dict(size=6, color=G.ARANCIO),
                                           hovertemplate="%{x:,.2f} → %{y:,.2f}<extra></extra>"))
                fig = G._stile(fig, f"{nome_var(vx, False)} vs {nome_var(vy, False)}", 340, nome_var(vy), False)
                fig.update_xaxes(title_text=nome_var(vx))
                fig.update_layout(hovermode="closest")
                col.plotly_chart(fig, width="stretch")
    if c.si_fa_grafico_SA in ("SI", "YES") and c.batteria == "SI":
        st.markdown("##### Sensitivity: batteria × elettrolizzatore")
        for v in c.SA_variable_list:
            st.plotly_chart(G.fig_sensitivity(c.risultati, v, nome_var(v)), width="stretch")


def scheda_equilibrio(c):
    cosa = "Incentivo di equilibrio [€/kg]" if c.attributo_simulazione in M.SIM_INCENTIVO else \
        "Prezzo di equilibrio [€/kg]"
    df = pd.DataFrame({"Progetto": [f"Progetto {i + 1}" for i in range(len(c.prezzi_eq))],
                       "Elettrolizzatore [kW]": [num(p.PotEle, 0) for p in c.top_progetti],
                       "Batteria [kWh]": [num(p.AccuE, 0) for p in c.top_progetti],
                       cosa: [num(v, 2) if v is not None else "n.d." for v in c.prezzi_eq],
                       "VAN all'equilibrio [€]": [eur(v) if v is not None else "n.d." for v in c.VAN_eq]})
    st.caption("Valore minimo (a passi di 0,5 €/kg, come nell'originale) che rende il VAN positivo. "
               f"'n.d.' = non raggiunto entro {M.MASSIMO_EQUILIBRIO:.0f} €/kg.")
    st.dataframe(df, hide_index=True, width="stretch")


def step3(s):
    c = s.risultato
    if c is None:
        st.info("Nessun risultato: esegui prima l'elaborazione.")
        st.button("← Vai all'elaborazione", on_click=vai, args=(2,))
        return
    if s.get("par_usati") != s.par:
        st.warning("Hai modificato i parametri dopo l'elaborazione: i risultati si riferiscono ai valori "
                   "precedenti. Torna allo step 2 per ricalcolare.")

    c_sel, c_xl1, c_xl2, c_nav = st.columns([3, 1.2, 1.2, 1.2])
    with c_sel:
        a, andam, idx = progetto_scelto(s, c)
    with c_xl1:
        st.write("")
        st.write("")
        if s.excel is None:
            if st.button("📊 Prepara Excel di output", width="stretch",
                         help="Stesso formato dell'OUTPUT.xlsx originale. Con molte configurazioni e i grafici "
                              "di sensitivity può richiedere fino a un minuto."):
                with st.spinner("Scrittura del file Excel..."):
                    s.excel = c.scarica_risultati()
                st.rerun()
        else:
            st.download_button("⬇️ Scarica OUTPUT.xlsx", s.excel, "H2FAST_OUTPUT.xlsx",
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               width="stretch", type="primary")
    with c_xl2:
        st.write("")
        st.write("")
        st.download_button("⬇️ Parametri (.json)", json.dumps(s.get("par_usati", s.par), indent=2, default=str),
                           "h2fast_parametri.json", "application/json", width="stretch")
    with c_nav:
        st.write("")
        st.write("")
        st.button("← Modifica parametri", on_click=vai, args=(1,), width="stretch")

    kpi(a, c)
    batteria = c.batteria == "SI"
    nomi = ["🏆 Migliori progetti", "⚡ Energia", "💶 Economia", "🎯 Esplora configurazioni"]
    mostra_rel = (c.relazione in ("SI", "YES") or (c.si_fa_grafico_SA in ("SI", "YES") and batteria)) \
        and c.attributo not in M.CRITERI_PREZZO
    mostra_eq = bool(c.prezzi_eq)
    if mostra_rel:
        nomi.append("📈 Relazioni e sensitivity")
    if mostra_eq:
        nomi.append("⚖️ Equilibrio")
    schede = st.tabs(nomi)
    with schede[0]:
        scheda_migliori(c)
    with schede[1]:
        scheda_energia(a, andam, batteria)
    with schede[2]:
        scheda_economia(a)
    with schede[3]:
        scheda_esplora(s, c)
    k = 4
    if mostra_rel:
        with schede[k]:
            scheda_relazioni(c)
        k += 1
    if mostra_eq:
        with schede[k]:
            scheda_equilibrio(c)


# =============================================================================================
# MAIN
# =============================================================================================
def main():
    s = stato()
    comune = accesso_h2ready()
    st.title("⚡ H2 FAST · Progettazione di impianti H2")
    st.caption("Produzione da rinnovabili, elettrolizzatore, batteria e stoccaggio: dimensionamento e business "
               "plan. Motore H2FAsT sviluppato nel progetto AMETHYST (Interreg Alpine Space).")
    if comune is not None and len(comune) and not s.get("_cerca_init"):
        s["_cerca_init"] = True
        nome = str(comune.get("NOME_COMUNE", "") or "")
        if nome:
            s["cerca_loc"] = nome
    stepper(s)
    if s.step == 1:
        step1(s)
    elif s.step == 2:
        step2(s)
    else:
        step3(s)


main()
