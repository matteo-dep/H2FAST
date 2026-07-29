import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import os
import requests
from geopy.geocoders import Nominatim
import folium
from streamlit_folium import st_folium

# Importiamo LE TUE classi originali e intatte
from motore_h2fast import Analisi_tecnica, Analisi_finanziaria

st.set_page_config(page_title="H2FAsT Simulator", layout="wide")
st.title("Simulatore H2FAsT - Dashboard Interattiva 🏭")

# Leggiamo il testo direttamente dal file README.md esterno!
try:
    with open("README.md", "r", encoding="utf-8") as file:
        testo_guida = file.read()
        
    with st.expander("ℹ️ Clicca qui per leggere la guida completa e scoprire come funziona"):
        st.markdown(testo_guida)
except FileNotFoundError:
    st.info("Crea un file README.md nella stessa cartella per visualizzare qui la guida.")

st.divider()

# ==========================================
# FUNZIONE EOLICA
# ==========================================
def calcola_potenza_eolica_kw(v_vento_ms, p_nominale_kw):
    if p_nominale_kw == 0: return 0.0
    p_nominale_w = p_nominale_kw * 1000.0
    v_cut_in = 3.0    
    v_rated = 12.0    
    v_cut_out = 25.0  
    
    if v_vento_ms < v_cut_in or v_vento_ms > v_cut_out:
        return 0.0
    elif v_cut_in <= v_vento_ms < v_rated:
        return (p_nominale_w * ((v_vento_ms**3 - v_cut_in**3) / (v_rated**3 - v_cut_in**3))) / 1000.0
    else:
        return p_nominale_kw

# Inizializziamo le coordinate nella memoria
if "lat_s" not in st.session_state: st.session_state.update({"lat_s": 41.89, "lon_s": 12.49, "lat_w": 41.89, "lon_w": 12.49})

# ==========================================
# STEP 0: SORGENTI DATI
# ==========================================
st.header("🌍 Step 0: Sorgenti Dati Energetici")
metodo_dati = st.radio("Come vuoi inserire i dati Base (Sole/Vento)?", ["Mappa Interattiva (API)", "Carica file CSV (Manuale)"], horizontal=True)

file_csv = None
tipo_file_passato = "SI"
usa_extra = False
file_extra = None
p_Extra = 0.0

if metodo_dati == "Mappa Interattiva (API)":
    luoghi_separati = st.toggle("📍 Usa due luoghi geografici differenti per Fotovoltaico ed Eolico")
    
    if not luoghi_separati:
        st.markdown("**Seleziona la località per l'Impianto Ibrido (Sole + Vento)**")
        m = folium.Map(location=[st.session_state.lat_s, st.session_state.lon_s], zoom_start=5)
        map_data = st_folium(m, height=350, width=700, key="mappa_unica")
        if map_data and map_data.get("last_clicked"):
            st.session_state.lat_s = map_data["last_clicked"]["lat"]
            st.session_state.lon_s = map_data["last_clicked"]["lng"]
            st.session_state.lat_w = st.session_state.lat_s
            st.session_state.lon_w = st.session_state.lon_s
        st.success(f"✅ Coordinate Selezionate: Lat {st.session_state.lat_s:.4f}, Lon {st.session_state.lon_s:.4f}")
    
    else:
        col_m1, col_m2 = st.columns(2)
        with col_m1:
            st.markdown("☀️ **Località Fotovoltaico**")
            m_s = folium.Map(location=[st.session_state.lat_s, st.session_state.lon_s], zoom_start=4)
            map_sun = st_folium(m_s, height=300, key="mappa_sole")
            if map_sun and map_sun.get("last_clicked"):
                st.session_state.lat_s = map_sun["last_clicked"]["lat"]
                st.session_state.lon_s = map_sun["last_clicked"]["lng"]
            st.info(f"Lat: {st.session_state.lat_s:.4f} | Lon: {st.session_state.lon_s:.4f}")
            
        with col_m2:
            st.markdown("💨 **Località Eolico**")
            m_w = folium.Map(location=[st.session_state.lat_w, st.session_state.lon_w], zoom_start=4)
            map_wind = st_folium(m_w, height=300, key="mappa_vento")
            if map_wind and map_wind.get("last_clicked"):
                st.session_state.lat_w = map_wind["last_clicked"]["lat"]
                st.session_state.lon_w = map_wind["last_clicked"]["lng"]
            st.info(f"Lat: {st.session_state.lat_w:.4f} | Lon: {st.session_state.lon_w:.4f}")
else:
    file_csv = st.file_uploader("Carica il tuo file Base di serie oraria CSV", type=["csv"])
    tipo_file_passato = st.selectbox("Il file proviene da PVGIS?", ["SI", "NO"])

with st.expander("➕ Integrazione Fonti Esterne (Opzionale)"):
    usa_extra = st.checkbox("Aggiungi un profilo utente (es. Idroelettrico, Biomasse)")
    if usa_extra:
        col_ex1, col_ex2 = st.columns(2)
        with col_ex1: file_extra = st.file_uploader("Carica file CSV (8760 valori orari in kW)", type=["csv", "txt"])
        with col_ex2: p_Extra = st.number_input("Potenza Nominale Extra (kW)", value=100.0)

st.divider()

# ==========================================
# STEP 1: PARAMETRI TECNICI
# ==========================================
st.header("⚙️ Step 1: Parametri Tecnici e Dimensionamento")
col_t1, col_t2, col_t3 = st.columns(3)

with col_t1:
    p_PV = st.number_input("Fotovoltaico (kWp)", value=1000.0)
    p_Wind = st.number_input("Eolico (kW)", value=0.0)
    lingua = st.selectbox("Lingua output", ["ITA", "ENG"])

with col_t2:
    dP_el = st.number_input("Electrolyzer range limit", value=1.0)
    min_elet = st.number_input("% Electrolyzer cutoff", value=0.2)
    n_progetti = st.number_input("Top N Progetti da salvare", value=5, min_value=1)

with col_t3:
    batteria = st.selectbox("Presenza Batteria", ["SI", "NO"])
    dP_bat = st.number_input("Battery range limit", value=2.0) if batteria == "SI" else 0.0
    min_batt = st.number_input("Cutoff SoC [%]", value=0.2)
    max_batteria = st.number_input("Max energy deliverable in 1h (%)", value=0.5)
    eff_batt = st.number_input("Efficienza Batteria", value=0.95)

st.divider()

# ==========================================
# STEP 2: CAPEX E COSTI
# ==========================================
st.header("💶 Step 2: CAPEX & Costi Impianto")
col_c1, col_c2, col_c3 = st.columns(3)

with col_c1:
    ImpPV1eurokW = st.number_input("Fotovoltaico [€/kWp]", value=800.0)
    ImpWind1eurokW = st.number_input("Eolico [€/kW]", value=1000.0)
    ImpExtra1eurokW = st.number_input("Fonte Extra [€/kW]", value=2000.0)

with col_c2:
    EletteuroKW = st.number_input("Elettrolizzatore [€/kW]", value=1650.0)
    CompreuroKW = st.number_input("Compressore [€/kW]", value=4000.0)
    AccuEeurokW = st.number_input("Batterie [€/kWh]", value=200.0)

with col_c3:
    Terr = st.number_input("Terreno/Land [€]", value=0.0)
    OpeE = st.number_input("Opere Edili [€]", value=0.0)
    StazzRif = st.number_input("Stazione Rifornimento [€]", value=0.0)
    costounitariostoccaggio = st.number_input("Stoccaggio [€/kg]", value=1200.0)

st.divider()

# ==========================================
# STEP 3: PARAMETRI FINANZIARI
# ==========================================
st.header("📊 Step 3: Parametri Finanziari")
col_f1, col_f2, col_f3 = st.columns(3)

with col_f1: prezzoindrogeno = st.number_input("Prezzo vendita Idrogeno [€/kg]", value=20.0)
with col_f2: tassoVAN = st.number_input("Tasso di Sconto (Discount rate)", value=0.07)
with col_f3: 
    inflazione = st.number_input("Tasso di Inflazione", value=0.01)
    DurPianEcon = st.number_input("Durata Piano Economico [anni]", value=20, step=1)

st.divider()

# ==========================================
# ESECUZIONE
# ==========================================
st.header("🚀 Step 4: Esecuzione Simulazione")
if st.button("AVVIA OTTIMIZZAZIONE", use_container_width=True, type="primary"):
    temp_filename = "temp_input_orario"

    if metodo_dati == "Mappa Interattiva (API)":
        with st.spinner('📡 Acquisizione dati solari ed eolici in corso...'):
            try:
                url_pv = f"https://re.jrc.ec.europa.eu/api/v5_2/seriescalc?lat={st.session_state.lat_s}&lon={st.session_state.lon_s}&startyear=2019&endyear=2019&pvcalculation=1&peakpower={p_PV}&loss=14&outputformat=json"
                df_pv = pd.DataFrame(requests.get(url_pv).json()['outputs']['hourly'])
                p_solare_kw = (df_pv['P'] / 1000.0).values
                
                url_wind = "https://archive-api.open-meteo.com/v1/archive"
                params_wind = {"latitude": st.session_state.lat_w, "longitude": st.session_state.lon_w, "start_date": "2019-01-01", "end_date": "2019-12-31", "hourly": "windspeed_100m", "wind_speed_unit": "ms", "timezone": "UTC"}
                v_vento = np.array(requests.get(url_wind, params=params_wind).json()['hourly']['windspeed_100m'])
                p_eolico_kw = np.array([calcola_potenza_eolica_kw(v, p_Wind) for v in v_vento])
                
                min_len = min(len(p_solare_kw), len(p_eolico_kw))
                df_export = pd.DataFrame({'P_kW': p_solare_kw[:min_len] + p_eolico_kw[:min_len]})
                df_export.to_csv(temp_filename + ".csv", index=False, header=['P_kW'], sep=';', decimal=',')
                tipo_file_passato = "NO"
            except Exception as e:
                st.error(f"❌ Errore API: {e}")
                st.stop()
    else:
        if file_csv is None: st.stop()
        with open(temp_filename + ".csv", "wb") as f: f.write(file_csv.getbuffer())

    with st.spinner('Calcolo Vettoriale in corso... L\'ottimizzazione è attiva.'):
        potenza_nominale_totale = p_PV + p_Wind + p_Extra
        analisi1 = Analisi_tecnica(
            file_csv=temp_filename, tipo_file=tipo_file_passato, p_PV=potenza_nominale_totale, dP_el=dP_el, 
            batteria=batteria, dP_bat=dP_bat, min_batt=min_batt, max_batteria=max_batteria, lingua=lingua, eff_batt=eff_batt, min_elet=min_elet
        )
        
        if usa_extra and file_extra is not None:
            df_extra = pd.read_csv(file_extra, sep=None, engine='python')
            extra_vals = pd.to_numeric(df_extra.iloc[:, 0].astype(str).str.replace(',', '.'), errors='coerce').fillna(0).values
            min_l = min(len(analisi1.E_PV), len(extra_vals))
            analisi1.E_PV[:min_l] += extra_vals[:min_l]
            analisi1.e_pv = np.sum(analisi1.E_PV)

        analisi1.progress_bar = st.progress(0)
        analisi1.status_text = st.empty()
        analisi1.run_analysis()
        if os.path.exists(temp_filename + ".csv"): os.remove(temp_filename + ".csv")
            
        top_progetti_tuples, dati_scatter = [], []
        
        # AGGIUNTO IL PAYBACK ALLA MEMORIA DEI GRAFICI!
        def salva_memoria(an_fin, andamenti_tecnici, val_batteria, cur_id):
            van = an_fin.VAN if an_fin.VAN is not None else -float('inf')
            lcoh_calc = an_fin.investimento / (an_fin.ProdAnnuaIdrogkg * DurPianEcon) if an_fin.ProdAnnuaIdrogkg > 0 else 0
            dati_scatter.append({
                'ID_Progetto': cur_id, 
                'VAN [€]': van, 
                'TIR [%]': an_fin.TIR, 
                'Payback Time [Anni]': an_fin.PAYBACK, # ECCOLO
                'LCOH Semplificato [€/kg]': lcoh_calc,
                'Investimento CAPEX [€]': an_fin.investimento, 
                'Produzione Idrogeno [kg]': an_fin.ProdAnnuaIdrogkg, 
                'Taglia Elettrolizzatore [kW]': an_fin.PotEle, 
                'Taglia Batteria [kWh]': val_batteria, 
                'Energia Sprecata (Curtailment) [kWh]': an_fin.ProdElettVend, 
                'Capacity Factor [%]': an_fin.CapFac,
            })
            top_progetti_tuples.append((van, an_fin, andamenti_tecnici))
            top_progetti_tuples.sort(key=lambda x: x[0], reverse=True)
            return top_progetti_tuples[:n_progetti]

        prog_bar_fin = st.progress(0)
        status_fin = st.empty()
        tot_progetti = len(analisi1.P_elc) if batteria == "NO" else len(analisi1.potenza_batt)

        for index, el in enumerate(zip(analisi1.potenza_batt if batteria == "SI" else analisi1.CF, analisi1.CF if batteria == "SI" else analisi1.P_elc, analisi1.potenza_elett if batteria == "SI" else analisi1.E_H2, analisi1.E_H2 if batteria == "SI" else analisi1.E_im, analisi1.E_im if batteria == "SI" else analisi1.M_H2, analisi1.M_H2 if batteria == "SI" else analisi1.andamenti, analisi1.andamenti if batteria == "SI" else analisi1.spegn_giorn, analisi1.spegn_giorn if batteria == "SI" else [0]*tot_progetti)):
            if batteria == "NO": el1, el2, el3, el4, el5, el6, el7 = el[:7]
            else: el1, el2, el3, el4, el5, el6, el7, el8 = el
            
            an_fin = Analisi_finanziaria(Terr=Terr, OpeE=OpeE, ImpPV1=p_PV, ImpPV1eurokW=ImpPV1eurokW, EletteuroKW=EletteuroKW, CompreuroKW=CompreuroKW, AccuE=el1 if batteria=="SI" else 0, AccuEeurokW=AccuEeurokW, idrogstocperc=0.1, StazzRif=StazzRif, SpeTOpere=0, BombSto=0, LavoImp=0, CarrEll=0, CapFac=el2 if batteria=="SI" else el1, PotEle=el3 if batteria=="SI" else el2, tassoDEN=0.005, ProdAnnuaIdrogkg=el6 if batteria=="SI" else el5, bar=500, costlitroacqua=0.025, costounitariostoccaggio=costounitariostoccaggio, PercEserImp=0.005, Percentimpianti=0.0025, PercentOpeEd=0.0005, SpesAmmGen=2000, Affitto=0, CostiPersonal=0, AltriCost=0, IVAsualtriCost="NO", DurPianEcon=DurPianEcon, inflazione=inflazione, inflazionePrezzoElet=0.015, inflazioneIdrog=0.01, tassoVAN=tassoVAN, incentpubb=2.0, duratincentpubb=20, prezzoindrogeno=prezzoindrogeno, ProdElettVend=el5 if batteria=="SI" else el4, EnergiaAutocons=el4 if batteria=="SI" else el3, prezzoElett=0.1, ContrPubb=0.9, DebitoSenior=0.0, DurDebitoSenior=10, tassoDebito=0.05, FreqPagamenti=1, tassoPonte=0, DurataPonte=0, aliquoMedia=0.275, MaxInterssDed=0.3, lingua=lingua, Perciva=0.22, spegn_giorn=el8 if batteria=="SI" else el7)
            an_fin.investimento += (p_Wind * ImpWind1eurokW) + (p_Extra * ImpExtra1eurokW)
            an_fin.RUN()
            top_progetti_tuples = salva_memoria(an_fin, el7 if batteria=="SI" else el6, el1 if batteria=="SI" else 0, index)
            prog_bar_fin.progress((index + 1) / tot_progetti)
            status_fin.text(f"Analisi Finanziaria: {((index + 1) / tot_progetti)*100:.1f}%")

        prog_bar_fin.empty(); status_fin.empty(); analisi1.progress_bar.empty(); analisi1.status_text.empty()
        
        top_progetti, andamenti_top = [], []
        for item in top_progetti_tuples:
            item[1].costruzione_tabelle()
            top_progetti.append(item[1])
            andamenti_top.append(item[2])

        st.session_state.update({'simulazione_completata': True, 'df_tutti': pd.DataFrame(dati_scatter).dropna(), 'top_progetti': top_progetti, 'andamenti_top': andamenti_top, 'batteria_usata': batteria})

# ==========================================
# DASHBOARD DEI RISULTATI
# ==========================================
if st.session_state.get('simulazione_completata', False):
    st.success("✅ Ottimizzazione Completata!")
    df_tutti = st.session_state['df_tutti']
    top_progetti = st.session_state['top_progetti']
    andamenti_top = st.session_state['andamenti_top']
    batteria_usata = st.session_state['batteria_usata']

    st.divider()
    tab1, tab2, tab3 = st.tabs(["📊 Finanza & Business Plan", "⚡ Flussi Energetici", "🎯 Frontiera di Pareto"])

    with tab1:
        df_sommario = pd.DataFrame({'VAN [€]': [p.VAN for p in top_progetti], 'TIR [%]': [p.TIR for p in top_progetti], 'PAYBACK': [p.PAYBACK for p in top_progetti], 'Elettrolizzatore [kW]': [p.PotEle for p in top_progetti], 'Batteria [kWh]': [p.AccuE for p in top_progetti] if batteria_usata == "SI" else [0]*len(top_progetti), 'H2 [kg/anno]': [p.ProdAnnuaIdrogkg for p in top_progetti]}).T
        df_sommario.columns = [f"Progetto {i+1}" for i in range(len(top_progetti))]
        st.dataframe(df_sommario, use_container_width=True)

        c_eco, c_cas = st.columns(2)
        idx_eco = c_eco.selectbox("Conto Economico:", range(len(top_progetti)), format_func=lambda x: f"Progetto {x+1}")
        c_eco.dataframe(top_progetti[idx_eco].dfContoEconomico, use_container_width=True)
        idx_cas = c_cas.selectbox("Flussi di Cassa:", range(len(top_progetti)), format_func=lambda x: f"Progetto {x+1}")
        c_cas.dataframe(top_progetti[idx_cas].dfFlussiMonetari, use_container_width=True)

    with tab2:
        idx_ene = st.selectbox("Grafico Flussi:", range(len(top_progetti)), format_func=lambda x: f"Progetto {x+1}")
        ore, andamenti = np.arange(8760), andamenti_top[idx_ene]
        fig_flussi = go.Figure()
        fig_flussi.add_trace(go.Scattergl(x=ore, y=andamenti[3], name='Rinnovabili Prodotte', line=dict(color='#92D050')))
        fig_flussi.add_trace(go.Scattergl(x=ore, y=andamenti[2], name='Energia Curtailment', line=dict(color='#FFC000')))
        fig_flussi.add_trace(go.Scattergl(x=ore, y=andamenti[0], name='In Elettrolizzatore', line=dict(color='#00AF50')))
        fig_flussi.add_trace(go.Scattergl(x=ore, y=andamenti[5], name='Taglia Elettrolizzatore', line=dict(color='gray', dash='dash')))
        if batteria_usata == "SI": fig_flussi.add_trace(go.Scattergl(x=ore, y=andamenti[6], name='Livello Batteria', line=dict(color='orange')))
        fig_flussi.update_layout(xaxis=dict(rangeslider=dict(visible=True), type="-"))
        st.plotly_chart(fig_flussi, use_container_width=True)

    with tab3:
        st.markdown("👇 **Clicca su un pallino per vedere tutti i dettagli dell'impianto!**")
        col_x, col_y, col_color = st.columns(3)
        cols = [c for c in df_tutti.columns if c != 'ID_Progetto']
        x_var = col_x.selectbox("Asse X", cols, index=cols.index('Investimento CAPEX [€]'))
        y_var = col_y.selectbox("Asse Y", cols, index=cols.index('VAN [€]'))
        color_var = col_color.selectbox("Colore", cols, index=cols.index('Taglia Elettrolizzatore [kW]'))

        max_x = not any(w in x_var.lower() for w in ['costo', 'lcoh', 'investimento'])
        max_y = not any(w in y_var.lower() for w in ['costo', 'lcoh', 'investimento'])
        
        df_sorted, pareto_front, best_y = df_tutti.sort_values(by=x_var, ascending=not max_x), [], -float('inf') if max_y else float('inf')
        for _, row in df_sorted.iterrows():
            if (max_y and row[y_var] >= best_y) or (not max_y and row[y_var] <= best_y):
                pareto_front.append(row); best_y = row[y_var]
        df_pareto = pd.DataFrame(pareto_front)

        fig_sa = px.scatter(df_tutti, x=x_var, y=y_var, color=color_var, color_continuous_scale=["#07d5df", "#f407fe"], custom_data=['ID_Progetto'])
        if not df_pareto.empty:
            fig_sa.add_trace(go.Scatter(x=df_pareto.sort_values(by=x_var)[x_var], y=df_pareto.sort_values(by=x_var)[y_var], mode='lines', line=dict(color='red', width=3, dash='dash'), name='Pareto', hoverinfo='skip'))
        
        # FIX CURSORE: Tolte le frecce di pan, messo il click pulito
        fig_sa.update_traces(marker=dict(size=9, opacity=0.8), selector=dict(mode='markers'))
        fig_sa.update_layout(dragmode=False, clickmode='event+select', hovermode='closest')
        
        sel = st.plotly_chart(fig_sa, use_container_width=True, on_select="rerun", selection_mode="points")
        
        if sel and sel.selection.points:
            punto_cliccato = sel.selection.points[0]
            c_data = punto_cliccato.get("customdata", None)
            
            id_selezionato = None
            if c_data is not None:
                if isinstance(c_data, dict):
                    id_selezionato = c_data.get("ID_Progetto", list(c_data.values())[0])
                elif isinstance(c_data, (list, tuple)):
                    id_selezionato = c_data[0]
                else:
                    id_selezionato = c_data
            
            if id_selezionato is not None:
                det = df_tutti[df_tutti['ID_Progetto'] == id_selezionato].iloc[0]
                
                st.success("🎯 **Dettagli Progetto Selezionato:**")
                
                # RIGA 1: Metriche Economiche
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("💰 VAN", f"€ {det['VAN [€]']:,.0f}")
                c2.metric("📈 TIR", f"{det['TIR [%]']:.2f} %")
                c3.metric("⏳ Payback Time", f"{det['Payback Time [Anni]']:.1f} Anni")
                c4.metric("⚖️ LCOH Semplificato", f"€ {det['LCOH Semplificato [€/kg]']:.2f} /kg")
                
                # RIGA 2: Metriche di Impianto
                c5, c6, c7, c8 = st.columns(4)
                c5.metric("⚡ Elettrolizzatore", f"{det['Taglia Elettrolizzatore [kW]']:,.0f} kW")
                c6.metric("🔋 Batteria", f"{det['Taglia Batteria [kWh]']:,.0f} kWh")
                c7.metric("💶 Investimento Totale", f"€ {det['Investimento CAPEX [€]']:,.0f}")
                c8.metric("⚙️ Capacity Factor", f"{det['Capacity Factor [%]']:.1f} %")
                
                # RIGA 3: Metriche Energetiche
                c9, c10, c11, c12 = st.columns(4)
                c9.metric("💨 Produzione H2", f"{det['Produzione Idrogeno [kg]']:,.0f} kg/anno")
                c10.metric("🗑️ Curtailment (Spreco)", f"{det['Energia Sprecata (Curtailment) [kWh]']:,.0f} kWh")
            else:
                st.warning("⚠️ Clicca al centro di un pallino colorato per vedere i dettagli!")
        else:
            st.info("👆 Clicca su un pallino nel grafico per caricare i dati della configurazione.")
