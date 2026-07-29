# 🏭 Simulatore H2FAsT - Dashboard Interattiva

H2FAsT (Hydrogen Financial and Technical Simulator) è un'applicazione web avanzata per la simulazione e l'ottimizzazione tecno-economica di impianti ibridi off-grid per la produzione di **Idrogeno Verde**.

Il software integra dati meteorologici reali, un motore di calcolo vettoriale per la simulazione oraria (8760 ore) e un modulo finanziario, aiutando a individuare le migliori configurazioni impiantistiche (Frontiera di Pareto) bilanciando fonti rinnovabili multiple, accumulo a batterie ed elettrolizzatori.

---

## ✨ Funzionalità Principali

* 🌍 **Geolocalizzazione Multipla:** Acquisizione automatica dei profili di produzione tramite API (PVGIS per il sole, Open-Meteo per il vento). È possibile mappare l'impianto fotovoltaico e quello eolico nella stessa località oppure **in due località geografiche differenti** per massimizzare la compensazione climatica.
* 📊 **Integrazione Fonti Extra:** Possibilità di iniettare un profilo orario personalizzato tramite CSV (es. idroelettrico, biomasse o misurazioni reali sul campo) integrandolo nel mix energetico.
* ⚡ **Motore Vettorizzato "Bare-Metal":** Il backend matematico risolve i bilanci energetici annuali di centinaia di scenari in pochi secondi, gestendo dinamicamente curtailment, SoC delle batterie e logiche di attivazione degli elettrolizzatori.
* 🎯 **Esploratore Interattivo di Pareto:** L'utente può incrociare variabili economiche e tecniche (es. CAPEX vs VAN) per identificare visivamente i progetti dominanti. Selezionando un nodo sul grafico si accede all'identikit completo (LCOH, Efficienza, Spegnimenti).

---

## 🚀 Flusso di Lavoro (Wizard)

L'interfaccia utente è strutturata in step sequenziali sulla schermata principale:

### Step 0: Sorgenti Dati Energetici
Scegli l'origine dei tuoi flussi energetici:
* **Mappa API:** Cerca una località per centrare la mappa e piazza il pin. Tramite un apposito "switch" puoi decidere di separare la localizzazione del Fotovoltaico da quella dell'Eolico, usando due mappe indipendenti.
* **Caricamento Manuale & Fonti Extra:** In alternativa alle mappe, carica un file CSV. In questo step puoi anche aggiungere una fonte "Extra" per simulare un sistema trigen/ibrido.

### Step 1: Parametri Tecnici e Dimensionamento
* Inserisci le potenze nominali installate (Fotovoltaico ed Eolico).
* Scegli se prevedere o meno una Batteria nel sistema.
* Imposta i limiti tecnici dell'Elettrolizzatore (cut-off) e i range di ricerca per l'algoritmo di ottimizzazione.

### Step 2: CAPEX & Costi Impianto
* Definisci i costi di installazione unitari (€/kW) per le fonti rinnovabili.
* Immetti i costi specifici di Elettrolizzatori, Compressori e Storage.
* Aggiungi i costi fissi (Opere edili, Terreni, Stazioni di rifornimento).

### Step 3: Parametri Finanziari
* Inserisci il prezzo di vendita target dell'Idrogeno (€/kg).
* Definisci la durata del Business Plan, il Tasso di Sconto (per il VAN) e il tasso generale di inflazione.

### Step 4: Esecuzione e Dashboard
Cliccando su "Avvia Ottimizzazione", l'app popolerà una cassaforte dati navigabile:
1. **Tabelle Finanziarie:** Accesso a Conto Economico e Flussi di Cassa per ogni configurazione esplorata.
2. **Flussi Energetici:** Grafico orario completo di 8760 ore con barra di zoom inferiore, utile per osservare le dinamiche stagionali di curtailment e utilizzo batteria.
3. **Pareto:** L'area interattiva in cui cliccare i nodi ottimali e valutarne LCOH e redditività.

---

## ⚙️ Architettura Tecnica

Il progetto è modulare:
* `app.py`: Frontend in **Streamlit**. Gestisce chiamate API esterne, interazioni utente (`on_select`), state management e visualizzazione interattiva (`Plotly`, `Folium`).
* `motore_h2fast.py`: Backend. Implementa l'`Analisi_tecnica` (basata su matrici NumPy) e l'`Analisi_finanziaria` (algoritmi di Business Plan).
