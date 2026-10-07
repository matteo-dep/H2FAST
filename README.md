# ⚡ H2 FAST · Progettazione di impianti di produzione e stoccaggio di idrogeno

Strumento avanzato del **toolkit H2READY** per dimensionare un impianto di idrogeno verde (fonti rinnovabili,
elettrolizzatore, batteria, compressione e stoccaggio) e valutarne il business plan.

Il motore di calcolo è **H2FAsT** (Hydrogen Financial and Technical Simulator), sviluppato nel progetto
[AMETHYST](https://www.alpine-space.eu/project/amethyst/) (Interreg Alpine Space). L'interfaccia Streamlit
mantiene la struttura del programma originale (`INPUT.xlsx` → calcolo → `OUTPUT.xlsx`), divisa in tre step.

---

## Come funziona

| Step | Cosa si fa |
|---|---|
| **1. Parametri** | Le stesse sezioni dell'`INPUT.xlsx`: fonti energetiche, dati tecnici, investimenti, costi operativi, dati economico-finanziari, opzioni di output. Si può importare un `INPUT.xlsx` originale o un file `.json` salvato dall'app. |
| **2. Elaborazione** | Riepilogo, controlli di coerenza, stima dei tempi e calcolo di tutte le configurazioni elettrolizzatore/batteria. |
| **3. Risultati** | Migliori progetti, flussi energetici, business plan, esplorazione delle configurazioni (frontiera di Pareto), grafici di relazione e sensitivity, prezzo o incentivo di equilibrio, download di `OUTPUT.xlsx` nel formato originale. |

### Fonti dei dati orari
- **Fotovoltaico**: [PVGIS](https://re.jrc.ec.europa.eu/pvg_tools/) (JRC). Si usa la versione 5.3 e, se non risponde, la 5.2. Inclinazione e orientamento sono ottimali oppure scelti dall'utente.
- **Eolico**: velocità del vento a 100 m da [Open-Meteo](https://open-meteo.com/) (archivio ERA5), con una curva di potenza semplificata (cubica tra avvio e velocità nominale).
- **CSV**: il file orario scaricato da PVGIS oppure una colonna di 8760 valori in kW.
- **Fonte extra**: un secondo CSV da sommare al profilo (idroelettrico, biomassa, misure reali).

---

## Struttura del repository

```
app.py                 interfaccia Streamlit (3 step)
motore_h2fast.py       motore: Analisi_tecnica, Analisi_finanziaria, Analisi_combinata (struttura originale)
fonti_energia.py       PVGIS, Open-Meteo, lettura CSV, curva eolica
grafici.py             grafici Plotly della dashboard
h2ready.py             modulo condiviso del toolkit H2READY (identico negli altri repository)
requirements.txt
.streamlit/config.toml           tema e limiti di upload
.streamlit/secrets.toml.example  esempio di Secrets per l'accesso H2READY
esempi/profilo_sintetico_FV_1000kWp.csv   profilo SINTETICO di prova (non è un dato misurato)
```

---

## Pubblicazione su Streamlit Community Cloud

1. Crea un repository su GitHub e carica tutti i file di questa cartella, comprese le cartelle nascoste `.streamlit/` ed `esempi/`.
2. Su [share.streamlit.io](https://share.streamlit.io) scegli **Create app**, poi il repository e il branch, con file principale `app.py`.
3. In **Advanced settings** imposta **Python 3.12**.
4. *(Facoltativo)* Per l'accesso del toolkit, apri **Settings → Secrets** e incolla la sezione `[connections.gsheets]` usata dagli altri tool H2READY (vedi `.streamlit/secrets.toml.example`). L'app chiederà il codice ISTAT e controllerà l'autonomia del Comune sulla categoria **`fast`**. Senza Secrets l'app funziona da sola.
5. *(Facoltativo)* Aggiungi la riga del tool nel foglio `LINK` del master, con categoria `fast`, perché compaia tra i tool successivi.
6. **Deploy.** La prima elaborazione dopo un riavvio dell'app richiede 10-20 secondi in più, perché Numba compila il ciclo orario. Le elaborazioni successive sono immediate.

### Provarla in locale
```bash
python -m venv .venv && source .venv/bin/activate      # Windows: .venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

---

## Prestazioni e memoria

| FV | Configurazioni con batteria | Codice originale | Questa versione |
|---|---|---|---|
| 1 MW | 1.700 | ~31 s, 1,3 GB di RAM | ~3,5 s, circa 300 MB |
| 2 MW | 6.700 | ~122 s, 5,2 GB | ~6-7 s, circa 420 MB con l'Excel |

Per arrivarci:
- il ciclo orario con batteria è compilato con **Numba**. Il codice è lo stesso riga per riga e i risultati sono identici all'originale (scarto massimo 1·10⁻⁹);
- le serie orarie si conservano solo per i migliori N progetti e non per tutte le configurazioni;
- le tabelle del business plan si costruiscono solo per i progetti mostrati;
- il prezzo e l'incentivo di equilibrio si cercano per bisezione sulla stessa griglia di 0,5 €/kg dell'originale, quindi con lo stesso risultato.

Il file Excel di output si genera solo su richiesta, perché con molte configurazioni e i grafici di sensitivity può richiedere fino a un minuto.

---

## Differenze rispetto al codice originale

**Correzioni**
- Con la durata del debito uguale o maggiore della durata del piano (default 20 e 20) l'originale andava in errore. Ora le rate oltre il piano vengono troncate.
- Con un tasso nullo, sul debito e sul prestito ponte, il calcolo della rata dava 0/0.
- Il costo di eolico e fonte extra entra in investimento, IVA e ammortamenti. Nella versione Streamlit precedente veniva sommato e poi sovrascritto.
- Nella ricerca dell'incentivo di equilibrio senza batteria l'originale variava anche il prezzo dell'idrogeno.
- `costruzione_tabelle()` sovrascriveva `CostiPers` con una lista di valori negativi.
- Senza batteria l'ora con produzione esattamente uguale alla taglia dell'elettrolizzatore non era gestita.

**Aggiunte**
- **LCOH attualizzato**: (investimento netto IVA + costi operativi attualizzati) / produzione attualizzata, al tasso del VAN. Gli indicatori originali restano (costo medio operativo, full cost).
- **Classifica su più criteri** (da 2 a 4, con pesi). Si tengono le configurazioni non dominate (frontiera di Pareto) e, tra queste, si mettono prima quelle più vicine al punto ideale. Resta disponibile il criterio singolo dell'originale. Ogni criterio è spiegato nell'interfaccia: cosa misura e che impianto tende a scegliere.
- **Vincoli di progetto**: produzione minima di idrogeno (kg/anno) e quota minima della produzione rinnovabile destinata all'elettrolizzatore.
- **Vendita dell'eccedenza in rete** attivabile: se spenta, l'impianto è dedicato all'idrogeno e l'energia non usata non produce ricavi.
- **Stazione di rifornimento** opzionale.
- **Stoccaggio espresso in giorni di produzione** (l'originale usava una quota della produzione annua).
- Criteri aggiuntivi: LCOH, payback, investimento, quota di rinnovabile all'idrogeno.
- Fonte eolica e fonte extra con i rispettivi costi.
- TIR calcolato con lo stesso metodo delle secanti dell'originale, compilato con Numba: risultati identici, 10 volte più veloce. Nei casi senza soluzione l'originale faceva 1.000 iterazioni per configurazione.

**Valori predefiniti diversi dalla classe originale**
- Classifica: frontiera di Pareto su **LCOH + produzione di idrogeno** (nell'originale il criterio era scelto nell'Excel).
- Vendita dell'eccedenza in rete **spenta**. Se attivata, il prezzo è 0,10 €/kWh (nell'originale 1 €/kWh, un valore segnaposto).
- Stazione di rifornimento **esclusa**. Se inclusa, costa 500.000 € come nell'originale.
- Stoccaggio: **3 giorni di produzione** a **1.200 €/kg**. L'originale usava il 10% della produzione annua (36,5 giorni) con costo 0. Con 36,5 giorni a 1.200 €/kg lo stoccaggio diventa la prima voce di costo e porta l'LCOH sopra i 30 €/kg.
- Tutti gli altri valori sono quelli dell'originale: contributo pubblico 0%, debito 80% per 20 anni, prezzo dell'idrogeno 10 €/kg.

---

## Aspetti del modello da verificare

Questi punti non sono stati modificati, per restare fedeli all'originale:
1. Nel modello batteria `p_batt` è usato sia come potenza (kW) sia come capacità (kWh): la carica è limitata a `p_batt`.
2. In alcuni rami del ciclo orario il confronto è un'uguaglianza tra numeri decimali (`e_H2 == energia_disp`). Se nessun ramo è vero, la batteria di quell'ora resta a 0.
3. In un ramo la carica residua è impostata a `min_batt × e_batt_max` invece che a `min_batt × p_batt`.
4. Il "Levelized full cost" originale somma due volte il full cost: è tenuto solo per compatibilità.
5. Tutto il CAPEX delle rinnovabili è attribuito al progetto idrogeno; i ricavi dell'energia immessa sono compresi.
