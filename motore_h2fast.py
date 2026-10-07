"""
H2FAsT - Hydrogen Financial and Technical Simulator
Motore di calcolo (progetto AMETHYST, Interreg Alpine Space) adattato a Streamlit per il toolkit H2READY.

La struttura è quella del codice originale:
    Analisi_tecnica      -> bilancio energetico orario (8760 h) per ogni configurazione elettrolizzatore/batteria
    Analisi_finanziaria  -> business plan del singolo progetto (conto economico, flussi, VAN, TIR, payback)
    Analisi_combinata    -> esplorazione di tutte le configurazioni, classifica, simulazioni, file Excel di output

Differenze rispetto all'originale (le logiche di calcolo sono le stesse):
  1. Il ciclo orario con batteria è compilato con Numba: stesso codice riga per riga, risultati identici,
     circa 150 volte più veloce. Se Numba non è installato il codice gira comunque in Python puro.
  2. La matrice 'andamenti' (8760 valori orari per ogni configurazione) non viene più conservata per tutte
     le configurazioni: si salvano i totali annui e le serie orarie si ricalcolano solo per i migliori N
     progetti (sono gli unici che finiscono nell'output). La RAM passa da alcuni GB a pochi MB.
  3. Le tabelle del business plan si costruiscono solo per i migliori N progetti.
  4. Analisi_combinata riceve i parametri da un dizionario (l'interfaccia Streamlit) invece che da INPUT.xlsx
     e restituisce il file Excel in memoria invece di scrivere OUTPUT.xlsx su disco.
  5. Correzioni: crash con durata debito >= durata piano; CAPEX eolico/fonte extra ora conteggiato;
     incentivo di equilibrio senza batteria (l'originale variava anche il prezzo); LCOH attualizzato aggiunto.
"""
import io
import warnings
import sys

import numpy as np
import pandas as pd
import xlsxwriter

try:  # Numba è opzionale: senza, le stesse funzioni girano in Python puro (più lente)
    from numba import njit
    NUMBA_DISPONIBILE = True
except ImportError:  # pragma: no cover
    NUMBA_DISPONIBILE = False

    def njit(*args, **kwargs):
        if len(args) == 1 and callable(args[0]):
            return args[0]
        return lambda f: f

K_H2 = 120 / 3.6  # kWh per kg di H2 (PCI), come nell'originale


def granularita(potenza):
    """Numero di taglie esplorate in funzione della potenza (stessa regola dell'originale)."""
    if potenza <= 100:
        return 10
    elif potenza <= 500:
        return 20
    elif potenza <= 1000:
        return 50
    return 100


# =====================================================================================================
# KERNEL NUMERICI (compilati con Numba)
# =====================================================================================================

@njit(cache=True)
def eff_elc_nb(x):
    return (-6.1371 * x ** 6 + 24.394 * x ** 5 - 39.663 * x ** 4 + 33.988 * x ** 3
            - 16.412 * x ** 2 + 4.2929 * x + 0.1022)


@njit(cache=True)
def ciclo_orario_batteria(E_PV, p_elc, p_batt, min_elet, min_batt_pct, max_batt_pct, eff_batt):
    """Copia fedele del ciclo di Analisi_tecnica.run_analysis_battery_static_min per UN progetto.

    Restituisce i totali annui e le serie orarie (stesso significato degli indici di 'andamenti').
    """
    n = E_PV.shape[0]
    p_elc_min = min_elet * p_elc
    e_batt_max = p_batt * max_batt_pct
    E_H2_h = np.zeros(n)
    M_H2_h = np.zeros(n)
    E_im_h = np.zeros(n)
    E_batt = np.zeros(n)
    E_batt_disponibile_h = np.zeros(n)
    Energia_disp = np.zeros(n)
    max_erogabile = np.zeros(n)
    off_count = 0
    flag = 1
    count_to_24 = 0
    count2 = 0
    j = 0
    batt_prev = 0.0

    for i in range(n):
        p_pv = E_PV[i]
        e_im = 0.0
        e_H2 = 0.0
        m_H2 = 0.0
        energia_disp_batt = max(0.0, batt_prev - min_batt_pct * p_batt) if i != 0 else 0.0
        energia_disp = p_pv + eff_batt * energia_disp_batt

        if p_pv == 0:  # Caso 1: FV non produce, l'eventuale produzione viene solo dalla batteria
            if energia_disp < p_elc_min:
                if i == 0:
                    E_batt[0] = 0.0
                else:
                    E_batt[i] = batt_prev
                j += 1
                count_to_24 += 1
                if flag == 0:
                    off_count += 1
                flag = 1
                if count_to_24 == 24:
                    count2 += 1
                    count_to_24 = 0
            elif p_elc_min <= energia_disp < p_elc:
                e_H2 = min(energia_disp, e_batt_max * eff_batt)
                m_H2 = (e_H2 * eff_elc_nb(e_H2 / p_elc)) / K_H2
                E_batt[i] = batt_prev - e_H2 / eff_batt
                j += 1
                flag = 0
                count_to_24 = 0
            elif p_elc <= energia_disp:
                e_H2 = min(p_elc, e_batt_max, energia_disp_batt) * eff_batt
                m_H2 = (e_H2 * eff_elc_nb(e_H2 / p_elc)) / K_H2
                E_batt[i] = batt_prev - e_H2 / eff_batt
                j += 1
                flag = 0
                count_to_24 = 0

        elif 0 < p_pv < p_elc_min:  # Caso 2: il FV da solo non basta per il minimo tecnico
            if energia_disp < p_elc_min:
                j += 1
                if i == 0:
                    E_batt[i] = min(p_pv, e_batt_max) * eff_batt
                else:
                    E_batt[i] = min((min(p_pv, e_batt_max) * eff_batt + batt_prev), p_batt)
                e_im = max(p_pv - (E_batt[i] - batt_prev) / eff_batt, 0.0)
                count_to_24 += 1
                if flag == 0:
                    off_count += 1
                flag = 1
                if count_to_24 == 24:
                    count2 += 1
                    count_to_24 = 0
            elif p_elc_min <= energia_disp < p_elc:
                e_H2 = min(energia_disp, e_batt_max * eff_batt + p_pv)
                m_H2 = (e_H2 * eff_elc_nb(e_H2 / p_elc)) / K_H2
                if i == j:
                    if e_H2 == energia_disp:
                        E_batt[i] = min_batt_pct * e_batt_max
                    elif e_H2 == e_batt_max * eff_batt + p_pv:
                        E_batt[i] = batt_prev - e_batt_max
                j += 1
                flag = 0
                count_to_24 = 0
            elif p_elc <= energia_disp:
                e_H2 = min(e_batt_max + p_pv, p_elc)
                if i == j:
                    if e_H2 == e_batt_max + p_pv:
                        m_H2 = (e_H2 * eff_elc_nb(e_H2 / p_elc)) / K_H2
                        E_batt[i] = batt_prev - e_batt_max
                    elif e_H2 == p_elc:
                        m_H2 = e_H2 * 0.565 / K_H2
                        E_batt[i] = batt_prev - (p_elc - p_pv) / eff_batt
                j += 1
                flag = 0
                count_to_24 = 0

        elif p_elc_min <= p_pv < p_elc:  # Caso 3: FV tra minimo tecnico e potenza nominale
            if p_elc_min <= energia_disp < p_elc:
                e_H2 = min(energia_disp, e_batt_max * eff_batt + p_pv)
                m_H2 = (e_H2 * eff_elc_nb(e_H2 / p_elc)) / K_H2
                if i == j:
                    if e_H2 == energia_disp:
                        E_batt[i] = batt_prev - (energia_disp - p_pv) / eff_batt
                    elif e_H2 == e_batt_max * eff_batt + p_pv:
                        E_batt[i] = batt_prev - e_batt_max
                j += 1
                flag = 0
                count_to_24 = 0
            elif p_elc <= energia_disp:
                e_H2 = min(e_batt_max * eff_batt + p_pv, p_elc)
                if i == j:
                    if e_H2 == e_batt_max * eff_batt + p_pv:
                        m_H2 = (e_H2 * eff_elc_nb(e_H2 / p_elc)) / K_H2
                        E_batt[i] = batt_prev - e_batt_max
                    elif e_H2 == p_elc:
                        m_H2 = e_H2 * 0.565 / K_H2
                        E_batt[i] = batt_prev - (p_elc - p_pv) / eff_batt
                j += 1
                flag = 0
                count_to_24 = 0

        elif p_elc <= p_pv:  # Caso 4: surplus, elettrolizzatore al massimo e carica batteria
            e_H2 = p_elc
            m_H2 = e_H2 * 0.565 / K_H2
            E_batt[i] = batt_prev + min((p_pv - p_elc) * eff_batt, p_batt - batt_prev, e_batt_max)
            e_im = p_pv - p_elc - (E_batt[i] - batt_prev) / eff_batt
            j += 1
            flag = 0
            count_to_24 = 0

        E_H2_h[i] = e_H2
        M_H2_h[i] = m_H2
        E_im_h[i] = e_im
        E_batt_disponibile_h[i] = energia_disp_batt
        Energia_disp[i] = energia_disp
        max_erogabile[i] = e_batt_max * eff_batt + p_pv
        batt_prev = E_batt[i]

    return (E_H2_h.sum(), E_im_h.sum(), M_H2_h.sum(), off_count, count2,
            E_H2_h, M_H2_h, E_im_h, E_batt, E_batt_disponibile_h, Energia_disp, max_erogabile)


@njit(cache=True)
def esplora_batteria(E_PV, elc, batt, min_elet, min_batt_pct, max_batt_pct, eff_batt):
    """Esegue il ciclo orario per un blocco di configurazioni; restituisce solo i totali annui.

    colonne: E_H2, E_im, M_H2, OFF, count2
    """
    n = elc.shape[0]
    out = np.zeros((n, 5))
    for k in range(n):
        r = ciclo_orario_batteria(E_PV, elc[k], batt[k], min_elet, min_batt_pct, max_batt_pct, eff_batt)
        out[k, 0] = r[0]
        out[k, 1] = r[1]
        out[k, 2] = r[2]
        out[k, 3] = r[3]
        out[k, 4] = r[4]
    return out


@njit(cache=True)
def npv_nb(flussi, rate):
    """Stessa formula di Analisi_finanziaria.npv: flussi / exp(t * log(1 + rate))."""
    lr = np.log(1.0 + rate)
    s = 0.0
    for t in range(flussi.shape[0]):
        s += flussi[t] / np.exp(t * lr)
    return s


@njit(cache=True)
def irr_nb(flussi, guess, tol, max_iter):
    """Stesso metodo delle secanti di Analisi_finanziaria.irr (stessi passi e criteri di arresto).

    Restituisce (trovato, tasso). Serve a velocizzare i casi che non convergono e facevano 1000 iterazioni.
    """
    rate0 = guess
    rate1 = rate0 + 0.05
    npv0 = npv_nb(flussi, rate0)
    npv1 = npv_nb(flussi, rate1)
    for _ in range(max_iter):
        if abs(npv1 - npv0) < tol:
            return False, 0.0
        rate2 = rate1 - npv1 * (rate1 - rate0) / (npv1 - npv0)
        npv2 = npv_nb(flussi, rate2)
        if abs(npv2) < tol:
            return True, rate2
        rate0, rate1 = rate1, rate2
        npv0, npv1 = npv1, npv2
    return False, 0.0


def flussi_senza_batteria(E_PV, p_elc, min_elet):
    """Bilancio orario senza batteria (versione vettorizzata del ciclo originale run_analysis_nobattery)."""
    p_pv = E_PV
    p_elc_min = min_elet * p_elc
    cond_max = p_pv >= p_elc                      # originale: p_pv > p_elc (p_pv == p_elc non era gestito)
    cond_mid = (p_pv >= p_elc_min) & (p_pv < p_elc)
    cond_off = p_pv < p_elc_min

    E_H2_h = np.zeros_like(p_pv)
    E_H2_h[cond_max] = p_elc
    E_H2_h[cond_mid] = p_pv[cond_mid]

    E_im_h = np.zeros_like(p_pv)
    E_im_h[cond_max] = p_pv[cond_max] - p_elc
    E_im_h[cond_off] = p_pv[cond_off]

    M_H2_h = np.zeros_like(p_pv)
    M_H2_h[cond_max] = p_elc * 0.565 / K_H2
    var_mid = p_pv[cond_mid] / p_elc
    M_H2_h[cond_mid] = p_pv[cond_mid] * Analisi_tecnica.eff_elc(var_mid) / K_H2

    # spegnimenti: passaggi acceso -> spento (si parte da spento) e giorni interi di fermo
    is_off = cond_off.astype(int)
    off = int(np.sum(np.diff(np.insert(is_off, 0, 1)) == 1))
    diffs = np.diff(np.pad(is_off, (1, 1), "constant", constant_values=0))
    starts = np.where(diffs == 1)[0]
    ends = np.where(diffs == -1)[0]
    count2 = int(np.sum((ends - starts) // 24))
    return E_H2_h, M_H2_h, E_im_h, off, count2


# =====================================================================================================
# ANALISI TECNICA
# =====================================================================================================

class Analisi_tecnica:
    def __init__(self, file_csv, tipo_file, p_PV, dP_el, batteria, dP_bat, min_batt, max_batteria, lingua, eff_batt,
                 min_elet, E_PV=None, salva_andamenti=False, progresso=None):
        # DATI ESOGENI
        self.batteria = batteria            # "SI" o "NO"
        self.file_csv = file_csv            # percorso, file caricato (BytesIO) oppure None se si passa E_PV
        self.tipo_file = tipo_file          # "SI" se il file proviene da PVGIS
        self.p_PV = p_PV                    # potenza nominale totale delle fonti (kW)
        self.dP_el = dP_el                  # limite superiore della ricerca elettrolizzatore (x p_PV)
        self.dP_bat = dP_bat                # limite superiore della ricerca batteria (x p_elc)
        self.lingua = lingua
        self.E_PV = np.asarray(E_PV, dtype=float) if E_PV is not None else self.load_data()
        self.min_batt = min_batt            # SoC minimo (frazione)
        self.max_batteria = max_batteria    # quota massima erogabile in 1 h (frazione)
        self.min_elet = min_elet            # minimo tecnico elettrolizzatore (frazione)
        self.eff_batt = eff_batt
        self.e_pv = np.sum(self.E_PV)
        self.salva_andamenti = salva_andamenti   # True = comportamento originale (tutte le serie orarie in RAM)
        self.progresso = progresso               # funzione(frazione, testo) per la barra di avanzamento

        gran_elc = granularita(self.p_PV)
        self.P_elc = np.linspace(self.p_PV / gran_elc, self.p_PV * dP_el, gran_elc)

        # elenco delle configurazioni (stesso ordine dei cicli dell'originale)
        if self.batteria == "SI":
            elc, batt = [], []
            for p_elc in self.P_elc:
                P_batt = np.linspace(0, self.dP_bat * p_elc, granularita(p_elc))
                elc.extend([p_elc] * len(P_batt))
                batt.extend(P_batt)
            self.potenza_elett = np.array(elc, dtype=float)
            self.potenza_batt = np.array(batt, dtype=float)
        else:
            self.potenza_elett = self.P_elc.astype(float)
            self.potenza_batt = np.zeros(len(self.P_elc))

        self.qt_progetti = len(self.potenza_elett)
        n_serie = 11 if self.batteria == "SI" else 6
        self.andamenti = np.zeros((self.qt_progetti, n_serie, len(self.E_PV))) if salva_andamenti else None
        self.E_H2 = np.zeros(self.qt_progetti)
        self.E_im = np.zeros(self.qt_progetti)
        self.M_H2 = np.zeros(self.qt_progetti)
        self.Auto = np.zeros(self.qt_progetti)
        self.CF = np.zeros(self.qt_progetti)
        self.OFF = np.zeros(self.qt_progetti)
        self.spegn_giorn = np.zeros(self.qt_progetti)

    def load_data(self):
        """Lettura del profilo orario (stesse regole dell'originale; accetta anche file caricati in memoria)."""
        sorgente = self.file_csv
        if isinstance(sorgente, str) and not sorgente.lower().endswith(".csv"):
            sorgente = sorgente + ".csv"
        if self.tipo_file == "SI":
            if hasattr(sorgente, "read"):
                testo = sorgente.read()
                testo = testo.decode("utf-8", errors="ignore") if isinstance(testo, bytes) else testo
                lines = testo.splitlines()
            else:
                with open(sorgente, "r") as file:
                    lines = file.readlines()
            header = "time,P,G(i),H_sun,T2m,WS10m,Int"
            header_idx = next((i for i, line in enumerate(lines) if line.strip() == header), None)
            if header_idx is None:
                if self.lingua == "ENG":
                    raise ValueError("Error opening the PVGIS file: download a new PVGIS file.")
                raise ValueError("Errore nell'apertura del file PVGIS: scarica un nuovo file PVGIS")
            rows = []
            for line in lines[header_idx + 1:]:
                columns = line.strip().split(',')
                if len(columns) == 7:
                    rows.append(columns)
                else:
                    break
            data = pd.DataFrame(rows, columns=header.split(','))
            E_PV = pd.to_numeric(data['P'], errors='coerce').dropna()
            E_PV /= 1000  # da W a kW
            return E_PV.to_numpy()
        data = pd.read_csv(sorgente, sep=';', decimal=',')
        E_PV = pd.to_numeric(data.iloc[:, 0], errors='coerce').dropna()
        return E_PV.to_numpy()

    @staticmethod
    def eff_elc(x):
        """Rendimento dell'elettrolizzatore in funzione del carico relativo (polinomio originale)."""
        return (-6.1371 * x ** 6 + 24.394 * x ** 5 - 39.663 * x ** 4 + 33.988 * x ** 3 - 16.412 * x ** 2
                + 4.2929 * x + 0.1022)

    def run_analysis(self):
        if self.batteria == "SI":
            self.run_analysis_battery_static_min()
        elif self.batteria == "NO":
            self.run_analysis_nobattery()

    def _aggiorna(self, frazione, testo):
        if self.progresso is not None:
            self.progresso(frazione, testo)

    def _salva_totali(self, index, p_elc, e_H2, e_im, m_H2, off, count2):
        e_TOT = len(self.E_PV) * p_elc
        self.OFF[index] = off
        self.spegn_giorn[index] = off + count2
        self.CF[index] = e_H2 / e_TOT * 100 if e_TOT > 0 else 0
        self.Auto[index] = e_H2 / self.e_pv * 100 if self.e_pv > 0 else 0
        self.E_H2[index] = e_H2
        self.E_im[index] = e_im
        self.M_H2[index] = m_H2

    def run_analysis_nobattery(self):
        testo = "Technical Analysis without Battery" if self.lingua == "ENG" else "Analisi tecnica senza batteria"
        for index, p_elc in enumerate(self.P_elc):
            E_H2_h, M_H2_h, E_im_h, off, count2 = flussi_senza_batteria(self.E_PV, p_elc, self.min_elet)
            self._salva_totali(index, p_elc, E_H2_h.sum(), E_im_h.sum(), M_H2_h.sum(), off, count2)
            if self.salva_andamenti:
                self.andamenti[index] = self.andamenti_progetto(index)
            self._aggiorna((index + 1) / self.qt_progetti, testo)

    def run_analysis_battery_static_min(self):
        testo = "Technical Analysis with Battery" if self.lingua == "ENG" else "Analisi tecnica con batteria"
        blocco = 200
        for inizio in range(0, self.qt_progetti, blocco):
            fine = min(inizio + blocco, self.qt_progetti)
            out = esplora_batteria(self.E_PV, self.potenza_elett[inizio:fine], self.potenza_batt[inizio:fine],
                                   float(self.min_elet), float(self.min_batt), float(self.max_batteria),
                                   float(self.eff_batt))
            for k in range(fine - inizio):
                idx = inizio + k
                self._salva_totali(idx, self.potenza_elett[idx], out[k, 0], out[k, 1], out[k, 2],
                                   int(out[k, 3]), int(out[k, 4]))
                if self.salva_andamenti:
                    self.andamenti[idx] = self.andamenti_progetto(idx)
            self._aggiorna(fine / self.qt_progetti, testo)

    def andamenti_progetto(self, index):
        """Serie orarie di un progetto, con lo stesso ordine di righe della matrice 'andamenti' originale.

        Senza batteria: 0 E_H2, 1 M_H2, 2 E_im, 3 P_pv, 4 Pmin elett, 5 Pmax elett
        Con batteria:   + 6 E_batt, 7 Pmax batt, 8 E batt disponibile, 9 Energia disponibile, 10 Max erogabile
        """
        p_elc = float(self.potenza_elett[index])
        n = len(self.E_PV)
        if self.batteria == "NO":
            E_H2_h, M_H2_h, E_im_h, _, _ = flussi_senza_batteria(self.E_PV, p_elc, self.min_elet)
            return np.vstack([E_H2_h, M_H2_h, E_im_h, self.E_PV, np.full(n, self.min_elet * p_elc), np.full(n, p_elc)])
        p_batt = float(self.potenza_batt[index])
        r = ciclo_orario_batteria(self.E_PV, p_elc, p_batt, float(self.min_elet), float(self.min_batt),
                                  float(self.max_batteria), float(self.eff_batt))
        return np.vstack([r[5], r[6], r[7], self.E_PV, np.full(n, self.min_elet * p_elc), np.full(n, p_elc),
                          r[8], np.full(n, p_batt), r[9], r[10], r[11]])


def conta_configurazioni(p_PV, dP_el, batteria, dP_bat):
    """Numero di configurazioni che verranno esplorate (serve per la stima prima del calcolo)."""
    g = granularita(p_PV)
    P_elc = np.linspace(p_PV / g, p_PV * dP_el, g)
    if batteria != "SI":
        return len(P_elc)
    return int(sum(granularita(p) for p in P_elc))


# =====================================================================================================
# ANALISI FINANZIARIA
# =====================================================================================================

class Analisi_finanziaria:
    def __init__(self, Terr = 0, OpeE = 0, ImpPV1 = 0, ImpPV1eurokW = 800, EletteuroKW = 1650, CompreuroKW = 4000, AccuE = 0, AccuEeurokW = 200,
                 idrogstocperc = 1/10, StazzRif = 500000, SpeTOpere = 0, BombSto = 0, LavoImp = 0, CarrEll = 0, CapFac = 0, PotEle = 0,
                 tassoDEN = 0.005, ProdAnnuaIdrogkg = 0, bar = 300, costlitroacqua = 0.035, costounitariostoccaggio = 0, PercEserImp = 0.005, Percentimpianti = 0.0025,
                 PercentOpeEd = 0.0005, SpesAmmGen = 7000, Affitto = 0, CostiPersonal = 0, AltriCost = 0, IVAsualtriCost = "SI", DurPianEcon = 20, inflazione = 0.02, inflazionePrezzoElet = 0.02,
                 inflazioneIdrog= 0.01,  tassoVAN = 0.1, incentpubb = 0, duratincentpubb = 0, prezzoindrogeno = 10, ProdElettVend = 0, EnergiaAutocons = 0,
                 prezzoElett = 1, ContrPubb = 0, DebitoSenior = 0.8, DurDebitoSenior = 20, tassoDebito = 0.05, FreqPagamenti = 1, tassoPonte = 0,
                 DurataPonte = 0, aliquoMedia = 0.275, MaxInterssDed = 0.3, lingua = "ITA", Perciva = 0.22, spegn_giorn = 0,
                 ImpWind = 0, ImpWindeurokW = 0, ImpExtra = 0, ImpExtraeurokW = 0):
        # [Streamlit] altre fonti rinnovabili: potenza (kW) e costo (€/kW); entrano in investimento, IVA e ammortamenti
        self.ImpWind = ImpWind
        self.ImpWindeurokW = ImpWindeurokW
        self.ImpExtra = ImpExtra
        self.ImpExtraeurokW = ImpExtraeurokW
        self.LCOH = 0 # [Streamlit] costo livellato attualizzato dell'idrogeno
        # Variabili di Input:
        self.Terr = Terr # prezzo del terreno (0 se terreno di proprietà)
        self.OpeE = OpeE # prezzo delle opere edili/vani tecnici impianti
        self.ImpPV1 = ImpPV1 # potenza in kW del primo impianto PV
        self.ImpPV1eurokW = ImpPV1eurokW # costo per kW del primo impianto PV
        self.AccuE = AccuE # taglia della batteria
        self.AccuEeurokW = AccuEeurokW # prezzo per kW della batteria
        self.EletteuroKW = EletteuroKW # prezzo per kW dell'elettrolizzatore
        self.CompreuroKW = CompreuroKW # prezzo per kW del compressore
        self.idrogstocperc = idrogstocperc # % della produzione annua stoccata
        self.StazzRif = StazzRif # prezzo della stazione di rifornimento
        self.SpeTOpere = SpeTOpere # prezzo delle spese tecniche ammortizzabili per gli impianti
        self.BombSto = BombSto # prezzo delle bombole di stoccaggio
        self.LavoImp = LavoImp # prezzo dei lavori impiantistici ammortizzabili
        self.CarrEll = CarrEll # prezzo del carrello elevatore
        self.CapFac = CapFac # Capacity factor
        self.PotEle = PotEle # potenza dell'elettrolizzatore
        self.spegn_giorn = spegn_giorn
        self.tassoDEN = tassoDEN # tasso di decrescita dell'efficienza nominale annua
        self.ProdAnnuaIdrogkg = ProdAnnuaIdrogkg # produzione di idrogeno annua in kg
        self.bar = bar # bar ai quali si vuole stoccare l'idrogeno
        self.costlitroacqua = costlitroacqua # prezzo per litro di acqua utilizzato
        self.costounitariostoccaggio = costounitariostoccaggio
        self.PercEserImp = PercEserImp # % del costo degli impianti per il loro utilizzo
        self.Percentimpianti = Percentimpianti # % degli costo degli impianti che va in manutenzione
        self.PercentOpeEd = PercentOpeEd # % degli costo delle opere edili che va in manutenzione
        self.SpesAmmGen = SpesAmmGen # spese amministrative e generali relative al processo
        self.Affitto = Affitto # costo dell'affitto
        self.CostiPersonal = CostiPersonal # altri costi connessi al progetto
        self.DurPianEcon = int(DurPianEcon) # durata stimata del piano economico
        self.inflazione = inflazione # tasso di inflazione annua (colpisce i costi)
        self.inflazionePrezzoElet = inflazionePrezzoElet # tasso di inflazione del prezzo dell'energia
        self.inflazioneIdrog = inflazioneIdrog # tasso inflazione dell'idrogeno
        self.tassoVAN = tassoVAN # tasso di attaulizzazione per il calcolo del VAN
        self.incentpubb = incentpubb  # contributo in euro per kg di H2 prodotta
        self.duratincentpubb = int(duratincentpubb) # durata in anni dell'erogazione dell'incentivo pubblico
        self.prezzoindrogeno = prezzoindrogeno # prezzo di vendita dell'idrogeno
        self.ProdElettVend = ProdElettVend # elettricità prodotta per venderla alla rete
        self.EnergiaAutocons = EnergiaAutocons # Energia autoconsumata dall'impianto
        self.prezzoElett = prezzoElett # prezzo dell'energia venduta alla rete
        self.ContrPubb = ContrPubb # % del pogetto coperto da contributo pubblico
        self.DebitoSenior = DebitoSenior # % del pogetto coperto da debito senior
        self.DurDebitoSenior = int(DurDebitoSenior) # durata del debito
        self.tassoDebito = tassoDebito # tasso d'interesse applicato sul debito
        self.FreqPagamenti = int(FreqPagamenti) # frequenza dei pagamenti delle rate del debito in un anno
        self.tassoPonte = tassoPonte # tasso d'interesse sul debito per l'IVA
        self.DurataPonte = int(DurataPonte) # durata del debito ponte
        self.aliquoMedia = aliquoMedia # aliquota media sugli utili
        self.MaxInterssDed = MaxInterssDed # interessi massimi deducibili sull'EBITDA
        self.lingua = lingua # qui salvo la lingua nel quale si vuole l'output
        self.Perciva = Perciva # percentuale dell'Iva
        self.AltriCost = AltriCost # altri costi
        self.IVAsualtriCost = IVAsualtriCost # si applica IVA sugli altri costi
        # variabili di output
        self.potenza_compressore = 0 # potenza del compressore
        self.lc = 0 # sono due variabili che mi servono per calcolare la potenza del compressore
        self.prohH2s = 0
        self.idrogeno_stocc = 0 # stoccaggio in magazzino
        self.impianto_stocc = 0 # costo dell'impianto di stoccaggio
        self.imponibile1 = 0 # imponibile sulle opere edili
        self.iva1 = 0 # iva sulle opere edili
        self.imponibile2 = 0 # imponibile sugli impianti
        self.iva2 = 0 # iva sugli impianti
        self.iva = 0 # iva totale sugli investimenti
        self.investimento = 0 # investimento al netto dell'iva
        self.prodNElett = 0 # produzione nominale elettrolizzatore
        self.EffNomElett = 0 # efficienza nominale elettrolizzatore
        self.LavoSpecElett = 0 # Lavoro specifico di compressione
        self.ConsAcqua = 0 # Consumo acqua di processo litri all'anno
        self.CostAnnAcq = 0 # Costo annuo in acqua
        self.EserImp = 0 # costo per l'esercizio dell'impianto; è il 0,5% del investimento al netto dell'iva
        self.CostManImp = 0 # Manutenzione programmata e guasti, pari al 25% per gli impianti e 5% per le opere edili
        self.ContrPubbAss = 0 # Contributo pubblico in valore assoluto dell'investimento iniziale
        self.DebitoSeniorAss = 0 # Debito in valore assoluto dell'investimento iniziale
        self.CapProp = 0 # capitale proprio in valore assoluto dell'investimento iniziale
        self.TIR = 0 # Tassi interno di rendimento del progetto
        self.VAN = 0 # Valore attuale netto del progetto
        self.PAYBACK = 0 # Pay back semplice del progetto
        # Ricavi e Costi operativi
        self.produzioneH2 = np.zeros(self.DurPianEcon) # qui ci sono i livelli di produzione in kg di idrogeno considerando una riduzione nella produzione
        self.prezziindrogeno = np.zeros(self.DurPianEcon) # qui ci sono i prezzi dell'idrogeno inflazionati
        self.RicaviVenditeH2 = np.zeros(self.DurPianEcon) # qui ci sono tutti i ricavi anno per anno
        self.RicaviContributi = np.zeros(self.DurPianEcon) # qui ci sono i contributi per kg di idrogeno prodotto
        self.RicaviVendEnerg = np.zeros(self.DurPianEcon) # ricavi derivanti dalla vendita dell'energia
        self.TotaliRicavi = np.zeros(self.DurPianEcon) # somma di tutti i ricavi per ogni anno
        self.CostiAcqua = np.zeros(self.DurPianEcon) # costi dell'acqua per ogni anno
        self.CostiEserImpi = np.zeros(self.DurPianEcon) # costi dell'eserizio degli impianti anno per anno
        self.CostiAmmGen = np.zeros(self.DurPianEcon)  # costi Amministrativi e generali anno per anno
        self.CostiManuten = np.zeros(self.DurPianEcon) # costi di manutenzione degli impianti
        self.CostiPers = np.zeros(self.DurPianEcon) # costi del personale
        self.OtherCOSTS = np.zeros(self.DurPianEcon) # altri costi
        self.CostiAffitti = np.zeros(self.DurPianEcon) # costi degli affitti anno per anno
        self.costi_flussi_monetari = np.zeros(self.DurPianEcon) # costi che prevedono un esborso monetario anno per anno
        self.costi_operativi = np.zeros(self.DurPianEcon) # costi operativi anno per anno
        # Ammortamenti
        self.AMMTerr = np.zeros(self.DurPianEcon) # ammortamento dei terreni anno per anno
        self.AMMOpeE = np.zeros(self.DurPianEcon) # ammortamenti delle opere edili anno per anno
        self.AMMCont = np.zeros(self.DurPianEcon) # ammortamenti dei container anno per anno
        self.AMMMurCon = np.zeros(self.DurPianEcon) # ammortamenti del muro di contenimento anno per anno
        self.AMMPlanI = np.zeros(self.DurPianEcon) # ammortamenti delle Platea per posizionamento impianti anno per anno
        self.AMMRec = np.zeros(self.DurPianEcon) # ammortamenti della recinzione per posizionamento impianti anno per anno
        self.AMMVial = np.zeros(self.DurPianEcon) # ammortamenti del vialetto di accesso anno per anno
        self.AMMSpesT = np.zeros(self.DurPianEcon) # ammortamento delle spese tecniche anno per anno
        self.AMMFabbrTerr = np.zeros(self.DurPianEcon) # somma degli ammortamenti riguardanti i fabbricati e terreni anno per anno
        self.AMMPV1 = np.zeros(self.DurPianEcon) # ammortamento immpianto PV1 anno per anno
        self.AMMElett = np.zeros(self.DurPianEcon) # ammortamento elettrolizzatore anno per anno
        self.AMMSTOC = np.zeros(self.DurPianEcon) # ammortamento impianto stoccaggio anno per anno
        self.AMMAccuE = np.zeros(self.DurPianEcon) # ammortamento batterie
        self.AMMSTAZZ = np.zeros(self.DurPianEcon) # ammortamento stazione di rifornimento anno per anno
        self.AMMSpeTec = np.zeros(self.DurPianEcon) # ammortamento spese tecniche anno per anno
        self.AMMBombStoc = np.zeros(self.DurPianEcon) # ammortamento bombole stoccaggio anno per anno
        self.AMMLavImp = np.zeros(self.DurPianEcon) # ammortamento lavori impiantistici anno per anno
        self.AMMCARR = np.zeros(self.DurPianEcon) # ammortamento carrello elevatore anno per anno
        self.AMMAltreFER = np.zeros(self.DurPianEcon) # [Streamlit] ammortamento eolico e fonte extra
        self.AMMMacchImp = np.zeros(self.DurPianEcon) # ammortamento Ammortamento macchinari ed impianti anno per anno
        # Flussi dlegati a debiti ed imposte
        self.Flussi_debito_non_corretti = np.zeros(self.DurPianEcon) # Sarà una lista con un dizionaria che ha al suo interno tutti i flussi dei debiti
        self.SumAnnRata = np.zeros(self.DurPianEcon) # Qui abbiamo tutte le rate in un specifico anno (caso semplice rata costante annuo)
        self.SumAnnInt = np.zeros(self.DurPianEcon) # interessi da pagare in un determinato anno
        self.SumAnnCapit = np.zeros(self.DurPianEcon) # capitale da rimborsare per un determinato anno
        self.SumAnnSaldo = np.zeros(self.DurPianEcon) # saldo del debito per un determinato anno
        self.Flussi_debito = np.zeros(self.DurPianEcon) # flusso di cassa in uscita o in entrata(solo all'inizio) del debito
        self.IvaDebito = np.zeros(self.DurPianEcon) # debito di iva da pagare allo stato anno per anno
        self.IvaCredito = np.zeros(self.DurPianEcon) # credito di iva da incassare dallo stato anno per anno
        self.IvaNetto = np.zeros(self.DurPianEcon) # posizione netta con lo stato anno per anno
        self.imposte = np.zeros(self.DurPianEcon) # imposte sull'utile di periodo anno per anno
        self.PrestPontInter = np.zeros(self.DurPianEcon) # Flusso degli interessi sul debito a ponte
        self.PrestPontCapital = np.zeros(self.DurPianEcon) # Rimborso del capitale sul debito a ponte
        self.TOTinteressi = np.zeros(self.DurPianEcon) # Totale interessi pagati anno per anno, per tutti i mutui
        self.TOTcapitale = np.zeros(self.DurPianEcon) # Totale capitale rimborsato anno per anno
        # Vettori dei risultati
        self.EBITDA = np.zeros(self.DurPianEcon) # Utile al netto degli ammortamenti, interessi ed imposte
        self.EBIT = np.zeros(self.DurPianEcon) # utile al netto degli interessi ed immposte
        self.EBT = np.zeros(self.DurPianEcon) # utile al netto delle tasse
        self.UtileNetto = np.zeros(self.DurPianEcon) # profitto o perdita di periodo anno per anno
        self.EBITDAsukW = np.zeros(self.DurPianEcon) # EBITDA/kW anno per anno
        self.EBITsukW = np.zeros(self.DurPianEcon) # EBIT/kW anno per anno
        self.EBTsukW = np.zeros(self.DurPianEcon) # EBT/kW anno per anno
        self.UtileNettosukW = np.zeros(self.DurPianEcon) # UTILE NETTO/kW anno per anno
        self.FlussOperativo = np.zeros(self.DurPianEcon) # Flusso di cassa operativo del periodo
        self.CAPEX = np.zeros(self.DurPianEcon) # nella nostro caso è previsto un solo investimento ad inizio periodo
        self.FlussInvestime = np.zeros(self.DurPianEcon) # Flusso di cassa considerando gli investimenti
        self.FlussiIvaNetta = np.zeros(self.DurPianEcon) # Flusso di cassa considerando la posizione IVA con lo stato
        self.FlussiConFinanz = np.zeros(self.DurPianEcon)  # Flusso di cassa considerando la posizione finanziaria
        self.ToTimposte = np.zeros(self.DurPianEcon) # Flusso di cassa considerando le imposte
        self.FlussoNettoCassa = np.zeros(self.DurPianEcon) # Flusso di cassa al netto
        self.costo_medio_operativo_anno1 = 0 # inizializzo le variabili dei costi
        self.costo_medio_operativo = 0
        self.costo_medio_investimenti = 0
        self.costo_full_cost = 0
        self.costo_full_cost_levelized = 0

    def calcolo_investimento(self): # prima cosa che deve essere calcolata
        self.ConsAcqua = self.ProdAnnuaIdrogkg*8.92 # questa relazione è chimica; per ogni kg di idrogeno mi servono tot chili acqua
        self.CostAnnAcq = self.costlitroacqua*self.ConsAcqua # calcolo del costo per l'utilizzo dell'acqua
        self.lc = (14960*293.15/0.75*(((self.bar*10**5)/(3*10**6))**(((7/5)-1)/(7/5))))/1000
        self.prohH2s = (((self.ProdAnnuaIdrogkg/8600)*24)/8)/3600
        self.potenza_compressore = self.prohH2s*0.5*self.lc # in base a questa produzione necessito compressore con x potenza
        self.idrogeno_stocc = self.idrogstocperc*self.ProdAnnuaIdrogkg
        # stoccaggio dellìidrogeno dipende dalla pressione di stoccaggio
        # in caso di cambiamento dei costi modificare nel codice
        self.impianto_stocc = self.idrogeno_stocc*self.costounitariostoccaggio



        # qui calcolo l'ivestimento e l'iva e altri costi legati alla loro manutenzione e utilizzo
        self.imponibile1 = self.OpeE + self.Terr
        self.iva1 = self.OpeE*(self.Perciva) + self.Terr*(1.05 - 1)


        self.imponibile2 = (self.ImpPV1*self.ImpPV1eurokW +
                            self.PotEle*self.EletteuroKW +
                            self.CompreuroKW*self.potenza_compressore +
                            self.AccuE*self.AccuEeurokW +
                            self.impianto_stocc +
                            self.BombSto +
                            self.StazzRif +
                            self.SpeTOpere +
                            self.LavoImp +
                            self.CarrEll)
        self.iva2 = self.ImpPV1*self.ImpPV1eurokW*(self.Perciva) + self.PotEle*self.EletteuroKW*(self.Perciva) + self.CompreuroKW*self.potenza_compressore*(self.Perciva) + self.AccuE*self.AccuEeurokW*(self.Perciva) + self.impianto_stocc*(self.Perciva) + self.StazzRif*(self.Perciva) + self.SpeTOpere*(self.Perciva) + self.BombSto*(self.Perciva) + self.LavoImp*(self.Perciva) + self.CarrEll*(self.Perciva)
        # [Streamlit] eolico e fonte extra (nell'app precedente il loro costo veniva sommato e poi sovrascritto)
        self.costo_altre_FER = self.ImpWind*self.ImpWindeurokW + self.ImpExtra*self.ImpExtraeurokW
        self.imponibile2 += self.costo_altre_FER
        self.iva2 += self.costo_altre_FER*self.Perciva

        #RISULTATI
        self.iva = self.iva1 + self.iva2
        self.investimento = self.imponibile1 + self.imponibile2
        self.EserImp = self.PercEserImp*self.imponibile2
        self.CostManImp = self.Percentimpianti*self.imponibile2 + self.PercentOpeEd*self.imponibile1

    def calcolo_econ_fin(self):
        # calcoli in valoro assoluto delle fonti di finanziamento
        self.ContrPubbAss = self.ContrPubb*self.investimento
        self.DebitoSeniorAss = self.DebitoSenior*self.investimento
        self.CapProp = 1 - self.ContrPubb - self.DebitoSenior
        self.CapPropAss = self.CapProp*self.investimento

    def calcolo_ricavi(self):
        # qui calcolo tutti i ricavi
        for index in range(self.DurPianEcon):
            if index == 0:
                # il primo anno non produco perché devo costruire l'impianto
                self.produzioneH2[index] = 0
                self.prezziindrogeno[index] = 0
                self.RicaviVendEnerg[index] = 0
            elif index == 1:
                self.produzioneH2[index] = self.ProdAnnuaIdrogkg
                self.prezziindrogeno[index] = self.prezzoindrogeno
                self.RicaviVendEnerg[index] = self.ProdElettVend*self.prezzoElett
            elif index != 1 and index != 0:
                self.produzioneH2[index] = self.produzioneH2[index - 1]*(1-self.tassoDEN) # la produzione si riduce annualmente
                self.prezziindrogeno[index] = self.prezziindrogeno[index - 1]*(1+self.inflazioneIdrog)
                self.RicaviVendEnerg[index] = self.RicaviVendEnerg[index - 1]*(1+self.inflazionePrezzoElet)
        for index,(el1,el2) in enumerate(zip(self.produzioneH2,self.prezziindrogeno)):
            self.RicaviVenditeH2[index] = el1*el2
        for index,el in enumerate(self.produzioneH2):
            if index <= self.duratincentpubb:
                self.RicaviContributi[index] = el*self.incentpubb
        for index,(el1,el2,el3) in enumerate(zip(self.RicaviContributi, self.RicaviVenditeH2, self.RicaviVendEnerg)):
            self.TotaliRicavi[index] = el1 + el2 + el3

    def calcolo_costi_operativi(self):
        # qui calcolo tutti i costi
        for index in range(self.DurPianEcon):
            if index == 0:
                self.CostiAcqua[index] = 0
                self.CostiEserImpi[index] = 0
                self.CostiAmmGen[index] = 0
                self.CostiAffitti[index] = 0
                self.CostiManuten[index] = 0
                self.CostiPers[index] = 0
                self.OtherCOSTS[index] = 0
            elif index == 1:
                self.CostiAcqua[index] = self.CostAnnAcq
                self.CostiEserImpi[index] = self.EserImp
                self.CostiAmmGen[index] = self.SpesAmmGen
                self.CostiAffitti[index] = self.Affitto
                self.CostiManuten[index] = self.CostManImp
                self.CostiPers[index] = self.CostiPersonal
                self.OtherCOSTS[index] = self.AltriCost
            elif index != 1 and index != 0:
                self.CostiAcqua[index] = self.CostiAcqua[index-1]*(1+self.inflazione)
                self.CostiEserImpi[index] = self.CostiEserImpi[index-1]*(1+self.inflazione)
                self.CostiAmmGen[index] = self.CostiAmmGen[index-1]*(1+self.inflazione)
                self.CostiAffitti[index] = self.CostiAffitti[index-1]*(1+self.inflazione)
                self.CostiManuten[index] = self.CostiManuten[index-1]*(1+self.inflazione)
                self.CostiPers[index] = self.CostiPers[index-1]*(1+self.inflazione)
                self.OtherCOSTS[index] = self.OtherCOSTS[index-1]*(1+self.inflazione)


        for index,(el1,el2,el3,el4,el5,el6,el7) in enumerate(zip(self.CostiAcqua, self.CostiEserImpi, self.CostiAmmGen, self.CostiAffitti,self.CostiManuten,self.CostiPers,self.OtherCOSTS)):
            self.costi_flussi_monetari[index] = el1 + el2 + el3 + el4 + el6 + el7
            self.costi_operativi[index] = el1 + el2 + el3 + el4 + el5 + el6 + el7


        # Ammortamento fabbricati e terreni, macchinari ed impianti
        for index in range(self.DurPianEcon):
            if index == 0:
                self.AMMTerr[index] = 0
                self.AMMOpeE[index] = 0
                self.AMMPV1[index] = 0
                self.AMMElett[index] = 0
                self.AMMAccuE[index] = 0
                self.AMMSTOC[index] = 0
                self.AMMSTAZZ[index] = 0
                self.AMMSpeTec[index] = 0
                self.AMMBombStoc[index] = 0
                self.AMMLavImp[index] = 0
                self.AMMCARR[index] = 0
            elif index != 0:
                self.AMMTerr[index] = (self.Terr*(1 - self.ContrPubb))/self.DurPianEcon
                self.AMMOpeE[index] = (self.OpeE*(1 - self.ContrPubb))/self.DurPianEcon
                self.AMMPV1[index] = ((self.ImpPV1*self.ImpPV1eurokW)*(1 - self.ContrPubb))/self.DurPianEcon
                self.AMMElett[index] = ((self.PotEle*self.EletteuroKW)*(1 - self.ContrPubb))/self.DurPianEcon
                self.AMMAccuE[index] = ((self.AccuE*self.AccuEeurokW)*(1 - self.ContrPubb))/self.DurPianEcon
                self.AMMSTOC[index] = ((self.impianto_stocc)*(1 - self.ContrPubb))/self.DurPianEcon
                self.AMMSTAZZ[index] = (self.StazzRif*(1 - self.ContrPubb))/self.DurPianEcon
                self.AMMSpeTec[index] = (self.SpeTOpere*(1 - self.ContrPubb))/self.DurPianEcon
                self.AMMBombStoc[index] = (self.BombSto*(1 - self.ContrPubb))/self.DurPianEcon
                self.AMMLavImp[index] = (self.LavoImp*(1 - self.ContrPubb))/self.DurPianEcon
                self.AMMCARR[index] = (self.CarrEll*(1 - self.ContrPubb))/self.DurPianEcon
                self.AMMAltreFER[index] = (self.costo_altre_FER*(1 - self.ContrPubb))/self.DurPianEcon



        for index,(el1,el2) in enumerate(zip(self.AMMTerr,self.AMMOpeE)):
            self.AMMFabbrTerr[index] = el1 + el2



        for index,(el1, el2, el3, el4, el5, el6, el7, el8, el9) in enumerate(zip(self.AMMPV1, self.AMMElett, self.AMMSTOC, self.AMMSTAZZ, self.AMMSpeTec, self.AMMBombStoc, self.AMMLavImp, self.AMMCARR, self.AMMAccuE)):
            self.AMMMacchImp[index] = el1 + el2 + el3 + el4 + el5 + el6 + el7 + el8 + el9 + self.AMMAltreFER[index]

    def calcola_flussi_capitali(self):
        # gestione dei flussi del debito
        if self.DurDebitoSenior != 0:
            tasso_periodico = ((1 + self.tassoDebito) ** (1 / self.FreqPagamenti)) - 1
            numero_rate = int(self.DurDebitoSenior * self.FreqPagamenti)
            if tasso_periodico == 0:  # [Streamlit] tasso nullo: rata costante senza interessi (evita 0/0)
                rata = self.DebitoSeniorAss / numero_rate
            else:
                rata = self.DebitoSeniorAss * (tasso_periodico * (1 + tasso_periodico) ** numero_rate) / ((1 + tasso_periodico) ** numero_rate - 1) #calcolo controllato in data 27/11/24 da Stefano Pagani e Giacomo Pamìo
            saldo = self.DebitoSeniorAss
            dtype = [('Periodo', int), ('Rata', float), ('Interesse', float), ('Capitale', float), ('Saldo', float)]
            self.Flussi_debito_non_corretti = np.zeros(numero_rate, dtype=dtype)
            for i in range(numero_rate):
                interesse = saldo * tasso_periodico
                capitale = rata - interesse
                saldo = saldo - capitale
                self.Flussi_debito_non_corretti[i] = (i + 1, rata, interesse, capitale, saldo)
            for anno in range(min(self.DurDebitoSenior, self.DurPianEcon)):  # [Streamlit] debito più lungo del piano: si tronca
                start = anno * self.FreqPagamenti
                end = start + self.FreqPagamenti
                ann_rata = np.sum(self.Flussi_debito_non_corretti['Rata'][start:end])
                ann_int = np.sum(self.Flussi_debito_non_corretti['Interesse'][start:end])
                ann_capit = np.sum(self.Flussi_debito_non_corretti['Capitale'][start:end])
                ann_saldo = self.Flussi_debito_non_corretti['Saldo'][end - 1] if end - 1 < numero_rate else 0
                self.SumAnnRata[anno] = ann_rata
                self.SumAnnInt[anno] = ann_int
                self.SumAnnCapit[anno] = ann_capit
                self.SumAnnSaldo[anno] = ann_saldo
            self.Flussi_debito[0] = self.DebitoSeniorAss - self.SumAnnRata[0]
            # [Streamlit] l'originale andava in errore con DurDebitoSenior >= DurPianEcon (es. 20 e 20)
            n_rate_piano = len(self.Flussi_debito[1:self.DurDebitoSenior+1])
            self.Flussi_debito[1:self.DurDebitoSenior+1] = self.SumAnnRata[:n_rate_piano]
            if self.DurDebitoSenior + 1 < self.DurPianEcon:
                self.Flussi_debito[self.DurDebitoSenior + 1:] = 0

    def calcolo_iva(self):
        if self.IVAsualtriCost == "SI":
            for index,(el1,el2,el3,el4,el5,el6,el7,el8) in enumerate(zip(self.RicaviVenditeH2, self.RicaviVendEnerg,self.CostiAcqua, self.CostiEserImpi, self.CostiAmmGen, self.CostiAffitti, self.CostiManuten, self.OtherCOSTS)):
                self.IvaDebito[index] = el1*self.Perciva + el2*self.Perciva
                self.IvaCredito[index] = el3*self.Perciva + el4*self.Perciva + el5*self.Perciva + el6*self.Perciva + el7*self.Perciva + el8*self.Perciva
        elif self.IVAsualtriCost == "NO":
            for index,(el1,el2,el3,el4,el5,el6,el7) in enumerate(zip(self.RicaviVenditeH2, self.RicaviVendEnerg,self.CostiAcqua, self.CostiEserImpi, self.CostiAmmGen, self.CostiAffitti, self.CostiManuten)):
                self.IvaDebito[index] = el1*self.Perciva + el2*self.Perciva
                self.IvaCredito[index] = el3*self.Perciva + el4*self.Perciva + el5*self.Perciva + el6*self.Perciva + el7*self.Perciva
        self.IvaCredito[0] += self.iva
        for index,(el1,el2) in enumerate(zip(self.IvaDebito, self.IvaCredito)):
            self.IvaNetto[index] = el2 - el1
        for index,el in enumerate(self.IvaNetto): # IvaNetto
            # noinspection PyTypeChecker
            if el > 0:
                if index + 1 < self.DurPianEcon:
                    self.IvaNetto[index + 1] += el
                    self.IvaNetto[index] = 0

    def calcolo_prestito_ponte(self):
        # gestione prestito a ponte
        tasso_annuo = self.tassoPonte
        if self.DurataPonte != 0:
            if tasso_annuo == 0:  # [Streamlit] evita 0/0 con tasso nullo
                rata_fissa = self.iva / self.DurataPonte
            else:
                rata_fissa = self.iva * (tasso_annuo * (1 + tasso_annuo) ** self.DurataPonte) / ((1 + tasso_annuo) ** self.DurataPonte - 1)
            saldo = self.iva
            for index in range(1, min(self.DurataPonte, self.DurPianEcon - 1) + 1):  # [Streamlit] non oltre il piano
                interesse = saldo * tasso_annuo
                capitale = rata_fissa - interesse
                saldo -= capitale
                self.PrestPontInter[index] = interesse
                self.PrestPontCapital[index] = capitale

    def somma_interessi(self):
        for index,(el1,el2) in enumerate(zip(self.PrestPontInter, self.SumAnnInt)):
            self.TOTinteressi[index] = el1 + el2
        for index,(el1,el2) in enumerate(zip(self.PrestPontCapital, self.SumAnnCapit)):
            self.TOTcapitale[index] = el1 + el2

    def calcolo_utili(self):
        for index,(el1,el2) in enumerate(zip(self.TotaliRicavi, self.costi_operativi)):
            self.EBITDA[index] = el1 - el2
        for index,(el1,el2,el3) in enumerate(zip(self.EBITDA, self.AMMFabbrTerr, self.AMMMacchImp)):
            self.EBIT[index] = el1 - el2 - el3
        for index,(el1,el2) in enumerate(zip(self.EBIT, self.TOTinteressi)):
            self.EBT[index] = el1 - el2
        for index,(el1, el2, el3) in enumerate(zip(self.EBIT, self.TOTinteressi, self.EBITDA)):
            if el1 >= 0:
                interest_deductible = min(el2, self.MaxInterssDed * el3)
                taxable_income = el1 - interest_deductible
                imposte = taxable_income * self.aliquoMedia
                self.UtileNetto[index] = el1 - imposte - el2
                self.ToTimposte[index] = imposte
            else:
                self.UtileNetto[index] = el1 - el2
                self.ToTimposte[index] = 0

    def calcolo_flussi_monenari(self):
        for index,(el1,el2) in enumerate(zip(self.TotaliRicavi, self.costi_flussi_monetari)):
            self.FlussOperativo[index] = el1 - el2


        #step 1
        self.CAPEX = self.CostiManuten.copy()
        # step 2
        self.CAPEX[0] -= self.CapPropAss
        # step 3
        if self.DurataPonte == 0:
            self.CAPEX[0] -= self.iva



        # step 4
        for index,(el1,el2) in enumerate(zip(self.FlussOperativo, self.CAPEX)):
            if index == 0:
                self.FlussInvestime[index] = el1 + el2
            if index != 0:
                self.FlussInvestime[index] = el1 - el2


        # step 5
        for index,(el1,el2) in enumerate(zip(self.FlussInvestime, self.IvaNetto)):
            self.FlussiIvaNetta[index] = el1 + el2

        # step 6
        for index,(el1,el2,el3,el4,el5) in enumerate(zip(self.FlussiIvaNetta, self.PrestPontInter, self.PrestPontCapital, self.SumAnnCapit, self.SumAnnInt)):
            self.FlussiConFinanz[index] = el1 - el2 - el3 - el4 - el5
        # step 7
        for index,(el1,el2) in enumerate(zip(self.FlussiConFinanz, self.ToTimposte)):
            self.FlussoNettoCassa[index] = el1 - el2

    def calcolo_costo_medio(self):
        #costi_op_totali_attualizzati = self.FlussOperativo.sum()  ---> precedente calcolo, modificato da giacomo pamìo e Stefano Pagani il 28/11/2024.

        costi_op_totali_attualizzati = self.costi_operativi[1] * self.DurPianEcon  #--->  la somma scontata dei costi poerativi al tasso di inflazione usato è uguale ai costi operativi del primo anno per la durata del progetto
        costi_op_1 = self.FlussOperativo[1]
        costi_cap = self.investimento
        interessi = self.TOTinteressi.sum()
        produzione = self.produzioneH2.sum()
        produzione_1 = self.produzioneH2[1]


        if produzione == 0:
            produzione = 1
        if produzione_1 == 0:
            produzione_1 = 1


        self.costo_medio_operativo_anno1 = costi_op_1/produzione_1
        self.costo_medio_operativo = costi_op_totali_attualizzati/produzione
        self.costo_medio_investimenti = costi_cap/produzione
        self.costo_full_cost = self.costo_medio_operativo + self.costo_medio_investimenti
        self.costo_full_cost_levelized = self.costo_full_cost + interessi / produzione + self.costo_full_cost

    def npv(self, rate):
        times = np.arange(len(self.FlussoNettoCassa))
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", category=RuntimeWarning)
            try:
                discounted_cash_flows = self.FlussoNettoCassa / np.exp(times * np.log(1 + rate))
                return np.sum(discounted_cash_flows)
            except (OverflowError, FloatingPointError):
                return None

    def irr(self, guess=0.1, tol=1e-6, max_iter=1000):
        # [Streamlit] stesso algoritmo compilato con Numba (vedi irr_nb); l'originale è in irr_python
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", category=RuntimeWarning)
            trovato, tasso = irr_nb(np.asarray(self.FlussoNettoCassa, dtype=float), float(guess), float(tol),
                                    int(max_iter))
        return float(tasso) if trovato else None

    def irr_python(self, guess=0.1, tol=1e-6, max_iter=1000):
        rate0 = guess
        rate1 = rate0 + 0.05
        npv0 = self.npv(rate0)
        npv1 = self.npv(rate1)
        if npv0 is None or npv1 is None:
            return None
        for _ in range(max_iter):
            if abs(npv1 - npv0) < tol:
                return None
            try:
                rate2 = rate1 - npv1 * (rate1 - rate0) / (npv1 - npv0)
                npv2 = self.npv(rate2)
            except (OverflowError, FloatingPointError, ZeroDivisionError):
                return None
            if npv2 is None or abs(npv2) < tol:
                return rate2
            rate0, rate1 = rate1, rate2
            npv0, npv1 = npv1, npv2
        return None

    def calcolo_indici_fin(self):
        self.VAN = self.npv(self.tassoVAN)
        try:
            self.TIR = self.irr()
        except Exception as e:
            self.TIR = None
        # calcolo Payback
        cumulativo = 0
        self.PAYBACK = None
        for i, flusso in enumerate(self.FlussoNettoCassa):
            cumulativo += flusso
            if cumulativo >= 0:
                self.PAYBACK = i
                break

    # noinspection PyUnboundLocalVariable
    def costruzione_tabelle(self):
        # qui vado a modificare i i valori (ovvero metto il valore negativo se sono costi...) in modo tale che quando vengono scritti nel file excel
        self.CostiAcquaDF = [-abs(value) if value != 0 else 0 for value in self.CostiAcqua]
        self.CostiEserImpiDF = [-abs(value) if value != 0 else 0 for value in self.CostiEserImpi]
        self.CostiAmmGenDF = [-abs(value) if value != 0 else 0 for value in self.CostiAmmGen]
        self.CostiManutenDF = [-abs(value) if value != 0 else 0 for value in self.CostiManuten]
        self.CostiAffittiDF = [-abs(value) if value != 0 else 0 for value in self.CostiAffitti]
        self.AMMMacchImpDF = [-abs(value) if value != 0 else 0 for value in self.AMMMacchImp]
        self.AMMFabbrTerrDF = [-abs(value) if value != 0 else 0 for value in self.AMMFabbrTerr]
        self.TOTinteressiDF = [-abs(value) if value != 0 else 0 for value in self.TOTinteressi]
        self.ToTimposteDF = [-abs(value) if value != 0 else 0 for value in self.ToTimposte]
        self.IvaDebitoDF = [-abs(value) if value != 0 else 0 for value in self.IvaDebito]
        self.TOTcapitaleDF = [-abs(value) if value != 0 else 0 for value in self.TOTcapitale]
        self.CAPEXDF = [-abs(value) if value != 0 else 0 for value in self.CAPEX]
        self.CostiPersDF = [-abs(value) if value != 0 else 0 for value in self.CostiPers]  # [Streamlit] prima sovrascriveva CostiPers
        self.OtherCOSTSDF = [-abs(value) if value != 0 else 0 for value in self.OtherCOSTS]
        # costruzione delle tabelle in inglese
        if self.lingua == "ENG":
            data1 = {
                'Revenue from H2 Sales': self.RicaviVenditeH2,
                'Revenue from Contributions': self.RicaviContributi,
                'Revenue from Energy Sales': self.RicaviVendEnerg,
                'Total Revenue': self.TotaliRicavi,
                'Water Costs': self.CostiAcquaDF,
                'Plant Operating Costs': self.CostiEserImpiDF,
                'General Administrative Costs': self.CostiAmmGenDF,
                'Maintenance': self.CostiManutenDF,
                'Passive Rents': self.CostiAffittiDF,
                'Labor Cost': self.CostiPersDF,
                'Other Costs' : self.OtherCOSTSDF,
                'EBITDA': self.EBITDA,
                'Depreciation of Machinery and Plants': self.AMMMacchImpDF,
                'Depreciation of Buildings and Land': self.AMMFabbrTerrDF,
                'EBIT': self.EBIT,
                'Interest Expenses': self.TOTinteressiDF,
                'EBT': self.EBT,
                'Average Taxes': self.ToTimposteDF,
                'Net Profit': self.UtileNetto
            }
            data2 = {
                'Revenue from H2 Sales': self.RicaviVenditeH2,
                'Revenue from Contributions': self.RicaviContributi,
                'Revenue from Energy Sales': self.RicaviVendEnerg,
                'Water Costs': self.CostiAcquaDF,
                'Plant Operating Costs': self.CostiEserImpiDF,
                'General Administrative Costs': self.CostiAmmGenDF,
                'Passive Rents': self.CostiAffittiDF,
                'Labor Cost': self.CostiPersDF,
                'Other Costs' : self.OtherCOSTSDF,
                'Operating Cash Flow': self.FlussOperativo,
                'CAPEX': self.CAPEXDF,
                'Cash Flow from Investments': self.FlussInvestime,
                'Net VAT': self.IvaNetto,
                'Cash Flows Net of VAT': self.FlussiIvaNetta,
                'Total Interests': self.TOTinteressiDF,
                'Capital Repayment': self.TOTcapitaleDF,
                'Financial Cash Flows': self.FlussiConFinanz,
                'Average Taxes': self.ToTimposteDF,
                'Net Cash Flow': self.FlussoNettoCassa,
            }
            data3 = {
                'NPV': [self.VAN],
                'IRR': [self.TIR],
                'PAYBACK': [self.PAYBACK],

                #'Average Operating Cost Year 1': [self.costo_medio_operativo_anno1],

                'Average Operating Cost [€/kg]': [self.costo_medio_operativo],
                'Average Investment Cost [€/kg]': [self.costo_medio_investimenti],
                'Average Full Cost [€/kg]': [self.costo_full_cost],
                'Discounted LCOH [€/kg]': [self.LCOH],

                #'Levelised Full Cost': [self.costo_full_cost_levelized],

                'Net Investment (excl. VAT)': [self.investimento],
                'Capacity Factor [%]' : [self.CapFac],
                'PV Plant Power [kW]' : [self.ImpPV1],
                'Battery size [kWh]' : [self.AccuE],
                'Electrolyzer Power [kW]' : [self.PotEle],
                'Compressor Power [kW]' : [self.potenza_compressore],
                'Self-Consumed Energy [kWh]' : [self.EnergiaAutocons],
                'Hydrogen Production [kg/year]' : [self.ProdAnnuaIdrogkg],
                'Electricity Injected into the Grid [kWh]' : [self.ProdElettVend],
                "Hydrogen Price  [€]" : [self.prezzoindrogeno],
                'Storage system capacity [kg]': [self.idrogeno_stocc],
                'Storage system cost  [€]' : [self.impianto_stocc],
                'Shutoff count': [self.spegn_giorn]
            }
        # costruzione delle tabelle in inglese
        if self.lingua == "ITA":
            data1 = {
                'Ricavi Vendite H2': self.RicaviVenditeH2,
                'Ricavi Contributi': self.RicaviContributi,
                'Ricavi Vendite Energia': self.RicaviVendEnerg,
                'Totali Ricavi': self.TotaliRicavi,
                'Costi Acqua': self.CostiAcquaDF,
                'Costi Esercizio Impianti': self.CostiEserImpiDF,
                'Costi Amministrativi Generali': self.CostiAmmGenDF,
                'Manutenzioni': self.CostiManutenDF,
                'Affitti passivi': self.CostiAffittiDF,
                'Costo del Personale': self.CostiPersDF,
                'Altri Costi' : self.OtherCOSTSDF,
                'EBITDA': self.EBITDA,
                'Ammortamento Macchinari e Impianti': self.AMMMacchImpDF,
                'Ammortamenti Fabbricati e Terreni': self.AMMFabbrTerrDF,
                'EBIT': self.EBIT,
                'interessi passivi': self.TOTinteressiDF,
                'EBT': self.EBT,
                'Imposte medie': self.ToTimposteDF,
                'UtileNetto': self.UtileNetto
            }
            data2 = {
                'Ricavi Vendite H2': self.RicaviVenditeH2,
                'Ricavi Contributi': self.RicaviContributi,
                'Ricavi Vendite Energia': self.RicaviVendEnerg,
                'Costi Acqua': self.CostiAcquaDF,
                'Costi Esercizio Impianti': self.CostiEserImpiDF,
                'Costi Amministrativi Generali': self.CostiAmmGenDF,
                'Affitti passivi': self.CostiAffittiDF,
                'Costo del Personale': self.CostiPersDF,
                'Altri Costi' : self.OtherCOSTSDF,
                'Flusso di Cassa Operativo': self.FlussOperativo,
                'CAPEX': self.CAPEXDF,
                'Flusso di cassa da investimenti': self.FlussInvestime,
                'IVA netto': self.IvaNetto,
                "Flussi al netto dell'Iva": self.FlussiIvaNetta,
                'Interessi TOT': self.TOTinteressiDF,
                'Rimborso Capitale': self.TOTcapitaleDF,
                'Flussi di cassa finanziari': self.FlussiConFinanz,
                'Imposte medie': self.ToTimposteDF,
                'Flusso di cassa netto': self.FlussoNettoCassa,
            }
            data3 = {
                'VAN': [self.VAN],
                'TIR': [self.TIR],
                'PAYBACK [anni]': [self.PAYBACK],

                #'Costo medio operativo anno 1' : [self.costo_medio_operativo_anno1],

                'Costo medio operativo [€/kg]' : [self.costo_medio_operativo],
                'Costo medio investimenti [€/kg]': [self.costo_medio_investimenti],
                'Full cost medio [€/kg]': [self.costo_full_cost],
                'LCOH attualizzato [€/kg]': [self.LCOH],

                #'Full cost levelised' : [self.costo_full_cost_levelized],

                'Investimento netto IVA' : [self.investimento],
                'Capacity Factor [%]' : [self.CapFac],
                "Potenza dell'impianto fotovoltaico [kW]" : [self.ImpPV1],
                'Taglia batteria [kWh]' : [self.AccuE],
                "Potenza dell'elettrolizzatore [kW]" : [self.PotEle],
                'Potenza del compressore [kW]' : [self.potenza_compressore],
                'Energia autoconsumata [kWh]' : [self.EnergiaAutocons],
                'Produzione di idrogeno [kg/anno]' : [self.ProdAnnuaIdrogkg],
                'Elettricità immessa in rete [kWh]' : [self.ProdElettVend],
                "Prezzo dell'idrogeno [€]" : [self.prezzoindrogeno],
                'Capacità impianto stoccaggio [kg]': [self.idrogeno_stocc],
                'Costo impianto di stoccaggio [€]' : [self.impianto_stocc],
                'Conteggio spegnimenti' : [self.spegn_giorn]
            }
        self.dfContoEconomico = pd.DataFrame(data1) # data frame del conto economico
        self.dfContoEconomico = self.dfContoEconomico.map(lambda x: f"{x:.2f}") # lascia due cifre decimali
        self.dfContoEconomico = self.dfContoEconomico.transpose()
        self.dfFlussiMonetari = pd.DataFrame(data2) # data frame del flusso di cassa
        self.dfFlussiMonetari = self.dfFlussiMonetari.map(lambda x: f"{x:.2f}")
        self.dfFlussiMonetari = self.dfFlussiMonetari.transpose()
        self.dfIndiciFinanziari = pd.DataFrame(data3) # data frame dei vari indici

    def calcolo_lcoh(self):
        # [Streamlit] LCOH attualizzato al tasso del VAN:
        # (investimento netto IVA + costi operativi attualizzati) / produzione attualizzata.
        # Esclude ricavi, contributi, imposte e oneri finanziari: è il costo industriale del kg di H2.
        anni = np.arange(self.DurPianEcon)
        sconto = (1 + self.tassoVAN) ** anni
        costi = self.investimento + np.sum(self.costi_operativi / sconto)
        produzione = np.sum(self.produzioneH2 / sconto)
        self.LCOH = costi / produzione if produzione > 0 else np.nan

    def RUN(self, tabelle=True):
        # qui creo una funzione che fa tutte le operazioni di sopra in modo ordinato
        self.calcolo_investimento()
        self.calcolo_econ_fin()
        self.calcolo_ricavi()
        self.calcolo_costi_operativi()
        self.calcola_flussi_capitali()
        self.calcolo_iva()
        self.calcolo_prestito_ponte()
        self.somma_interessi()
        self.calcolo_utili()
        self.calcolo_flussi_monenari()
        self.calcolo_costo_medio()
        self.calcolo_indici_fin()
        self.calcolo_lcoh()
        if tabelle:  # [Streamlit] le tabelle servono solo per i progetti mostrati
            self.costruzione_tabelle()


# =====================================================================================================
# VARIABILI DI RISULTATO (criteri di ottimizzazione, grafici di relazione e sensitivity)
# =====================================================================================================
# codice interno: (etichetta ITA, etichetta ENG, unità, verso "max" = meglio alto / "min" = meglio basso)
VARIABILI = {
    "VAN": ("VAN", "NPV", "€", "max"),
    "TIR": ("TIR", "IRR", "-", "max"),
    "PAYBACK": ("Payback", "Payback", "anni", "min"),
    "LCOH": ("LCOH attualizzato", "Discounted LCOH", "€/kg", "min"),
    "costo_full_cost": ("Full cost medio", "Full Cost", "€/kg", "min"),
    "costo_medio_operativo": ("Costo medio operativo", "Average operating cost", "€/kg", "min"),
    "costo_full_cost_levelized": ("Levelized full cost (formula originale)", "Levelized Full Cost", "€/kg", "min"),
    "prezzoindrogeno": ("Prezzo idrogeno", "Hydrogen Price", "€/kg", "min"),
    "investimento": ("Investimento (netto IVA)", "Net investment", "€", "min"),
    "PotEle": ("Elettrolizzatore", "Electrolyzer", "kW", "max"),
    "AccuE": ("Batteria", "Batteries", "kWh", "max"),
    "potenza_compressore": ("Potenza compressore", "Compressor power", "kW", "max"),
    "ProdAnnuaIdrogkg": ("Produzione idrogeno", "Hydrogen production", "kg/anno", "max"),
    "ProdElettVend": ("Energia immessa in rete", "Energy injected into the grid", "kWh", "min"),
    "EnergiaAutocons": ("Energia autoconsumata", "Self-consumed energy", "kWh", "max"),
    "Autoconsumo": ("Autoconsumo", "Self-consumption", "%", "max"),
    "CapFac": ("Capacity factor", "Capacity factor", "%", "max"),
    "impianto_stocc": ("Costo stoccaggio idrogeno", "Hydrogen storage cost", "€", "min"),
    "spegn_giorn": ("Conteggio spegnimenti", "Shutoff count", "n.", "min"),
}

# Ordinamento come nell'originale: per questi attributi "meglio basso" (ordinamento decrescente,
# il migliore finisce in fondo alla lista); per tutti gli altri ordinamento crescente.
ORDINE_DECRESCENTE = ["prezzoindrogeno", "costo_medio_operativo", "costo_full_cost", "costo_full_cost_levelized",
                      "spegn_giorn", "LCOH", "investimento", "PAYBACK", "impianto_stocc"]

# Nomi usati nel file INPUT.xlsx originale -> codice interno
ALIAS_VARIABILI = {
    "IRR": "TIR", "NPV": "VAN", "Electrolyzer": "PotEle", "Elettrolizzatore": "PotEle",
    "Batteries": "AccuE", "Batteria": "AccuE", "Compressor power": "potenza_compressore",
    "Potenza compressore": "potenza_compressore", "Hydrogen production": "ProdAnnuaIdrogkg",
    "Produzione idrogeno": "ProdAnnuaIdrogkg", "Energy injected into the grid": "ProdElettVend",
    "Energia immessa": "ProdElettVend", "Self-consumed energy": "EnergiaAutocons",
    "Energia autoconsumata": "EnergiaAutocons", "Capacity factor": "CapFac",
    "Average operating cost": "costo_medio_operativo", "Costo medio operativo": "costo_medio_operativo",
    "Full Cost": "costo_full_cost", "Levelized Full Cost": "costo_full_cost_levelized",
    "Hydrogen storage cost": "impianto_stocc", "Shutoff count": "spegn_giorn",
    "Prezzo idrogeno": "prezzoindrogeno", "Hydrogen Price": "prezzoindrogeno",
}
CRITERI_PREZZO = ["prezzoindrogeno", "Prezzo idrogeno", "Hydrogen Price"]
SIM_PREZZO = ["prezzo", "Prezzo idrogeno", "Hydrogen Price"]
SIM_INCENTIVO = ["incentivo", "Incentivo pubblico per kg di idrogeno venduto",
                 "Public incentive per kilogram of hydrogen sold"]

PASSO_EQUILIBRIO = 0.5      # €/kg, come nell'originale
MASSIMO_EQUILIBRIO = 200.0  # €/kg: oltre questo valore la configurazione è considerata non sostenibile

# Parametri letti dal foglio INPUT.xlsx originale (stesso ordine di Analisi_combinata.read_excel)
LISTA_NOMI1 = ["Terr", "OpeE", "StazzRif", "SpeTOpere", "BombSto", "LavoImp", "CarrEll", "ImpPV1eurokW",
               "EletteuroKW", "CompreuroKW", "AccuEeurokW", "costounitariostoccaggio", "idrogstocperc"]
LISTA_NOMI2 = ["costlitroacqua", "PercEserImp", "Percentimpianti", "PercentOpeEd", "SpesAmmGen", "Affitto",
               "CostiPersonal", "AltriCost", "IVAsualtriCost"]
LISTA_NOMI3 = ["DurPianEcon", "inflazione", "inflazionePrezzoElet", "tassoVAN", "incentpubb", "duratincentpubb",
               "prezzoindrogeno", "inflazioneIdrog", "prezzoElett", "ContrPubb", "DebitoSenior", "DurDebitoSenior",
               "tassoDebito", "FreqPagamenti", "DurataPonte", "tassoPonte", "aliquoMedia", "MaxInterssDed", "Perciva"]
LISTA_NOMI4 = ["file_csv", "tipo_file", "p_PV", "batteria", "min_batt", "max_batteria", "tassoDEN", "bar",
               "eff_batt", "min_elet"]
LISTA_NOMI5 = ["dP_el", "dP_bat"]


def nome_variabile(codice, lingua="ITA", con_unita=True):
    ita, eng, um, _ = VARIABILI.get(codice, (codice, codice, "", "max"))
    nome = eng if lingua == "ENG" else ita
    return f"{nome} [{um}]" if con_unita and um not in ("", "-") else nome


def leggi_input_excel(file):
    """Legge un INPUT.xlsx nel formato originale H2FAST e restituisce un dizionario di parametri.

    Usa le stesse posizioni di cella di Analisi_combinata.read_excel; il profilo orario (file_csv) va
    comunque caricato a parte, perché nell'Excel c'è solo il nome del file.
    """
    df = pd.read_excel(file, sheet_name=0)
    valori = (df.iloc[7:20, 3].tolist() + df.iloc[7:16, 6].tolist() + df.iloc[7:26, 9].tolist()
              + df.iloc[7:17, 12].tolist() + df.iloc[7:9, 15].tolist())
    nomi = LISTA_NOMI1 + LISTA_NOMI2 + LISTA_NOMI3 + LISTA_NOMI4 + LISTA_NOMI5
    par = {}
    for nome, valore in zip(nomi, valori):
        if isinstance(valore, str) and valore.strip().upper() == "YES":
            valore = "SI"
        if isinstance(valore, float) and np.isnan(valore):
            continue
        par[nome] = valore

    def cella(r, c):
        try:
            v = df.iloc[r, c]
            return None if (isinstance(v, float) and np.isnan(v)) else v
        except IndexError:
            return None

    def colonna(start, col):
        out = []
        for v in df.iloc[start:, col]:
            if pd.isna(v):
                break
            out.append(ALIAS_VARIABILI.get(v, v))
        return out

    attributo = cella(29, 2)
    if attributo is not None:
        par["attributo"] = ALIAS_VARIABILI.get(attributo, attributo)
    if cella(31, 2) is not None:
        par["n_progetti"] = int(cella(31, 2))
    if cella(33, 2) in ("ITA", "ENG"):
        par["lingua"] = cella(33, 2)
    rel = cella(28, 8)
    par["relazione"] = "SI" if rel in ("SI", "YES") else "NO"
    if par["relazione"] == "SI":
        par["variable_1_list"] = colonna(30, 5)
        par["variable_2_list"] = colonna(30, 8)
    sim = cella(28, 12)
    par["si_fa_simulazione"] = "SI" if sim in ("SI", "YES") else "NO"
    if par["si_fa_simulazione"] == "SI":
        att = cella(29, 11)
        par["attributo_simulazione"] = "incentivo" if att in SIM_INCENTIVO else "prezzo"
    sa = cella(28, 15)
    par["si_fa_grafico_SA"] = "SI" if sa in ("SI", "YES") else "NO"
    if par["si_fa_grafico_SA"] == "SI":
        par["SA_variable_list"] = colonna(29, 14)
    for k in ("p_PV",):
        if k in par:
            par[k] = float(par[k])
    return par


# =====================================================================================================
# ANALISI COMBINATA
# =====================================================================================================

class Analisi_combinata:
    def __init__(self, parametri, E_PV, progresso=None):
        """parametri: dizionario con gli stessi nomi del file INPUT.xlsx (più p_Wind, p_Extra e relativi costi).
        E_PV: profilo orario totale delle fonti (kW, 8760 valori)."""
        warnings.filterwarnings("ignore", category=UserWarning, module='openpyxl')
        self.progresso = progresso
        self.read_parametri(parametri)
        self.p_totale = self.p_PV + self.p_Wind + self.p_Extra

        self._aggiorna(0.0, "Analisi tecnica...")
        self.analisi1 = Analisi_tecnica(None, "NO", self.p_totale, self.dP_el, self.batteria, self.dP_bat,
                                        self.min_batt, self.max_batteria, self.lingua, self.eff_batt, self.min_elet,
                                        E_PV=E_PV, progresso=lambda f, t: self._aggiorna(0.45 * f, t))
        self.analisi1.run_analysis()
        self.potenza_impianto = self.p_PV
        self.valori_simulatore = []
        self.prezzi_eq = []
        self.VAN_eq = []
        self.traduci_attributo()
        self.analisi_tutti_i_progetti()
        self.tabella_topN_per(self.attributo, self.n_progetti)
        if self.attributo not in CRITERI_PREZZO and self.si_fa_simulazione in ["SI", "YES"]:
            self.simulatore_simulatore(self.attributo_simulazione)
        self._aggiorna(1.0, "Fatto" if self.lingua == "ITA" else "Done")

    # ---------------------------------------------------------------------------------------------
    def _aggiorna(self, frazione, testo):
        if self.progresso is not None:
            self.progresso(min(max(frazione, 0.0), 1.0), testo)

    def read_parametri(self, parametri):
        """Equivalente di read_excel: crea gli attributi della simulazione a partire dal dizionario."""
        predefiniti = {
            "p_Wind": 0.0, "ImpWindeurokW": 0.0, "p_Extra": 0.0, "ImpExtraeurokW": 0.0,
            "attributo": "VAN", "n_progetti": 5, "lingua": "ITA", "relazione": "NO",
            "variable_1_list": [], "variable_2_list": [], "si_fa_simulazione": "NO",
            "attributo_simulazione": "prezzo", "si_fa_grafico_SA": "NO", "SA_variable_list": [],
            # [Streamlit] classifica su più criteri (frontiera di Pareto) e vincoli di progetto
            "criteri": [], "pesi": [], "vincolo_H2_min": 0.0, "vincolo_auto_min": 0.0,
        }
        for nome, valore in {**predefiniti, **parametri}.items():
            if valore == "YES":
                valore = "SI"
            setattr(self, nome, valore)
        self.n_progetti = int(self.n_progetti)
        self.criteri = [ALIAS_VARIABILI.get(c, c) for c in (self.criteri or []) if c not in CRITERI_PREZZO]
        if len(self.criteri) >= 2:
            self.attributo = self.criteri[0]   # per compatibilità con le parti che leggono un solo criterio
        if not self.pesi or len(self.pesi) != len(self.criteri):
            self.pesi = [1.0] * len(self.criteri)

    def traduci_attributo(self):
        self.attributo = ALIAS_VARIABILI.get(self.attributo, self.attributo)

    def traduci_lista_per_il_codice(self, lista):
        return [ALIAS_VARIABILI.get(el, el) for el in lista]

    def traduci_nome(self, el):
        return nome_variabile(ALIAS_VARIABILI.get(el, el), self.lingua, con_unita=False)

    def aggiungi_unita_di_misura(self, el):
        return nome_variabile(ALIAS_VARIABILI.get(el, el), self.lingua, con_unita=True)

    # ---------------------------------------------------------------------------------------------
    def _analisi_finanziaria(self, idx, prezzo=None, incentivo=None):
        """Crea l'Analisi_finanziaria della configurazione idx (stessi argomenti dei cicli originali)."""
        a1 = self.analisi1
        return Analisi_finanziaria(
            Terr=self.Terr, OpeE=self.OpeE, ImpPV1=self.potenza_impianto, ImpPV1eurokW=self.ImpPV1eurokW,
            EletteuroKW=self.EletteuroKW, CompreuroKW=self.CompreuroKW, AccuE=a1.potenza_batt[idx],
            AccuEeurokW=self.AccuEeurokW, idrogstocperc=self.idrogstocperc, StazzRif=self.StazzRif,
            SpeTOpere=self.SpeTOpere, BombSto=self.BombSto, LavoImp=self.LavoImp, CarrEll=self.CarrEll,
            CapFac=a1.CF[idx], PotEle=a1.potenza_elett[idx], tassoDEN=self.tassoDEN,
            ProdAnnuaIdrogkg=a1.M_H2[idx], bar=self.bar, costlitroacqua=self.costlitroacqua,
            costounitariostoccaggio=self.costounitariostoccaggio, PercEserImp=self.PercEserImp,
            Percentimpianti=self.Percentimpianti, PercentOpeEd=self.PercentOpeEd, SpesAmmGen=self.SpesAmmGen,
            Affitto=self.Affitto, CostiPersonal=self.CostiPersonal, AltriCost=self.AltriCost,
            IVAsualtriCost=self.IVAsualtriCost, DurPianEcon=self.DurPianEcon, inflazione=self.inflazione,
            inflazionePrezzoElet=self.inflazionePrezzoElet, inflazioneIdrog=self.inflazioneIdrog,
            tassoVAN=self.tassoVAN, incentpubb=self.incentpubb if incentivo is None else incentivo,
            duratincentpubb=self.duratincentpubb,
            prezzoindrogeno=self.prezzoindrogeno if prezzo is None else prezzo,
            ProdElettVend=a1.E_im[idx], EnergiaAutocons=a1.E_H2[idx], prezzoElett=self.prezzoElett,
            ContrPubb=self.ContrPubb, DebitoSenior=self.DebitoSenior, DurDebitoSenior=self.DurDebitoSenior,
            tassoDebito=self.tassoDebito, FreqPagamenti=self.FreqPagamenti, tassoPonte=self.tassoPonte,
            DurataPonte=self.DurataPonte, aliquoMedia=self.aliquoMedia, MaxInterssDed=self.MaxInterssDed,
            lingua=self.lingua, Perciva=self.Perciva, spegn_giorn=a1.spegn_giorn[idx],
            ImpWind=self.p_Wind, ImpWindeurokW=self.ImpWindeurokW,
            ImpExtra=self.p_Extra, ImpExtraeurokW=self.ImpExtraeurokW)

    def _cerca_equilibrio(self, idx, variabile="prezzo", stretto=False):
        """Minimo prezzo (o incentivo), a passi di 0,5 €/kg, che rende il VAN >= 0 (o > 0 se stretto).

        L'originale aumentava il valore di 0,5 alla volta; qui si cerca lo stesso valore della stessa
        griglia per bisezione (il VAN cresce con il prezzo), con un numero di calcoli molto minore.
        """
        def prova(k):
            v = k * PASSO_EQUILIBRIO
            a = self._analisi_finanziaria(idx, prezzo=v if variabile == "prezzo" else None,
                                          incentivo=v if variabile == "incentivo" else None)
            a.RUN(tabelle=False)
            ok = a.VAN is not None and (a.VAN > 0 if stretto else a.VAN >= 0)
            return ok, a

        hi = int(round(MASSIMO_EQUILIBRIO / PASSO_EQUILIBRIO))
        ok, migliore = prova(hi)
        if not ok:
            return None, None
        lo = 0
        while hi - lo > 1:
            mid = (lo + hi) // 2
            ok, a = prova(mid)
            if ok:
                hi, migliore = mid, a
            else:
                lo = mid
        return hi * PASSO_EQUILIBRIO, migliore

    @staticmethod
    def _riga_risultati(idx, a, e_pv):
        return {
            "ID": idx, "VAN": a.VAN, "TIR": a.TIR, "PAYBACK": a.PAYBACK, "LCOH": a.LCOH,
            "costo_full_cost": a.costo_full_cost, "costo_medio_operativo": a.costo_medio_operativo,
            "costo_medio_investimenti": a.costo_medio_investimenti,
            "costo_full_cost_levelized": a.costo_full_cost_levelized, "prezzoindrogeno": a.prezzoindrogeno,
            "investimento": a.investimento, "PotEle": a.PotEle, "AccuE": a.AccuE, "ImpPV1": a.ImpPV1,
            "potenza_compressore": a.potenza_compressore, "ProdAnnuaIdrogkg": a.ProdAnnuaIdrogkg,
            "ProdElettVend": a.ProdElettVend, "EnergiaAutocons": a.EnergiaAutocons,
            "Autoconsumo": a.EnergiaAutocons / e_pv * 100 if e_pv > 0 else 0.0,
            "CapFac": a.CapFac, "impianto_stocc": a.impianto_stocc, "idrogeno_stocc": a.idrogeno_stocc,
            "spegn_giorn": a.spegn_giorn,
        }

    def analisi_tutti_i_progetti(self):
        """Analisi finanziaria di ogni configurazione: si conservano solo gli indicatori (non le istanze)."""
        n = self.analisi1.qt_progetti
        righe = []
        criterio_prezzo = self.attributo in CRITERI_PREZZO
        testo = "Analisi finanziaria" if self.lingua == "ITA" else "Financial analysis"
        passo = max(1, n // 100)
        for idx in range(n):
            if criterio_prezzo:
                prezzo, a = self._cerca_equilibrio(idx, "prezzo", stretto=False)
                if a is None:
                    continue
            else:
                a = self._analisi_finanziaria(idx)
                a.RUN(tabelle=False)
            righe.append(self._riga_risultati(idx, a, self.analisi1.e_pv))
            if idx % passo == 0 or idx == n - 1:
                self._aggiorna(0.45 + 0.45 * (idx + 1) / n, testo)
        self.risultati = pd.DataFrame(righe)
        for col in ("VAN", "TIR", "PAYBACK"):
            if col in self.risultati:
                self.risultati[col] = pd.to_numeric(self.risultati[col], errors="coerce")
        # [Streamlit] vincoli di progetto: produzione minima di H2 e quota minima di rinnovabile all'elettrolizzatore
        if len(self.risultati):
            self.risultati["ammissibile"] = ((self.risultati["ProdAnnuaIdrogkg"] >= float(self.vincolo_H2_min))
                                             & (self.risultati["Autoconsumo"] >= float(self.vincolo_auto_min) * 100))
        self.vincoli_rispettati = bool(len(self.risultati)) and bool(self.risultati["ammissibile"].any())

    def ordina_lista(self, attributo):
        """Ordina le configurazioni come nell'originale: il migliore è l'ULTIMO elemento."""
        attributo = ALIAS_VARIABILI.get(attributo, attributo)
        df = self.risultati.dropna(subset=[attributo])
        return df.sort_values(attributo, ascending=attributo not in ORDINE_DECRESCENTE, kind="stable")

    def ordina_lista2(self, attributo1, attributo2):
        attributo1 = ALIAS_VARIABILI.get(attributo1, attributo1)
        attributo2 = ALIAS_VARIABILI.get(attributo2, attributo2)
        df = self.risultati.dropna(subset=[attributo1, attributo2])
        return df.sort_values(attributo1, ascending=attributo1 not in ORDINE_DECRESCENTE, kind="stable")

    def _candidati(self):
        """[Streamlit] Configurazioni che rispettano i vincoli; se nessuna li rispetta si usano tutte."""
        df = self.risultati
        if "ammissibile" in df and df["ammissibile"].any():
            return df[df["ammissibile"]]
        return df

    def ordina_multicriterio(self, df, criteri, pesi):
        """[Streamlit] Classifica su più criteri.

        1. Ordinamento per fronti di Pareto: il fronte 1 contiene le configurazioni per cui nessun'altra è
           migliore o uguale su tutti i criteri e strettamente migliore su almeno uno; il fronte 2 quelle non
           dominate una volta tolto il fronte 1, e così via.
        2. Dentro ogni fronte si mettono prima le configurazioni più vicine al punto ideale (il valore migliore
           di ogni criterio), con i criteri normalizzati 0-1 e pesati.
        Restituisce df ordinato dal migliore e le colonne 'fronte' e 'distanza_ideale'.
        """
        X = []
        for c in criteri:
            v = pd.to_numeric(df[c], errors="coerce").to_numpy(dtype=float)
            if VARIABILI.get(c, ("", "", "", "max"))[3] == "max":
                v = -v                                   # tutto diventa "da minimizzare"
            finiti = np.isfinite(v)
            peggiore = (np.max(v[finiti]) + 1 + abs(np.max(v[finiti])) * 0.01) if finiti.any() else 0.0
            X.append(np.where(finiti, v, peggiore))      # valori mancanti (es. payback mai raggiunto) = peggiori
        X = np.column_stack(X)
        lo, hi = X.min(axis=0), X.max(axis=0)
        Z = (X - lo) / np.where(hi - lo > 0, hi - lo, 1.0)
        w = np.asarray(pesi, dtype=float)
        w = w / w.sum() if w.sum() > 0 else np.full(len(criteri), 1 / len(criteri))
        distanza = np.sqrt((Z ** 2 * w).sum(axis=1))

        fronte = np.zeros(len(df), dtype=int)
        rimasti = np.arange(len(df))
        k = 0
        while len(rimasti) and k < 50:
            k += 1
            sub = X[rimasti]
            non_dom = [a for a in range(len(sub))
                       if not (np.all(sub <= sub[a], axis=1) & np.any(sub < sub[a], axis=1)).any()]
            fronte[rimasti[non_dom]] = k
            rimasti = np.delete(rimasti, non_dom)
            if (fronte > 0).sum() >= max(self.n_progetti, 1):
                break                                    # bastano i fronti che servono per i top N
        fronte[fronte == 0] = k + 1
        out = df.copy()
        out["fronte"] = fronte
        out["distanza_ideale"] = distanza
        return out.sort_values(["fronte", "distanza_ideale"], kind="stable")

    def tabella_topN_per(self, attributo1, n=10):
        attributo1 = ALIAS_VARIABILI.get(attributo1, attributo1)
        candidati = self._candidati()
        if len(self.criteri) >= 2:
            ordinata = self.ordina_multicriterio(candidati, self.criteri, self.pesi)
            top = ordinata.iloc[:n]
            # frontiera (fronte 1) di tutte le configurazioni ammissibili, per i grafici
            self.risultati["pareto_criteri"] = self.risultati["ID"].isin(ordinata.loc[ordinata["fronte"] == 1, "ID"])
        else:
            base, self.risultati = self.risultati, candidati
            try:
                ordinata = self.ordina_lista(attributo1)
            finally:
                self.risultati = base
            top = ordinata.iloc[-n:].iloc[::-1]          # dal migliore (Progetto 1) in giù
        self.top_righe = top.reset_index(drop=True)
        self.top_idx = top["ID"].astype(int).tolist()
        self._aggiorna(0.92, "Business plan dei migliori progetti" if self.lingua == "ITA" else "Top projects")

        self.top_progetti, self.top_andamenti = [], []
        for idx, prezzo in zip(self.top_idx, top["prezzoindrogeno"]):
            a = self._analisi_finanziaria(idx, prezzo=prezzo if attributo1 in CRITERI_PREZZO else None)
            a.RUN(tabelle=True)
            self.top_progetti.append(a)
            self.top_andamenti.append(self.analisi1.andamenti_progetto(idx))

        self.valori_simulatore = [list(top["CapFac"]), list(top["ImpPV1"]), list(top["PotEle"]),
                                  list(top["EnergiaAutocons"]), list(top["ProdAnnuaIdrogkg"]),
                                  list(top["ProdElettVend"]), list(top["AccuE"])]

        # flussi energetici dei migliori progetti (stesse righe dell'originale)
        self.dfs_flussi_energetici = []
        for el in self.top_andamenti:
            data = {
                'Energia in elettrolizzatore per ogni ora': el[0],
                'Energia immessa in rete per ogni ora': el[2],
                'Energia prodotta dal fotovoltaico': el[3],
                'Potenza minima elettrolizzatore': el[4],
                'Potenza massima elettrolizzatore': el[5],
            }
            if self.batteria == "SI":
                data['Energia in batteria ora per ora'] = el[6]
                data['Potenza massima batteria'] = el[7]
                data['Energia disponibile in batteria'] = el[8]
                data['Energia disponibile'] = el[9]
                data['Max erogabile'] = el[10]
            self.dfs_flussi_energetici.append(pd.DataFrame(data).transpose())

        ita = self.lingua != "ENG"
        data = {
            ('VAN Progetto' if ita else 'Project NPV'): list(top["VAN"]),
            ('TIR Progetto' if ita else 'Project IRR'): list(top["TIR"]),
            ('PAYBACK Progetto [anni]' if ita else 'Project PAYBACK [years]'): list(top["PAYBACK"]),
            ('LCOH attualizzato [€/kg]' if ita else 'Discounted LCOH [€/kg]'): list(top["LCOH"]),
            'Capacity Factor [%]': list(top["CapFac"]),
            ('Potenza impianto PV [kW]' if ita else 'PV Plant Power [kW]'): list(top["ImpPV1"]),
            ('Potenza Elettrolizzatore [kW]' if ita else 'Electrolyzer Power [kW]'): list(top["PotEle"]),
            ('Potenza Compressore [kW]' if ita else 'Compressor Power [kW]'): list(top["potenza_compressore"]),
            ('Energia Autoconsumata [kWh]' if ita else 'Self-Consumed Energy [kWh]'): list(top["EnergiaAutocons"]),
            ('Autoconsumo [%]' if ita else 'Self-Consumption [%]'): list(top["Autoconsumo"]),
            ('Produzione idrogeno [kg] (primo anno)' if ita else 'Hydrogen Production [kg] (first year)'):
                list(top["ProdAnnuaIdrogkg"]),
            ('Elettricità immessa in rete [kWh]' if ita else 'Electricity Injected into the Grid [kWh]'):
                list(top["ProdElettVend"]),
        }
        if self.batteria == "SI":
            data['Dimensione Batteria [kWh]' if ita else 'Battery Size [kWh]'] = list(top["AccuE"])
        if attributo1 in CRITERI_PREZZO:
            data['Prezzo idrogeno [€]' if ita else 'Hydrogen Price [€]'] = list(top["prezzoindrogeno"])
        self.TabellaMax = pd.DataFrame(data).transpose()
        self.TabellaMax.columns = [f'{"Progetto" if ita else "Project"} {i + 1}' for i in range(len(top))]

        self.dfs_conto_economico = [p.dfContoEconomico for p in self.top_progetti]
        self.dfs_indici_finanziari = [p.dfIndiciFinanziari for p in self.top_progetti]
        self.dfs_flussi_monetari = [p.dfFlussiMonetari for p in self.top_progetti]

    def simulatore_simulatore(self, attributo):
        """Prezzo (o incentivo pubblico) di equilibrio per i migliori progetti (VAN > 0)."""
        variabile = "incentivo" if attributo in SIM_INCENTIVO else "prezzo"
        self.prezzi_eq, self.VAN_eq = [], []
        for i, idx in enumerate(self.top_idx):
            valore, a = self._cerca_equilibrio(idx, variabile, stretto=True)
            self.prezzi_eq.append(valore)
            self.VAN_eq.append(a.VAN if a is not None else None)
            self._aggiorna(0.95 + 0.05 * (i + 1) / len(self.top_idx),
                           "Equilibrio" if self.lingua == "ITA" else "Equilibrium")

    def crea_grafico_relazione(self, attributo1, attributo2):
        df = self.ordina_lista2(attributo1, attributo2)
        return (df[ALIAS_VARIABILI.get(attributo1, attributo1)].tolist(),
                df[ALIAS_VARIABILI.get(attributo2, attributo2)].tolist())

    def crea_grafico_SA(self, attributo1, attributo2, attributo3):
        df = self.ordina_lista2(attributo1, attributo3)
        a1, a2, a3 = (ALIAS_VARIABILI.get(x, x) for x in (attributo1, attributo2, attributo3))
        return df[a1].tolist(), df[a2].tolist(), df[a3].tolist()

    # ---------------------------------------------------------------------------------------------
    # colori per i grafici di sensitivity (Gradients 1 -> Gradients 2)
    def interpolate_color(self, start_color, end_color, factor):
        return [int(start_color[i] + (end_color[i] - start_color[i]) * factor) for i in range(3)]

    def rgb_to_hex(self, rgb):
        return "#{:02x}{:02x}{:02x}".format(*rgb)

    def hex_to_rgb(self, hex_color):
        hex_color = hex_color.lstrip("#")
        return tuple(int(hex_color[i:i + 2], 16) for i in (0, 2, 4))

    def scarica_risultati(self):
        """Costruisce il file Excel di output (stesso formato di OUTPUT.xlsx) e ne restituisce i byte."""
        buffer = io.BytesIO()
        workbook = xlsxwriter.Workbook(buffer, {'nan_inf_to_errors': True, 'in_memory': True})
        ita = self.lingua != "ENG"
        titolo = 'Sommario generale' if ita else "General Summary"
        worksheet1 = workbook.add_worksheet(titolo)
        color_palette = {
            "title_bg": "#5F85EA", "header_bg": "#dee4f2", "cell_bg": "#f8f8f8", "gridlines": "#666666",
            "chart_fill1": "#d03938", "chart_fill2": "#e0b544", "font_color": "#282a2b", "white_color": "#FFFFFF",
            "negative_red": "#aa312e", "header_bg2": "#95b3d7", "Elements_1": "#8f60fe",
            "Gradients_1": "#07d5df", "Gradients_2": "#f407fe", "Text_2": "#4c4c54", "illustrations_4": "#e7cc65",
        }
        title_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter',
                                            'bg_color': color_palette['title_bg'],
                                            'font_color': color_palette["white_color"],
                                            'font_name': 'Rajdhani Medium', 'font_size': 22})
        subtitle_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter',
                                               'bg_color': color_palette['title_bg'],
                                               'font_color': color_palette["white_color"],
                                               'font_name': 'Rajdhani Medium', 'font_size': 16})
        header_format = workbook.add_format({'bold': True, 'text_wrap': True, 'align': 'center', 'valign': 'vcenter',
                                             'bg_color': color_palette['header_bg'],
                                             'font_color': color_palette['font_color'], 'border': 1,
                                             'border_color': color_palette['gridlines'],
                                             'font_name': 'Rajdhani Medium', 'font_size': 11})
        generic_cell_format = workbook.add_format({'bg_color': color_palette['cell_bg'],
                                                   'font_color': color_palette['font_color'], 'border': 1,
                                                   'border_color': color_palette['gridlines']})
        number_format = workbook.add_format({'num_format': '#,##0.00;[Red]-#,##0.00;0.00',
                                             'bg_color': color_palette['cell_bg'],
                                             'font_color': color_palette['font_color'], 'border': 1,
                                             'border_color': color_palette['gridlines'],
                                             'font_name': 'Roboto', 'font_size': 11})

        def scrivi(ws, r, c, v, fmt):
            if v is None or (isinstance(v, float) and np.isnan(v)):
                ws.write_blank(r, c, None, fmt)
            elif isinstance(v, (int, float, np.integer, np.floating)):
                ws.write_number(r, c, float(v), fmt)
            else:
                ws.write(r, c, v, fmt)

        worksheet1.set_column('A:A', 26.5, generic_cell_format)
        worksheet1.set_column('B:XFD', 12, generic_cell_format)
        numero_colonne = len(self.TabellaMax.columns)
        worksheet1.merge_range(0, 0, 0, max(numero_colonne, 1), titolo, title_format)
        for idx, col in enumerate(self.TabellaMax.columns):
            worksheet1.write(2, idx + 1, col, header_format)
        for row_idx, index in enumerate(self.TabellaMax.index):
            worksheet1.write(row_idx + 3, 0, index, header_format)
            for col_idx, value in enumerate(self.TabellaMax.iloc[row_idx]):
                scrivi(worksheet1, row_idx + 3, col_idx + 1, value, number_format)

        start_row = len(self.TabellaMax.index) + 5
        if self.relazione in ["SI", "YES"] and self.attributo not in CRITERI_PREZZO:
            for el1, el2 in zip(self.traduci_lista_per_il_codice(self.variable_1_list),
                                self.traduci_lista_per_il_codice(self.variable_2_list)):
                valori_x, valori_y = self.crea_grafico_relazione(el1, el2)
                if not valori_x:
                    continue
                n1, n2 = self.aggiungi_unita_di_misura(el1), self.aggiungi_unita_di_misura(el2)
                for col_idx, value in enumerate(valori_x):
                    scrivi(worksheet1, start_row, col_idx, value, generic_cell_format)
                for col_idx, value in enumerate(valori_y):
                    scrivi(worksheet1, start_row + 1, col_idx, value, generic_cell_format)
                chart = workbook.add_chart({'type': 'scatter', 'subtype': 'straight_with_markers'})
                chart.add_series({
                    'name': f'{n1} vs {n2}',
                    'categories': [worksheet1.name, start_row, 0, start_row, len(valori_x) - 1],
                    'values': [worksheet1.name, start_row + 1, 0, start_row + 1, len(valori_y) - 1],
                    'marker': {'type': 'circle', 'size': 2, 'fill': {'color': color_palette['chart_fill1']},
                               'border': {'color': color_palette['chart_fill1']}},
                    'line': {'none': True}})
                font_titolo = {'bold': True, 'size': 18, 'color': color_palette["font_color"], 'name': 'Rajdhani Medium'}
                font_asse = {'bold': True, 'size': 13, 'color': color_palette["font_color"], 'name': 'Rajdhani Medium'}
                font_num = {'name': 'Roboto', 'size': 9, 'color': color_palette["font_color"]}
                chart.set_title({'name': f'Scatter Plot: {n1} vs {n2}', 'name_font': font_titolo})
                vx = [v for v in valori_x if v is not None and np.isfinite(v)]
                vy = [v for v in valori_y if v is not None and np.isfinite(v)]
                ax = {'name': n1, 'major_gridlines': {'visible': True, 'line': {'color': color_palette['gridlines']}},
                      'name_font': font_asse, 'num_font': font_num}
                ay = {'name': n2, 'major_gridlines': {'visible': True, 'line': {'color': color_palette['gridlines']}},
                      'name_font': font_asse, 'num_font': font_num}
                if vx:
                    ax.update({'min': min(vx) - abs(min(vx) * 0.1), 'max': max(vx) + abs(max(vx) * 0.1)})
                if vy:
                    ay.update({'min': min(vy) - abs(min(vy) * 0.1), 'max': max(vy) + abs(max(vy) * 0.1)})
                chart.set_x_axis(ax)
                chart.set_y_axis(ay)
                chart.set_legend({'position': 'bottom'})
                worksheet1.insert_chart(f'A{start_row + 3}', chart,
                                        {'x_offset': 5, 'y_offset': 10, 'x_scale': 1.5, 'y_scale': 1.5})
                start_row += 23

        if self.si_fa_grafico_SA in ["SI", "YES"] and self.batteria == "SI" and self.attributo not in CRITERI_PREZZO:
            k_pos, j_pos = start_row + 1, 8
            for el1 in self.traduci_lista_per_il_codice(self.SA_variable_list):
                valori_x, valori_y, valori_z = self.crea_grafico_SA("PotEle", "AccuE", el1)
                if not valori_z:
                    continue
                unique_x = sorted(set(valori_x))
                unique_y = sorted(set(valori_y))
                max_el1_pot_ele = valori_x[valori_z.index(max(valori_z))]
                filtered_pot_ele = [max_el1_pot_ele] + [unique_x[i] for i in range(0, len(unique_x), max(1, len(unique_x) // 16))
                                                        if unique_x[i] != max_el1_pot_ele]
                lookup = {(x, y): z for x, y, z in zip(valori_x, valori_y, valori_z)}  # al posto della ricerca lineare
                for col_idx, pot_ele_value in enumerate(filtered_pot_ele, start=1):
                    worksheet1.write(start_row, col_idx, round(pot_ele_value, 2), generic_cell_format)
                for row_idx, accu_e_value in enumerate(unique_y, start=start_row + 1):
                    worksheet1.write(row_idx, 0, round(accu_e_value, 2), generic_cell_format)
                    for col_idx, pot_ele_value in enumerate(filtered_pot_ele, start=1):
                        scrivi(worksheet1, row_idx, col_idx, lookup.get((pot_ele_value, accu_e_value)),
                               generic_cell_format)
                chart = workbook.add_chart({'type': 'scatter', 'subtype': 'straight_with_markers'})
                for i, nomecateg in enumerate(filtered_pot_ele):
                    factor = nomecateg / max(filtered_pot_ele)
                    color = self.rgb_to_hex(self.interpolate_color(self.hex_to_rgb(color_palette["Gradients_1"]),
                                                                   self.hex_to_rgb(color_palette["Gradients_2"]), factor))
                    chart.add_series({
                        'name': f'{round(nomecateg, 1)}',
                        'categories': [worksheet1.name, start_row + 1, 0, start_row + len(unique_y), 0],
                        'values': [worksheet1.name, start_row + 1, i + 1, start_row + len(unique_y), i + 1],
                        'marker': {'type': 'circle', 'size': 4, 'fill': {'color': color_palette['font_color']},
                                   'border': {'color': color_palette['font_color']}},
                        'line': {'width': 2, 'color': color}})
                nome = self.aggiungi_unita_di_misura(el1)
                chart.set_title({'name': f'Sensitivity analysis chart: {nome} vs Battery size by electrolizer size',
                                 'name_font': {'bold': True, 'size': 18, 'color': color_palette["font_color"],
                                               'name': 'Rajdhani Medium'}})
                chart.set_x_axis({'name': "Capacità accumulo [kWh]" if ita else "Battery size [kWh]",
                                  'major_gridlines': {'visible': True, 'line': {'color': color_palette['gridlines']}},
                                  'name_font': {'bold': True, 'size': 13, 'color': color_palette["font_color"],
                                                'name': 'Rajdhani Medium'},
                                  'num_font': {'name': 'Roboto', 'size': 9, 'color': color_palette["font_color"]}})
                chart.set_y_axis({'name': nome,
                                  'major_gridlines': {'visible': True, 'line': {'color': color_palette['gridlines']}},
                                  'name_font': {'bold': True, 'size': 13, 'color': color_palette["font_color"],
                                                'name': 'Rajdhani Medium'},
                                  'num_font': {'name': 'Roboto', 'size': 9, 'color': color_palette["font_color"]}})
                chart.show_blanks_as('span')
                chart.set_legend({'position': 'right', 'title': {'name': "Electrolyzer size"}})
                worksheet1.insert_chart(k_pos, j_pos, chart, {'x_offset': 5, 'y_offset': 10, 'x_scale': 3, 'y_scale': 2.5})
                start_row += len(unique_y) + 1
                k_pos += 37

        # fogli dei singoli progetti: PA 1 = migliore
        for i in range(len(self.dfs_conto_economico)):
            project_num = i + 1
            worksheet = workbook.add_worksheet(f"PA {project_num}")
            worksheet.freeze_panes(0, 2)
            worksheet.merge_range('A1:B1', f'Project Analysis {project_num}', title_format)
            worksheet.set_column('A:A', 8, generic_cell_format)
            worksheet.set_column('B:B', 54, generic_cell_format)
            worksheet.set_column('C:XFD', 13, generic_cell_format)
            worksheet.set_row(1, 63)
            ind_fin = self.dfs_indici_finanziari[i]
            for col_num, value in enumerate(ind_fin.columns):
                worksheet.write(1, col_num + 1, value, header_format)
            for row_num, row in enumerate(ind_fin.values):
                for col_num, value in enumerate(row):
                    if isinstance(value, np.ndarray):
                        value = value.item() if value.size == 1 else str(value.tolist())
                    scrivi(worksheet, row_num + 2, col_num + 1, value, number_format)

            if self.si_fa_simulazione in ["SI", "YES"] and self.attributo not in CRITERI_PREZZO and self.prezzi_eq:
                if self.attributo_simulazione in SIM_INCENTIVO:
                    worksheet.write('V2', "L'incentivo di eq:" if ita else "Eq incentive is:", header_format)
                else:
                    worksheet.write('V2', 'il prezzo di eq:' if ita else "Eq price is:", header_format)
                scrivi(worksheet, 2, 21, self.prezzi_eq[i], number_format)

            start_row = 4
            worksheet.merge_range(start_row, 0, start_row, 1, "Conto Economico" if ita else "Income Statement",
                                  subtitle_format)
            cont_eco = self.dfs_conto_economico[i]
            for col_num, value in enumerate(cont_eco.columns):
                worksheet.write(start_row, col_num + 2, value, header_format)
            for row_idx, (index, row) in enumerate(cont_eco.iterrows()):
                worksheet.write(row_idx + start_row + 1, 1, index, header_format)
                for col_idx, value in enumerate(row):
                    scrivi(worksheet, row_idx + start_row + 1, col_idx + 2, float(value), number_format)

            start_row += len(cont_eco) + 2
            worksheet.merge_range(start_row, 0, start_row, 1, "Flussi di Cassa" if ita else "Cash Flows",
                                  subtitle_format)
            flussi_mon = self.dfs_flussi_monetari[i]
            for col_num, value in enumerate(flussi_mon.columns):
                worksheet.write(start_row, col_num + 2, value, header_format)
            for row_idx, (index, row) in enumerate(flussi_mon.iterrows()):
                worksheet.write(row_idx + start_row + 1, 1, index, header_format)
                for col_idx, value in enumerate(row):
                    try:
                        value = float(value)
                    except ValueError:
                        pass
                    scrivi(worksheet, row_idx + start_row + 1, col_idx + 2, value, number_format)

            # flussi energetici orari: riga 47 in poi, come nell'originale
            fl = self.dfs_flussi_energetici[i]
            righe = [
                ('Energia in elettrolizzatore per ogni ora',
                 "Energia usata dall'elettrolizzatore [kWh]" if ita else "Energy used by electrolyzer [kWh]"),
                ('Energia immessa in rete per ogni ora',
                 "Energia immessa in rete [kWh]" if ita else "Energy fed into the grid [kWh]"),
                ('Energia prodotta dal fotovoltaico', "Prodotta FER [kWh]" if ita else "RES production [kWh]"),
                ('Potenza minima elettrolizzatore',
                 "Potenza minima elettrolizzatore [kW]" if ita else "Electrolyzer cutoff power [kW]"),
                ('Potenza massima elettrolizzatore',
                 "Taglia elettrolizzatore [kW]" if ita else "Electrolyzer nominal power [kW]"),
            ]
            if self.batteria == "SI":
                righe += [('Energia in batteria ora per ora',
                           "Energia in batteria al netto dell'efficienza di scarica [kWh]" if ita
                           else "Energy in battery net of discharge efficiency [kWh]"),
                          ('Potenza massima batteria', "Taglia batteria [kW]" if ita else "Battery nominal power [kW]")]
            n_ore = fl.shape[1]
            titolo_fl = "Flussi energetici orari" if ita else "Hourly Energy Flows"
            worksheet.merge_range(46, 0, 46, 1, titolo_fl, subtitle_format)
            worksheet.write("B47", titolo_fl, header_format)
            worksheet.write_row("C47", list(range(n_ore)), header_format)
            for r, (chiave, etichetta) in enumerate(righe):
                worksheet.write(47 + r, 1, etichetta, header_format)
                worksheet.write_row(47 + r, 2, [float(v) for v in fl.loc[chiave].tolist()])

            sheet_name = f"'PA {project_num}'"
            chart = workbook.add_chart({'type': 'line'})
            ultima = xlsxwriter.utility.xl_col_to_name(2 + n_ore - 1)
            serie = [(3, '#92D050', None), (2, '#FFC000', None), (1, '#00AF50', None),
                     (5, '#767171', 'dash'), (4, '#767171', 'dash')]
            if self.batteria == "SI":
                serie += [(6, 'orange', None), (7, 'orange', 'dash')]
            for r, colore, tratto in serie:
                riga_excel = 48 + r - 1
                linea = {'color': colore}
                if tratto:
                    linea.update({'dash_type': tratto, 'transparency': 50})
                chart.add_series({'name': f"={sheet_name}!$B${riga_excel}",
                                  'categories': f"={sheet_name}!$C$47:${ultima}$47",
                                  'values': f"={sheet_name}!$C${riga_excel}:${ultima}${riga_excel}",
                                  'line': linea})
            chart.set_title({'name': 'Grafico dei Flussi Energetici' if ita else "Energy Flow Chart", 'align': 'left',
                             'name_font': {'name': 'Rajdhani Medium', 'size': 18, 'bold': True,
                                           'color': color_palette["font_color"]}})
            chart.set_x_axis({'name': 'Ora' if ita else "Hour"})
            chart.set_y_axis({'name': 'Quantità' if ita else "Quantity"})
            chart.set_size({'width': 100000, 'height': 700})
            chart.set_legend({'position': 'left'})
            worksheet.insert_chart('B54' if self.batteria == "NO" else 'B58', chart)

        workbook.close()
        return buffer.getvalue()
