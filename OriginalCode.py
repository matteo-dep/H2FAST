import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import xlsxwriter
import warnings
import sys

class Analisi_tecnica:
    def __init__(self, file_csv, tipo_file, p_PV, dP_el, batteria, dP_bat, min_batt, max_batteria, lingua, eff_batt,
                 min_elet):
        # DATI ESOGENI
        self.batteria = batteria  # è una variabile di input ("SI" o  "NO")
        self.file_csv = file_csv + ".csv"
        self.tipo_file = tipo_file  # se il file proviene o meno da pvgis
        self.p_PV = p_PV  # taglia dell'impianto PV
        self.dP_el = dP_el  # granularità nella ricerca dell'elettrolizzatore
        self.dP_bat = dP_bat  # granularità nella ricerca delle batterie
        self.E_PV = self.load_data()  # così apro il file (guarda cosa fa la funzione load data)
        self.min_batt = min_batt  # definisco il minimo e il
        self.max_batteria = max_batteria  # massimo erogabile in valore relativo dalla batteria
        self.min_elet = min_elet  # potenza minima dell'elettrolizzatore
        self.e_pv = np.sum(self.E_PV)  # così ottengo l'energia totale prodotta in un'anno
        # essendo la granularità nella ricerca dell'elettrolizzatore una variabile che influenza enormemente il tempo computazionale nella simulazione considero una ricerca minore per taglie di fotovoltaico maggiori di 1 Mw
        """if self.p_PV > 1000: # se l'impianto PV è maggiore di 1 MW
            self.P_elc = np.arange(10, p_PV*2+1, dP_el*2) # comincio con un'elettrolizzatore di 10 kW fino a uno il doppio dellla potenza del fotovoltaico con una granularità il doppio maggiore rispetto alla simmulazione con l'impianto inferiore di 1 mW
        if self.p_PV <= 1000:
            self.P_elc = np.arange(10, p_PV*2+1, dP_el) # essendo assente il *2 che moltiplica dP_el; significa che controllo il doppio degli elettrolizzatori"""
        gran_elc = 0
        if self.p_PV <= 100:
            gran_elc = 10
        elif 100 < self.p_PV <= 500:
            gran_elc = 20
        elif 500 < self.p_PV <= 1000:
            gran_elc = 50
        elif self.p_PV > 1000:
            gran_elc = 100
        self.P_elc = np.linspace(self.p_PV / gran_elc, self.p_PV*dP_el, gran_elc)

        self.lingua = lingua  # qui abbiamo la lingua dell'output
        self.eff_batt = eff_batt  # qui abbiamo l'efficienza della batteria
        # DATI ENDOGENI:
        # questi vanno inizializzati (ovvero devo allocare un spazio nella memoria a queste variabili)
        # Sto utilizzando le array di numpy per questioni di memoria; il problema è che devo già determinare la loro dimensioni
        # inoltre l'inizializzazione è diversa nel caso l'analisi venga effettuata con o senza batteria
        if self.batteria == "NO":
            # in questo caso io faccio l'analisi solo per ogni taglia di elettrolizzatore ovvero la lugnhezza di P_elc
            self.qt_progetti = len(self.P_elc)  # questo verttore mi conta quanti progetti vado ad analizzare
            self.andamenti = np.zeros((self.qt_progetti, 6,len(self.E_PV)))  # qui verranno salvati i flussi energetici in caso di assenza della batteria
            self.E_H2 = np.zeros(self.qt_progetti)  # qui c'è l'energia totale che è entrata nell'elettrolizzatore per ogni progetto
            self.E_im = np.zeros(self.qt_progetti)  # qui c'è l'energia totale immessa in rete per ogni progetto
            self.M_H2 = np.zeros(self.qt_progetti)  # qui c'è la massa di igrogeno prodotto per ogni progetto
            self.Auto = np.zeros(self.qt_progetti)  # definisco il vettore autoconsumo con il suo rispettivo valore per ogni progetto
            self.CF = np.zeros(self.qt_progetti)  # vettore Capacity Factor per progetto
            self.OFFgiorn = np.zeros(self.qt_progetti)  # servirà a capire il numero di spegnimento per progetto; considera come spegnimento anche il caso in cui l'elettrolizzatore non ha lavorato
            self.OFF = np.zeros(self.qt_progetti)  # numero di spegnimenti in totale per progetto senza considerare i casi in cui non lavora per più di 24 ore
            self.spegn_giorn = np.zeros(self.qt_progetti)
            # inizializzazione in caso di presenza delle batteria
        # con la presenza delle batterie l'analisi diventa molto più comlessa essendo che devo fare l'analisi per ogni taglia di elettrolizzatore e batteria
        if self.batteria == "SI":
            len_batt = 0  # contiamo il numero di progetti
            for el in self.P_elc:  # per ogni taglia di elettrolizzatore
                # anche in questo caso la ricerca parte de una batteria di grandezza pari a 0 fino al doppio della potenza dell'elettrolizzatore
                """if self.p_PV > 1000: # granularità è doppia se il PV è maggiore di 1 MW
                    P_batt = np.arange(0,2*el + 1,self.dP_bat*2)
                if self.p_PV <= 1000:
                    P_batt = np.arange(0,2*el + 1,self.dP_bat)
                len_batt += len(P_batt) # in questo loop vado a determinare il numero di progetti in caso di presenza della batteria"""
                gran_bat = 0
                if el <= 100:
                    gran_bat = 10
                elif 100 < el <= 500:
                    gran_bat = 20
                elif 500 < el <= 1000:
                    gran_bat = 50
                elif el > 1000:
                    gran_bat = 100
                P_batt = np.linspace(el * self.dP_bat / gran_bat, el * self.dP_bat, gran_bat)
                len_batt += len(P_batt)  # in questo loop vado a determinare il numero di progetti in caso di presenza della batteria

            self.qt_progetti = len_batt
            self.andamenti = np.zeros((self.qt_progetti, 11,len(self.E_PV)))  # qui verrranno salvati i flussi energetici in caso di assenza della batteria
            self.E_H2 = np.zeros(self.qt_progetti)  # qui c'è l'energia in elettrolizzatore per ogni progetto
            self.E_im = np.zeros(self.qt_progetti)  # qui c'è l'energia immessa per ogni progetto
            self.M_H2 = np.zeros(self.qt_progetti)  # qui c'è la massa di igrogenoprodotto per ogni progetto
            self.Auto = np.zeros(self.qt_progetti)  # definisco il vettore autoconsumo per progetto
            self.CF = np.zeros(self.qt_progetti)  # vettore Capacity Factor per progetto
            self.OFFgiorn = np.zeros(self.qt_progetti)  # servirà a capire il numero di spegnimento per progetto
            self.OFF = np.zeros(self.qt_progetti)  # numero di spegnimenti in totale per progetto
            self.potenza_batt = np.zeros(self.qt_progetti)  # qui ci saranno tutti i valori delle batterie per progetto
            self.potenza_elett = np.zeros(self.qt_progetti )  # qui ci sono invece tutti i valori delle taglie di elettrolizzatore per progetto
            self.spegn_giorn = np.zeros(self.qt_progetti)  # queste tre variabili mi servono per andare a capire quante volte si spegne durante il giorno (quando c'è giorno)
        # Variabili che gestiscono gli spegnimenti
        self.off = 0
        self.flag = 0
        self.count_to_24 = 0
        self.count2 = 0

    def load_data(self):
        # qui definisco le logiche per leggere i file per i dati orari
        if self.tipo_file == "SI":
            with open(self.file_csv, 'r') as file:lines = file.readlines()
            header = "time,P,G(i),H_sun,T2m,WS10m,Int"  # Definire l'intestazione esatta desiderata (quella derivante da PVGIS)
            header_idx = next((i for i, line in enumerate(lines) if line.strip() == header),None)  # Trova l'indice della riga di intestazione che corrisponde esattamente
            if header_idx is None:
                if self.lingua == "ITA":
                    raise ValueError(
                        "Errore nell'apertura del file PVGIS: scarica un nuovo file PVGIS")  # se non la trova dimmi che c'è un errore nel file PVGIS che hai mandato
                elif self.lingua == "ENG":
                    raise ValueError("Error opening the PVGIS file: download a new PVGIS file.")
            rows = []
            for line in lines[header_idx + 1:]:  # salva tutti i valori che sono sotto l'intestazione
                columns = line.strip().split(',')
                if len(columns) == 7:  # so che il numero di colonne è sempre uguale a 7
                    rows.append(columns)
                else:
                    break  # Interrompe la lettura se incontra una riga con un numero di colonne differente
            data = pd.DataFrame(rows, columns=header.split(','))  # qui creo il dataframe che da cui nella riga successiva ottengo i valori della produzione da PV
            E_PV = pd.to_numeric(data['P'], errors='coerce').dropna()  #
            E_PV /= 1000  # Converti da W a kW
            return E_PV.to_numpy()
        if self.tipo_file == "NO":  # in caso di file non proveniente da PVGIS; vogliamo che questo abbia un formato tale che sia leggibile dal nostro modulo
            data = pd.read_csv(self.file_csv, sep=';',decimal=',')  # in questo caso una sola colonna in cui interno ci siano 8760 valori della produzione annua da PV
            E_PV = pd.to_numeric(data.iloc[:, 0],errors='coerce').dropna()  # Noi pretendiamo che il file sia già definito in kW
            return E_PV.to_numpy()

    @staticmethod
    def eff_elc(x):  # definisco la funzione che esprime la variazione di efficienza (kWhel per kg H2 prodotto da idrolizzatore) al variare della Potenza di esercizio rispetto al nominale
        y = (-6.1371 * x ** 6 + 24.394 * x ** 5 - 39.663 * x ** 4 + 33.988 * x ** 3 - 16.412 * x ** 2 + 4.2929 * x + 0.1022)
        return y

    def run_analysis_nobattery(self):  # Qui troviamo le legiche della simulazione dei flussi energetici nel caso di assenza di batteria
        for index, p_elc in enumerate(self.P_elc):  # per ogni taglia dell'eletrolizzatore salvo il suo indice e il suo valore
            p_elc_min = self.min_elet * p_elc  # Potenza minima di soglia per entrata in lavoro dell'elettrolizzatore
            E_H2_h = np.zeros(len(self.E_PV))  # vettore energia autoconsumata per produrre idrogeno per ogni ora dell'anno
            M_H2_h = np.zeros(len(self.E_PV))  # vettore massa idrogeno prodotto per ogni ora dell'anno
            E_im_h = np.zeros(len(self.E_PV))  # vettore energia immessa in rete per ogni ora dell'anno
            P_pv_h = np.zeros(len(self.E_PV))  # vettore energia prodotta dal fotovoltaico per ogni ora dell'anno
            Pot_min_elett = np.full(len(self.E_PV),p_elc_min)  # questi sono dei vettori che hanno sempre il valore minimo per far funzionare l'elttrolizzatore (servono per essere disegnati)
            Pot_max_elett = np.full(len(self.E_PV),p_elc)  # questi sono dei vettori che hanno sempre il valore massimo per far funzionare l'elttrolizzatore (servono per essere disegnati)

            # variabili per lo spegnimento
            self.off = 0  # per il momento l'elttrolizzatore non si è mai spento quindi il numero di spegnimenti è uguale a zero
            self.flag = 1  # l'elttrolizzatore comincia da spento
            self.count_to_24 = 0  # qeusta variabile mi serve per capire se l'elettrolizzatore ha mai non lavorato per un numero maggiore di 24 ore; ovvero è stato spento tutto il giorno
            self.count2 = 0  # qui conto i casi in cui l'lettrolizzatore non ha lavorato per un intero giorno

            for i, p_pv in enumerate(self.E_PV):  # qui
                if p_pv > p_elc:  # se la potenza derivante da PV è maggiore della potenza massima dell'elettrolizzaatore
                    e_H2 = p_elc  # produco idrogeno al massimo della potenza dell'elettrolizzatore
                    e_im = p_pv - e_H2  # l'energia immessa è la differenza tra quella che produco e quella che utilizzo
                    m_H2 = e_H2 * 0.565 / (120 / 3.6)  # così ottengo la massa di idrogeno; 0.565 = self.eff_elc(1)

                    # noinspection PyTypeChecker
                    self.flag = 0  # qui l'elettrolizzatore è acceso perché produco
                    self.count_to_24 = 0  # sicuro ha prodotto oggi

                elif p_elc_min <= p_pv < p_elc:  # se la potenza derivante da PV è minore della potenza massima e maggiore della potenza minima dell'elettrolizzaatore
                    e_H2 = p_pv  # produco idrogeno non al massimo ma in base alla prodotta in quest'ora dal fotovoltaico
                    e_im = 0  # non immetto energia nell'elettrolizzatore
                    var = p_pv / p_elc  # questa variabile è utilizzata per capire quanto produce l'elettolizzatore se non lavora al massimmo
                    eta = self.eff_elc(var)
                    m_H2 = e_H2 * eta / (120 / 3.6)  # ottengo la massa di idrogneo

                    self.flag = 0  # qui l'elettrolizzatore è a lavoro
                    self.count_to_24 = 0  # qui l'elettrolizzatore sta riposando. fa le ninne e le nanne

                elif p_pv <= p_elc_min:  # in questo caso non produco se l'energia da PV è inferiore della potenza minima
                    e_H2 = 0  # se non produco non do energia all'elettrolizzatore
                    e_im = p_pv  # tutta l'energia che produco è immessa in rete
                    m_H2 = 0  # non ho alcuna massa di idrogeno

                    if self.flag == 0:  # se l'ora precedente era acceso
                        self.off += 1  # consideralo come spegnimento
                        self.flag = 1  # lo spengo
                    else:  # se l'ora precedente era spento
                        self.flag = 1  # continua a tenerlo spento
                    self.count_to_24 += 1  # qui aumento il conto delle ore in cui è spento
                    if self.count_to_24 == 24:  # se è spento per un giorno di fila
                        self.count2 += 1  # questa varibile conta i giorni in cui è spento tutto il giorno
                        self.count_to_24 = 0  # ricomincia a contare

                # Costruzione dei vettori energetici
                # noinspection PyUnboundLocalVariable
                E_H2_h[i] = e_H2  # vettore orario dell'energia in elettrolizzatore
                # noinspection PyUnboundLocalVariable
                M_H2_h[i] = m_H2  # vettore orario della massa di idrogeno
                # noinspection PyUnboundLocalVariable
                E_im_h[i] = e_im  # vettore orario dell'energia immessa
                P_pv_h[i] = p_pv  # vettore orario dell'energia da PV


            # calcolo delle infromazioni principali per ogni progetto
            e_H2 = np.sum(E_H2_h)  # energia in elettrolizzatore in un'anno
            e_im = np.sum(E_im_h)  # energia immessa in un'anno
            m_H2 = np.sum(M_H2_h)  # massa idrogeno prodotta in un'anno
            spegn_giorn_proj = self.count2 + self.off  #- len(self.E_PV) / 24      # qui ci sono i numeri di spegnimenti durante il giorno
            e_TOT = len(self.E_PV) * p_elc  # energia totale ovvero l'energia totale se avesse funzionato al 100% tutti i giorni
            cf = e_H2 / e_TOT * 100  # calcolo del capacity factor
            autoconsumo = e_H2 / self.e_pv * 100  # calcolo dell'autoconsumo

            # index fa a capo al progetto con determinata taglia di elettrolizzatore:
            self.OFF[index] = self.off  # numero di spegnimenti per ogni progetto
            self.spegn_giorn[index] = spegn_giorn_proj  # numero di spegnimento per ogni progetto (considerando anche i giorni in cui non ha lavorato come spegnimenti)
            self.CF[index] = cf  # capacity factor per ogni progetto
            self.Auto[index] = autoconsumo  # valore dell'autconsumo per ogni progetto
            self.E_H2[index] = e_H2  # energia utilizzata per elettrolizzare per ogni progetto
            self.E_im[index] = e_im  # energia immessa per ogni progetto
            self.M_H2[index] = m_H2  # massa di idrogeno per ogni progetto

            # qui invece ci sono i vettori dei flussi energetici per ogni progetto (servono per dopo disegnare i grafici)
            self.andamenti[index, 0, :] = E_H2_h
            self.andamenti[index, 1, :] = M_H2_h
            self.andamenti[index, 2, :] = E_im_h
            self.andamenti[index, 3, :] = P_pv_h
            self.andamenti[index, 4, :] = Pot_min_elett
            self.andamenti[index, 5, :] = Pot_max_elett

            # qui invece disegno le barre di progresso; per capire a che punto è la simulazione
            if self.lingua == "ENG":
                self.print_progress_bar(index + 1, len(self.E_H2), suffix="Technical Analysis without Battery")
            elif self.lingua == "ITA":
                self.print_progress_bar(index + 1, len(self.E_H2), suffix="Analisi Tecnica senza batteria")

    def run_analysis_battery_static_min(self):  # queste sono le logiche della produzione con batteria
        index = 0  # qui ho gli indici dei progetti
        for p_elc in self.P_elc:  # per ogni taglia di elettrolizzatore
            p_elc_min = self.min_elet * p_elc  # Potenza minima di soglia per entrata in lavoro dell'elettrolizzatore

            #vecchio snippet per popolamento vettore dimensioni accumulo
            """if self.p_PV > 1000: # aumento la granularità se la potenza del PV è maggiore di 1 mW
                P_batt = np.arange(0,2*p_elc + 1,self.dP_bat*2)# Range di valori per P_elc (val min, val max, step) in funzione di p_PV, taglia dell'impianto PV
            if self.p_PV <= 1000:
                P_batt = np.arange(0,2*p_elc + 1,self.dP_bat)"""

            if p_elc <= 100:
                gran_bat = 10
            elif 100 < p_elc <= 500:
                gran_bat = 20
            elif 500 < p_elc <= 1000:
                gran_bat = 50
            elif p_elc > 1000:
                gran_bat = 100
            P_batt = np.linspace(0, self.dP_bat * p_elc, gran_bat) #self.dP_bat * p_elc / gran_bat

            # noinspection PyUnboundLocalVariable
            for p_batt in P_batt: # per ogni taglia della batteria
                    e_batt_max = p_batt*self.max_batteria  # questa è la potenza massima erogabile dalle batterie (nel calcolo della potenza max è considerata anche l'efficienza della batteria stessa)
                    E_H2_h = np.zeros(len(self.E_PV))  # vettore energia autoconsumata per produrre idrogeno
                    M_H2_h = np.zeros(len(self.E_PV))  # vettore massa idrogeno prodotto
                    E_im_h = np.zeros(len(self.E_PV))  # vettore energia immessa in rete
                    E_batt_disponibile_h = np.zeros(len(self.E_PV)) # vettore energia disponibile in batteria
                    E_batt = np.zeros(len(self.E_PV)) # vettore energia in batteria
                    P_pv_h = np.zeros(len(self.E_PV)) # energia prodotta da fotovoltaico ora per ora
                    Energia_disp = np.zeros(len(self.E_PV)) # vettore con l'energia disponibile PV e batteria
                    Pot_min_elett = np.full(len(self.E_PV), p_elc_min) # vettore potenza minima elettrolizzatore
                    Pot_max_elett = np.full(len(self.E_PV), p_elc) # vettore potenza massima elettrolizzatore
                    Pot_max_batt = np.full(len(self.E_PV), p_batt) # max erogabile dalla batteria
                    max_erogabile = np.zeros(len(self.E_PV)) # max erogabile considerando PV e batteria
                    # variabile per lo spegnimento e per evitare che la batteria si carichi/scarichi più volte al giorno
                    self.off = 0 # ad inizio periodo gli spegnimenti sono uguali a zero
                    self.flag = 1  # indica l'elettrolizzatore spento
                    self.count_to_24 = 0 # è uno scalare (se conta fino a 24, significa che per intero un giorno l'elttrolizzatore non ha funzionato)
                    self.count2 = 0  # conta i giorni in cui l'elettrolizzatore è spento tutto il giorno
                    j = 0 # mi serve per evitare che le batterie si scarichino/carichino più di una volta all'ora



                    for i,p_pv in enumerate(self.E_PV): # per ogni ora di produzione
                        e_im = 0

                        if i == 0: # la prima volta determino che in batteria non c'è niente
                            energia_disp_batt = 0
                        elif i != 0:
                            energia_disp_batt = max(0, E_batt[i - 1] - self.min_batt*p_batt) # energia disponibile nella batteria è quella che va oltre la carica minima (ad esempio il 20%); se la carica è <= minimo, allora l'energia disponibile viene assunta come 0
                        # noinspection PyUnboundLocalVariable

                        energia_disp = p_pv + self.eff_batt*energia_disp_batt # l'energia disponibile è la somma dell'energia disponibile in batteria (considerando che in eletrrolizzatore entra solo il 95% di essa es.) e l'energia da PV


                        if p_pv == 0: # Caso 1; PV non produce, quindi la produzione se esistente viene solo da batteria

                            if energia_disp < p_elc_min: # se l'energia disponibile non è sufficiente per produrre, non produco
                                e_H2 = 0 # non produco

                                m_H2 = 0 # non produco idrogeno

                                if i == 0: # se questa è la prima ora ovvero quando l'energia all'inizio è uguale a zero
                                    E_batt[0] = 0 # qui metto 0 nel primo valore della batteria
                                    j += 1
                                elif i != 0:
                                    E_batt[i] = E_batt[i-1] # il livello nella batteria non cambia, perché la produzione del PV = 0
                                    j += 1

                                self.count_to_24 += 1 # continua a contare fino a 24
                                if self.flag == 0:  # se l'ora precedente era acceso
                                    self.off += 1
                                self.flag = 1 # qui è spento
                                if self.count_to_24 == 24: # se è già da 24 ore che non lavora
                                    self.count2 += 1 # questo considererà i giorni in cui l'impianto non ha prodotto, perché il PV ha prodotto pochissimo o non ha prodotto
                                    self.count_to_24 = 0 # qui riporta a 0 il count to 24

                            elif p_elc_min <= energia_disp < p_elc: # se invece l'energia disponibile è tra il minimo e il massimo
                                e_H2 = min(energia_disp, e_batt_max * self.eff_batt) # deve essere erogato all'elettrolizzatore il min tra questi due valori
                                # o erogo tutta l'energia disponibile; in questo caso il 95% dell'energia disponibile in batteria; oppure gli do tutta l'energia erogabile dalla batteria (il 95% del massimo della batteria)

                                var = e_H2 / p_elc  # diverso rispetto a prima, perché non lavora a max potenza l'elett
                                eta = self.eff_elc(var)
                                m_H2 = e_H2 * eta / (120 / 3.6)


                                #↓↓↓↓↓↓↓↓↓ ottimizzato il 09/12/2024 per evitare un ingresso ininfluente in un if
                                E_batt[i] = E_batt[i - 1] - e_H2 / self.eff_batt
                                j += 1

                                """if e_H2 == energia_disp and i == j: # se nell'elettrolizzatore va tutta l'energia
                                    E_batt[i] = E_batt[i -1] - energia_disp/self.eff_batt
                                    j += 1
                                
                                
                                elif e_H2 == e_batt_max and i == j: # se l'energia che esce è il max possibile
                                    E_batt[i] = E_batt[i -1] - e_batt_max/self.eff_batt # uscirà dall'elettrolizzatore il massimo possibile
                                    j += 1
                                    """


                                self.flag = 0 # qui l'elett lavora
                                self.count_to_24 = 0 # qui porta a 0 il contatore che conta fino a 24

                            elif p_elc <= energia_disp: # se ho disponibile più di quanto serve all'elettrolizzatore
                                e_H2 = min(p_elc, e_batt_max, energia_disp_batt)*self.eff_batt # simile a prima ma in questo caso non può eccedere la potenza del l'elett

                                var = e_H2 / p_elc  # come vale nel caso di elett non a max xapacità
                                eta = self.eff_elc(var)
                                m_H2 = e_H2 * eta / (120 / 3.6)  # massa idrogeno


                                E_batt[i] = E_batt[i - 1] - e_H2 / self.eff_batt

                                j += 1

                                self.flag = 0 # qui l'elett lavora
                                self.count_to_24 = 0 # qui porta il contatore che conta fino a 24 a 0


                        elif 0 < p_pv < p_elc_min: # Caso 2: se il fotovoltaico da solo non basta ad alimentare l'elettrolizzatore a minima potenza

                            # non considero il caso in cui l'energia disponibile sia uguale a 0 per definizione di questa (p_pv > 0)
                            if energia_disp < p_elc_min: # se l'energia disponibile non è abbastanza

                                e_H2 = 0 # non produco idrogeno

                                m_H2 = 0

                                j += 1

                                if i == 0: # questo è il primo caso in cui si carica la batteria
                                    E_batt[i] = min(p_pv, e_batt_max) * self.eff_batt# nella batteria entra solo il 95% dell'energia del fotovoltaico

                                elif i != 0: # se non è la prima volta dipende da quanto è carica la batteria
                                    E_batt[i] = min(  (min(p_pv, e_batt_max)*self.eff_batt + E_batt[i-1]),      p_batt )# nella batteria entra solo il 95% dell'energia del fotovolotaico


                                        #l'energia che immetto è uguale alla differenza tra la massima potenza a cui la batteria può caricare e la potenza prodotta dal fotovoltaico;
                                e_im = max(   p_pv - (E_batt[i] - E_batt[i-1])/self.eff_batt,    0)

                                self.count_to_24 += 1  # continua a contare fino a 24
                                if self.flag == 0:  # se l'ora precedente era acceso
                                    self.off += 1
                                self.flag = 1  # qui è spento
                                if self.count_to_24 == 24:  # se è già da 24 ore che non lavora
                                    self.count2 += 1  # questo considererà i giorni in cui l'impianto non ha prodotto, perché il PV ha prodotto pochissimo o non ha prodotto
                                    self.count_to_24 = 0  # qui riporta a 0 il count to 24

                                """if self.flag==0: #era già spento
                                    self.off += 1

                                self.flag = 1 # si spegne/è spento l'lettrolizzatore
                                self.count_to_24 += 1 # non sta lavorando

                                if self.count_to_24 == 24: # controlla se non lavora da più di 24 ore
                                    self.count2 += 1
                                    self.count_to_24 = 0"""

                            elif p_elc_min <= energia_disp < p_elc: # se seono in grado di produrre

                                e_H2 = min(energia_disp, e_batt_max*self.eff_batt + p_pv) # alimento eletrolizzatore con il minimo tra (tutta l'energia disponibile, cioè batt + produzione) e (la massima potenza che la batteria può fornire nell'ora + produzione)

                                var = e_H2 / p_elc  # come vale nel caso di elett non a max xapacità
                                eta = self.eff_elc(var)
                                m_H2 = e_H2 * eta / (120 / 3.6)



                                if i == j:
                                    if e_H2 == energia_disp: # scarico completamente la batteria
                                        E_batt[i] = self.min_batt*e_batt_max

                                    elif e_H2 == e_batt_max*self.eff_batt + p_pv: # scarico la batteria quanto serve
                                        E_batt[i] = E_batt[i-1] - e_batt_max

                                j += 1
                                self.flag = 0 # qui è acceso
                                self.count_to_24 = 0 # cosi come prima

                            elif p_elc <= energia_disp: # se ho energia abbastanza da dare energia all'eletrolizzatore affinchè lavori al massimo
                                e_H2 = min(e_batt_max + p_pv, p_elc) # alimento eletrolizzatore con il minimo tra (la massima potenza che la batteria può fornire nell'ora + produzione) e (la potenza nominale dell'elettrolizzatore)

                                if i==j:
                                    if e_H2 == e_batt_max + p_pv: # lavoro con la massima potenza che la batteria può fornire nell'ora + produzione
                                        var = e_H2 / p_elc  # come vale nel caso di elett non a max xapacità
                                        eta = self.eff_elc(var)
                                        m_H2 = e_H2 * eta / (120 / 3.6)

                                        E_batt[i] = E_batt[i-1] - e_batt_max

                                    elif e_H2 == p_elc:
                                        m_H2 = e_H2 * 0.565 / (120 / 3.6)
                                        E_batt[i] = E_batt[i-1] - (p_elc - p_pv)/self.eff_batt

                                j += 1

                                self.flag = 0 # qui è acceso
                                self.count_to_24 = 0 # cosi come prima


                        elif p_elc_min <= p_pv < p_elc: # Caso 3: ovvero quanto la potenza da PV è tra il minimo e il massimo

                            if p_elc_min <= energia_disp < p_elc:
                                e_H2 = min(energia_disp, e_batt_max*self.eff_batt + p_pv)

                                var = e_H2 / p_elc  # come vale nel caso di elett non a max xapacità
                                eta = self.eff_elc(var)
                                m_H2 = e_H2 * eta / (120 / 3.6)

                                if i == j:
                                    if e_H2 == energia_disp:
                                        E_batt[i] = E_batt[i-1] - (energia_disp  - p_pv) / self.eff_batt


                                    elif e_H2 == e_batt_max*self.eff_batt + p_pv:
                                        E_batt[i] = E_batt[i-1] - e_batt_max

                                j += 1

                                self.flag = 0 # qui è acceso
                                self.count_to_24 = 0 # cosi come prima

                            elif p_elc <= energia_disp:

                                e_H2 = min(e_batt_max * self.eff_batt + p_pv, p_elc)

                                if i == j:

                                    if e_H2 == e_batt_max * self.eff_batt + p_pv:
                                        var = e_H2 / p_elc  # come vale nel caso di elett non a max xapacità
                                        eta = self.eff_elc(var)
                                        m_H2 = e_H2 * eta / (120 / 3.6)

                                        E_batt[i] = E_batt[i-1] - e_batt_max



                                    elif e_H2 == p_elc:
                                        m_H2 = e_H2 * 0.565 / (120 / 3.6)

                                        E_batt[i] = E_batt[i-1] - (p_elc - p_pv)/self.eff_batt


                                j += 1

                                self.flag = 0 # qui è acceso
                                self.count_to_24 = 0 # qui porta il contatore che conta fino a 24 a 0


                        elif p_elc <= p_pv:

                            e_H2 = p_elc

                            m_H2 = e_H2 * 0.565 / (120 / 3.6)

                            # noinspection PyTypeChecker

                            #                          surplus energetico      spazio in accumulo     potenza massima batteria
                            E_batt[i] = E_batt[i-1]   +     min((p_pv - p_elc) * self.eff_batt, p_batt - E_batt[i - 1], e_batt_max)

                            e_im = p_pv - p_elc - (E_batt[i] - E_batt[i-1]) / self.eff_batt

                            j += 1

                            self.flag = 0 # qui è acceso
                            self.count_to_24 = 0 # qui porta il contatore che conta fino a 24 a 0


                        """if E_batt[i] > p_batt: # qui gestisco l'immissione in rete; infatti immetto in rete solo se la batteria è completamente carica oppure se ho un surplus che la batteria non riesce ad assorbire perchè non ha una potenza sufficiente
                            e_im += (E_batt[i] - p_batt)/self.eff_batt
                            E_batt[i] -= e_im * self.eff_batt
                        if p_elc - energia_disp > e_batt_max:
#                            e_im +="""

                        # noinspection PyUnboundLocalVariable
                        E_H2_h[i] = e_H2  # vettore orario di energia in ingresso all'elettrolizzatore
                        # noinspection PyUnboundLocalVariable
                        M_H2_h[i] = m_H2  # vettore orario della massa di energia
                        E_im_h[i] = e_im  # vettore orario dell'energia immessa in rete
                        E_batt_disponibile_h[i] = energia_disp_batt
                        Energia_disp[i] = energia_disp
                        P_pv_h[i] = p_pv
                        max_erogabile[i] = e_batt_max*self.eff_batt + p_pv
                        # fino a qui lo fa per ogni ora

                    e_H2_ = np.sum(E_H2_h)
                    e_im_ = np.sum(E_im_h)
                    m_H2_ = np.sum(M_H2_h)
                    spegn_giorn_proj = self.off + self.count2 #- len(self.E_PV)/24 # qui ci sono i numeri di spegnimenti durante il giorno
                    e_TOT = len(self.E_PV) * p_elc
                    cf = e_H2_ / e_TOT * 100
                    autoconsumo = e_H2_ / self.e_pv * 100








                    # qui creo i vettori che alla posizione k-esima ci sono tutte le informazioni sul progetto k
                    self.spegn_giorn[index] = spegn_giorn_proj
                    self.OFF[index] = self.off
                    self.potenza_elett[index] = p_elc
                    self.potenza_batt[index] = p_batt
                    self.CF[index] = cf
                    self.Auto[index] = autoconsumo
                    self.E_H2[index] = e_H2_
                    self.E_im[index] = e_im_
                    self.M_H2[index] = m_H2_
                    self.andamenti[index, 0, :] = E_H2_h
                    self.andamenti[index, 1, :] = M_H2_h
                    self.andamenti[index, 2, :] = E_im_h
                    self.andamenti[index, 3, :] = P_pv_h
                    self.andamenti[index, 4, :] = Pot_min_elett
                    self.andamenti[index, 5, :] = Pot_max_elett
                    self.andamenti[index, 6, :] = E_batt
                    self.andamenti[index, 7, :] = Pot_max_batt
                    self.andamenti[index, 8, :] = E_batt_disponibile_h
                    self.andamenti[index, 9, :] = Energia_disp
                    self.andamenti[index, 10, :] = max_erogabile
                    if self.lingua == "ENG":
                        self.print_progress_bar(index + 1, len(self.E_H2), suffix = "Technical Analysis with Battery")
                    elif self.lingua == "ITA":
                        self.print_progress_bar(index + 1, len(self.E_H2), suffix = "Analisi Tecnica con batteria")
                    index += 1

    def run_analysis(self):
        if self.batteria == "SI":
            self.run_analysis_battery_static_min()
        elif self.batteria == "NO":
            self.run_analysis_nobattery()



    #da qui in poi ci sono le funzioni per la barra di completamento
    def interpolate_color(self, start_rgb, end_rgb, factor):
        return tuple(int(start + (end - start) * factor) for start, end in zip(start_rgb, end_rgb))
    def rgb_to_ansi(self, rgb):
        return f'\033[38;2;{rgb[0]};{rgb[1]};{rgb[2]}m'
    def print_progress_bar(self, iteration, total, prefix='', suffix="Completamento", decimals=1, length=50, fill='█'):


        start_color = (244, 7, 254)  # #F407FE (Magenta)
        end_color = (7, 213, 223)  # #07D5DF (Cyan)

        percent = iteration / float(total)
        filled_length = int(length * percent)

        bar = ""
        for i in range(length):
            color = self.interpolate_color(start_color, end_color, i / length )
            bar += self.rgb_to_ansi(color) + (fill if i < filled_length else '-')

        sys.stdout.write(f'\r{prefix} |{bar}\033[0m| {percent * 100:.{decimals}f}% {suffix}')
        sys.stdout.flush()

        if iteration == total:
            print()

class Analisi_finanziaria:
    def __init__(self, Terr = 0, OpeE = 0, ImpPV1 = 0, ImpPV1eurokW = 800, EletteuroKW = 1650, CompreuroKW = 4000, AccuE = 0, AccuEeurokW = 200,
                 idrogstocperc = 1/10, StazzRif = 500000, SpeTOpere = 0, BombSto = 0, LavoImp = 0, CarrEll = 0, CapFac = 0, PotEle = 0,
                 tassoDEN = 0.005, ProdAnnuaIdrogkg = 0, bar = 300, costlitroacqua = 0.035, costounitariostoccaggio = 0, PercEserImp = 0.005, Percentimpianti = 0.0025,
                 PercentOpeEd = 0.0005, SpesAmmGen = 7000, Affitto = 0, CostiPersonal = 0, AltriCost = 0, IVAsualtriCost = "SI", DurPianEcon = 20, inflazione = 0.02, inflazionePrezzoElet = 0.02,
                 inflazioneIdrog= 0.01,  tassoVAN = 0.1, incentpubb = 0, duratincentpubb = 0, prezzoindrogeno = 10, ProdElettVend = 0, EnergiaAutocons = 0,
                 prezzoElett = 1, ContrPubb = 0, DebitoSenior = 0.8, DurDebitoSenior = 20, tassoDebito = 0.05, FreqPagamenti = 1, tassoPonte = 0,
                 DurataPonte = 0, aliquoMedia = 0.275, MaxInterssDed = 0.3, lingua = "ITA", Perciva = 0.22, spegn_giorn = 0):
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



        for index,(el1,el2) in enumerate(zip(self.AMMTerr,self.AMMOpeE)):
            self.AMMFabbrTerr[index] = el1 + el2



        for index,(el1, el2, el3, el4, el5, el6, el7, el8, el9) in enumerate(zip(self.AMMPV1, self.AMMElett, self.AMMSTOC, self.AMMSTAZZ, self.AMMSpeTec, self.AMMBombStoc, self.AMMLavImp, self.AMMCARR, self.AMMAccuE)):
            self.AMMMacchImp[index] = el1 + el2 + el3 + el4 + el5 + el6 + el7 + el8 + el9

    def calcola_flussi_capitali(self):
        # gestione dei flussi del debito
        if self.DurDebitoSenior != 0:
            tasso_periodico = ((1 + self.tassoDebito) ** (1 / self.FreqPagamenti)) - 1
            numero_rate = int(self.DurDebitoSenior * self.FreqPagamenti)
            rata = self.DebitoSeniorAss * (tasso_periodico * (1 + tasso_periodico) ** numero_rate) / ((1 + tasso_periodico) ** numero_rate - 1) #calcolo controllato in data 27/11/24 da Stefano Pagani e Giacomo Pamìo
            saldo = self.DebitoSeniorAss
            dtype = [('Periodo', int), ('Rata', float), ('Interesse', float), ('Capitale', float), ('Saldo', float)]
            self.Flussi_debito_non_corretti = np.zeros(numero_rate, dtype=dtype)
            for i in range(numero_rate):
                interesse = saldo * tasso_periodico
                capitale = rata - interesse
                saldo = saldo - capitale
                self.Flussi_debito_non_corretti[i] = (i + 1, rata, interesse, capitale, saldo)
            for anno in range(self.DurDebitoSenior):
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
            self.Flussi_debito[1:self.DurDebitoSenior+1] = self.SumAnnRata[:self.DurDebitoSenior]
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
            rata_fissa = self.iva * (tasso_annuo * (1 + tasso_annuo) ** self.DurataPonte) / ((1 + tasso_annuo) ** self.DurataPonte - 1)
            saldo = self.iva
            for index in range(1, self.DurataPonte + 1):
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
        self.CostiPers = [-abs(value) if value != 0 else 0 for value in self.CostiPers]
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
                'Labor Cost': self.CostiPers,
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
                'Labor Cost': self.CostiPers,
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
                'Costo del Personale': self.CostiPers,
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
                'Costo del Personale': self.CostiPers,
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

    def RUN(self):
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
        self.costruzione_tabelle()

class Analisi_combinata:
    def __init__(self, file_path): # qui c'è solo un'attributo della classe (dovuto al fatto che al suo interno ci sono tutti i valori)
        warnings.filterwarnings("ignore", category=UserWarning, module='openpyxl')  #riga aggiunta da capire dove dovrebbe stare
        self.df = pd.read_excel(file_path, sheet_name=0)  # Leggi il file Excel
        self.file_path = file_path
        self.read_excel() # in questa funzione ci sono le logiche di ottenimento dei valori e di creazioni degli attributi della simmulazione
        # noinspection PyUnresolvedReferences
        self.analisi1 = Analisi_tecnica(self.file_csv, self.tipo_file, self.p_PV, self.dP_el, self.batteria, self.dP_bat,self.min_batt, self.max_batteria, self.lingua, self.eff_batt, self.min_elet)
        self.analisi1.run_analysis() # qui praticamente faccio l'analisi tecnica della mia simulazione
        self.potenza_impianto = self.analisi1.p_PV # era solo per avere chiaro la potenza dell'impianto (avrei potuto lasciare self.analisi1.p_PV)
        self.lista_con_istanze_analisi1 = np.empty(self.analisi1.qt_progetti, dtype=object) # qui salverò tutte le istanze create durante l'analisi finanziaria
        self.lista_con_andamenti_annui = np.empty(self.analisi1.qt_progetti, dtype=object) # qui salverò i flussi energetici dei precedenti progetti (nello stesso ordine)
        self.lista_con_istanze_analisi1_ordinate = np.empty(self.analisi1.qt_progetti, dtype=object) # poi verranno ordinati e salvati inn questa lista
        self.lista_con_andamenti_annui_ordinate = np.empty(self.analisi1.qt_progetti, dtype=object)
        self.valori_simulatore = [] # qui salvo i valori per cui andrò a fare la simulazione per il prezzo di equilibrio
        self.prezzi_eq = [] # qui dentro invece salvo i prezzi quando vado ad ottimizzare per il prezzo di equilibrio; diverso rispetto a prima
        self.VAN_eq = [] # qui dentro invece salvo i van quando vado ad ottimizzare per il prezzo di equilibrio; diverso rispetto a prima
        self.traduci_attributo()
        if self.attributo not in ["Prezzo idrogeno","Hydrogen Price"]: # in caso in cui il criterio di ottimizzazione sia diverso rispetto al prezzo
            if self.batteria == "NO": # e la simulazione è senza batteria

                for index,(el1,el2,el3,el4,el5,el6,el7) in enumerate(zip(self.analisi1.CF, self.analisi1.P_elc, self.analisi1.E_H2, self.analisi1.E_im, self.analisi1.M_H2, self.analisi1.andamenti, self.analisi1.spegn_giorn)):
                    self.analisi2 = Analisi_finanziaria(Terr = self.Terr, OpeE = self.OpeE, ImpPV1 = self.potenza_impianto, ImpPV1eurokW = self.ImpPV1eurokW, EletteuroKW = self.EletteuroKW,
                                                        CompreuroKW = self.CompreuroKW, AccuE = 0, AccuEeurokW = self.AccuEeurokW,idrogstocperc = self.idrogstocperc, StazzRif = self.StazzRif,
                                                        SpeTOpere = self.SpeTOpere, BombSto = self.BombSto, LavoImp = self.LavoImp, CarrEll = self.CarrEll, CapFac = el1, PotEle = el2,
                                                        tassoDEN = self.tassoDEN, ProdAnnuaIdrogkg = el5, bar = self.bar, costlitroacqua = self.costlitroacqua, costounitariostoccaggio = self.costounitariostoccaggio,
                                                        PercEserImp = self.PercEserImp, Percentimpianti = self.Percentimpianti, PercentOpeEd = self.PercentOpeEd, SpesAmmGen = self.SpesAmmGen,
                                                        Affitto = self.Affitto, CostiPersonal = self.CostiPersonal, AltriCost = self.AltriCost, IVAsualtriCost  = self.IVAsualtriCost, DurPianEcon = self.DurPianEcon, inflazione = self.inflazione,
                                                        inflazionePrezzoElet = self.inflazionePrezzoElet, inflazioneIdrog= self.inflazioneIdrog,  tassoVAN = self.tassoVAN,
                                                        incentpubb = self.incentpubb, duratincentpubb = self.duratincentpubb, prezzoindrogeno = self.prezzoindrogeno, ProdElettVend = el4,
                                                        EnergiaAutocons = el3, prezzoElett = self.prezzoElett, ContrPubb = self.ContrPubb, DebitoSenior = self.DebitoSenior,
                                                        DurDebitoSenior = self.DurDebitoSenior, tassoDebito = self.tassoDebito, FreqPagamenti = self.FreqPagamenti, tassoPonte = self.tassoPonte,
                                                        DurataPonte = self.DurataPonte, aliquoMedia = self.aliquoMedia, MaxInterssDed = self.MaxInterssDed, lingua = self.lingua, Perciva = self.Perciva, spegn_giorn = el7)
                    self.analisi2.RUN() # fai l'analisi finanziaria
                    self.lista_con_istanze_analisi1[index] = self.analisi2 # qui trovi le istanze della classe analisi_finanziaria
                    self.lista_con_andamenti_annui[index] = el6 # qui trovi gli andamenti energetici degli stessi progetti che sono sopra (hanno gli stessi indici)
                    if self.lingua == "ENG":
                        self.print_progress_bar(index + 1, len(self.lista_con_istanze_analisi1), suffix = "Financial Analysis without Battery")
                    elif self.lingua == "ITA":
                        self.print_progress_bar(index + 1, len(self.lista_con_istanze_analisi1), suffix = "Analisi finanziaria senza batteria")
            if self.batteria == "SI": # se invce la batteria c'è fai così

                for index,(el1,el2,el3,el4,el5,el6,el7, el8) in enumerate(zip(self.analisi1.potenza_batt, self.analisi1.CF, self.analisi1.potenza_elett, self.analisi1.E_H2, self.analisi1.E_im, self.analisi1.M_H2, self.analisi1.andamenti, self.analisi1.spegn_giorn)):
                    self.analisi2 = Analisi_finanziaria(Terr = self.Terr, OpeE = self.OpeE, ImpPV1 = self.potenza_impianto, ImpPV1eurokW = self.ImpPV1eurokW, EletteuroKW = self.EletteuroKW,
                                                        CompreuroKW = self.CompreuroKW, AccuE = el1, AccuEeurokW = self.AccuEeurokW,idrogstocperc = self.idrogstocperc, StazzRif = self.StazzRif,
                                                        SpeTOpere = self.SpeTOpere, BombSto = self.BombSto, LavoImp = self.LavoImp,CarrEll = self.CarrEll, CapFac = el2, PotEle = el3,
                                                        tassoDEN = self.tassoDEN, ProdAnnuaIdrogkg = el6, bar = self.bar, costlitroacqua = self.costlitroacqua, costounitariostoccaggio = self.costounitariostoccaggio,
                                                        PercEserImp = self.PercEserImp, Percentimpianti = self.Percentimpianti, PercentOpeEd = self.PercentOpeEd, SpesAmmGen = self.SpesAmmGen,
                                                        Affitto = self.Affitto, CostiPersonal = self.CostiPersonal, AltriCost = self.AltriCost, IVAsualtriCost  = self.IVAsualtriCost, DurPianEcon = self.DurPianEcon, inflazione = self.inflazione,
                                                        inflazionePrezzoElet = self.inflazionePrezzoElet, inflazioneIdrog= self.inflazioneIdrog,  tassoVAN = self.tassoVAN,
                                                        incentpubb = self.incentpubb, duratincentpubb = self.duratincentpubb, prezzoindrogeno = self.prezzoindrogeno, ProdElettVend = el5,
                                                        EnergiaAutocons = el4, prezzoElett = self.prezzoElett, ContrPubb = self.ContrPubb, DebitoSenior = self.DebitoSenior,
                                                        DurDebitoSenior = self.DurDebitoSenior, tassoDebito = self.tassoDebito, FreqPagamenti = self.FreqPagamenti, tassoPonte = self.tassoPonte,
                                                        DurataPonte = self.DurataPonte, aliquoMedia = self.aliquoMedia, MaxInterssDed = self.MaxInterssDed, lingua = self.lingua, Perciva = self.Perciva, spegn_giorn = el8)
                    self.analisi2.RUN()
                    self.lista_con_istanze_analisi1[index] = self.analisi2 # qui trovi le istanze della classe analisi_finanziaria
                    self.lista_con_andamenti_annui[index] = el7 # qui trovi gli andamenti energetici degli stessi progetti che sono sopra (hanno gli stessi indici)
                    if self.lingua == "ENG":
                        self.print_progress_bar(index + 1, len(self.lista_con_istanze_analisi1), suffix = "Financial Analysis with Battery")
                    elif self.lingua == "ITA":
                        self.print_progress_bar(index + 1, len(self.lista_con_istanze_analisi1), suffix = "Analisi finanziaria con batteria")
            self.scarica_risultati(self.attributo,self.n_progetti)
        if self.attributo in ["Prezzo idrogeno","Hydrogen Price"]: # se invece il criterio di ottimizzaazione sia il prezzo dell'idrogeno
            if self.batteria == "NO": # e la simulazione si fa senza batteria
                print(3)
                for index,(el1,el2,el3,el4,el5,el6,el7) in enumerate(zip(self.analisi1.CF, self.analisi1.P_elc, self.analisi1.E_H2, self.analisi1.E_im, self.analisi1.M_H2, self.analisi1.andamenti, self.analisi1.spegn_giorn)):
                    VAN = - 1 # ad inizio definisco che il VAN e negativo
                    prezzo = 0.5 # comicio con un prezzo uguale a 50 centesimi (è il minimo di prezzo possibile)
                    while VAN < 0: # fino a quando il VAN del progetto ha un valore negativo
                        self.analisi2 = Analisi_finanziaria(Terr = self.Terr, OpeE = self.OpeE, ImpPV1 = self.potenza_impianto, ImpPV1eurokW = self.ImpPV1eurokW, EletteuroKW = self.EletteuroKW,
                                                            CompreuroKW = self.CompreuroKW, AccuE = 0, AccuEeurokW = self.AccuEeurokW,idrogstocperc = self.idrogstocperc, StazzRif = self.StazzRif,
                                                            SpeTOpere = self.SpeTOpere, BombSto = self.BombSto, LavoImp = self.LavoImp, CarrEll = self.CarrEll, CapFac = el1, PotEle = el2,
                                                            tassoDEN = self.tassoDEN, ProdAnnuaIdrogkg = el5, bar = self.bar, costlitroacqua = self.costlitroacqua, costounitariostoccaggio = self.costounitariostoccaggio,
                                                            PercEserImp = self.PercEserImp, Percentimpianti = self.Percentimpianti, PercentOpeEd = self.PercentOpeEd, SpesAmmGen = self.SpesAmmGen,
                                                            Affitto = self.Affitto, CostiPersonal = self.CostiPersonal, AltriCost = self.AltriCost, IVAsualtriCost  = self.IVAsualtriCost, DurPianEcon = self.DurPianEcon, inflazione = self.inflazione,
                                                            inflazionePrezzoElet = self.inflazionePrezzoElet, inflazioneIdrog= self.inflazioneIdrog,  tassoVAN = self.tassoVAN,
                                                            incentpubb = self.incentpubb, duratincentpubb = self.duratincentpubb, prezzoindrogeno = prezzo , ProdElettVend = el4,
                                                            EnergiaAutocons = el3, prezzoElett = self.prezzoElett, ContrPubb = self.ContrPubb, DebitoSenior = self.DebitoSenior,
                                                            DurDebitoSenior = self.DurDebitoSenior, tassoDebito = self.tassoDebito, FreqPagamenti = self.FreqPagamenti, tassoPonte = self.tassoPonte,
                                                            DurataPonte = self.DurataPonte, aliquoMedia = self.aliquoMedia, MaxInterssDed = self.MaxInterssDed, lingua = self.lingua, Perciva = self.Perciva, spegn_giorn = el7)
                        self.analisi2.RUN() # fai l'analisi
                        prezzo += 0.5 # aumenta il prezzo di 50 centesimi
                        if self.analisi2.VAN >= 0: # se il VAN è positivo
                            VAN = self.analisi2.VAN # questo mi farà uscire dal loop; essendo adesso il VAN positivo
                            self.lista_con_istanze_analisi1[index] = self.analisi2 # qui trovi le istanze della classe analisi_finanziaria
                            self.lista_con_andamenti_annui[index] = el6 # qui trovi gli andamenti energetici degli stessi progetti che sono sopra (hanno gli stessi indici)
                            if self.lingua == "ENG":
                                self.print_progress_bar(index + 1, len(self.lista_con_istanze_analisi1), suffix = "Financial Analysis without Battery")
                            elif self.lingua == "ITA":
                                self.print_progress_bar(index + 1, len(self.lista_con_istanze_analisi1), suffix = "Analisi finanziaria senza batteria")
            if self.batteria == "SI": # se invece la batteria c'è fai lo stesso ma con la batteria
                for index,(el1,el2,el3,el4,el5,el6,el7,el8) in enumerate(zip(self.analisi1.potenza_batt, self.analisi1.CF, self.analisi1.potenza_elett, self.analisi1.E_H2, self.analisi1.E_im, self.analisi1.M_H2, self.analisi1.andamenti, self.analisi1.spegn_giorn)):
                    VAN = - 1
                    prezzo = 0.5
                    while VAN < 0:
                        self.analisi2 = Analisi_finanziaria(Terr = self.Terr, OpeE = self.OpeE, ImpPV1 = self.potenza_impianto, ImpPV1eurokW = self.ImpPV1eurokW, EletteuroKW = self.EletteuroKW,
                                                            CompreuroKW = self.CompreuroKW, AccuE = el1, AccuEeurokW = self.AccuEeurokW,idrogstocperc = self.idrogstocperc, StazzRif = self.StazzRif,
                                                            SpeTOpere = self.SpeTOpere, BombSto = self.BombSto, LavoImp = self.LavoImp,CarrEll = self.CarrEll, CapFac = el2, PotEle = el3,
                                                            tassoDEN = self.tassoDEN, ProdAnnuaIdrogkg = el6, bar = self.bar, costlitroacqua = self.costlitroacqua, costounitariostoccaggio = self.costounitariostoccaggio,
                                                            PercEserImp = self.PercEserImp, Percentimpianti = self.Percentimpianti, PercentOpeEd = self.PercentOpeEd, SpesAmmGen = self.SpesAmmGen,
                                                            Affitto = self.Affitto, CostiPersonal = self.CostiPersonal, AltriCost = self.AltriCost, IVAsualtriCost  = self.IVAsualtriCost, DurPianEcon = self.DurPianEcon, inflazione = self.inflazione,
                                                            inflazionePrezzoElet = self.inflazionePrezzoElet, inflazioneIdrog= self.inflazioneIdrog,  tassoVAN = self.tassoVAN,
                                                            incentpubb = self.incentpubb, duratincentpubb = self.duratincentpubb, prezzoindrogeno = prezzo, ProdElettVend = el5,
                                                            EnergiaAutocons = el4, prezzoElett = self.prezzoElett, ContrPubb = self.ContrPubb, DebitoSenior = self.DebitoSenior,
                                                            DurDebitoSenior = self.DurDebitoSenior, tassoDebito = self.tassoDebito, FreqPagamenti = self.FreqPagamenti, tassoPonte = self.tassoPonte,
                                                            DurataPonte = self.DurataPonte, aliquoMedia = self.aliquoMedia, MaxInterssDed = self.MaxInterssDed, lingua = self.lingua, Perciva = self.Perciva, spegn_giorn = el7)
                        self.analisi2.RUN()
                        prezzo += 0.5
                        if self.analisi2.VAN >= 0:
                            VAN = self.analisi2.VAN
                            self.lista_con_istanze_analisi1[index] = self.analisi2 # qui trovi le istanze della classe analisi_finanziaria
                            self.lista_con_andamenti_annui[index] = el7 # qui trovi gli andamenti energetici degli stessi progetti che sono sopra (hanno gli stessi indici)
                            if self.lingua == "ENG":
                                self.print_progress_bar(index + 1, len(self.lista_con_istanze_analisi1), suffix = "Financial Analysis with Battery")
                            elif self.lingua == "ITA":
                                self.print_progress_bar(index + 1, len(self.lista_con_istanze_analisi1), suffix = "Analisi finanziaria con batteria")
            self.scarica_risultati(self.attributo,self.n_progetti)
        print()
        if self.lingua == "ITA":
            print ("Finito :)")
        elif self.lingua == "ENG":
            print ("Done :)")

    def traduci_attributo(self):
        if self.attributo == "IRR":
            self.attributo = "TIR"
        if self.attributo == "NPV":
            self.attributo = "VAN"
        if self.attributo in ["Electrolyzer", "Elettrolizzatore"]:
            self.attributo = "PotEle"
        if self.attributo in ["Batteries","Batteria"]:
            self.attributo = "AccuE"
        if self.attributo in ["Compressor power","Potenza compressore"]:
            self.attributo = "potenza_compressore"
        if self.attributo in ["Hydrogen production","Produzione idrogeno"]:
            self.attributo = "ProdAnnuaIdrogkg"
        if self.attributo in ["Energy injected into the grid","Energia immessa"]:
            self.attributo = "ProdElettVend"
        if self.attributo in ["Self-consumed energy","Energia autoconsumata"]:
            self.attributo = "EnergiaAutocons"
        if self.attributo == "Capacity factor":
            self.attributo = "CapFac"
        if self.attributo in ["Average operating cost","Costo medio operativo"]:
            self.attributo = "costo_medio_operativo"
        if self.attributo == "Full Cost":
            self.attributo = "costo_full_cost"
        if self.attributo == "Levelized Full Cost":
            self.attributo = "costo_full_cost_levelized"
        if self.attributo == "Hydrogen storage cost":
            self.attributo = "impianto_stocc"
        if self.attributo == "Shutoff count":
            self.attributo = "spegn_giorn"

    def read_excel(self):
        # estrai i valori per la realizzazione dei grafici
        self.lista_valori_investimenti = self.df.iloc[7:20,3].tolist()  # dove 6:13 (da riga 6 a 13, la prima è 0) e 2 è la colonna C (0 è la prima)
        self.lista_valori_dati_econ1 = self.df.iloc[7:16, 6].tolist()
        self.lista_valori_dati_econ2 = self.df.iloc[7:26, 9].tolist()
        self.lista_valori_dati_tec = self.df.iloc[7:17, 12].tolist()
        self.lista_valori_facolt = self.df.iloc[7:9, 15].tolist()
        self.lista_valori = self.lista_valori_investimenti + self.lista_valori_dati_econ1 + self.lista_valori_dati_econ2 + self.lista_valori_dati_tec + self.lista_valori_facolt
        self.lista_nomi1 = ["Terr", "OpeE", "StazzRif", "SpeTOpere", "BombSto", "LavoImp", "CarrEll", "ImpPV1eurokW", "EletteuroKW", "CompreuroKW", "AccuEeurokW", "costounitariostoccaggio", "idrogstocperc"]
        self.lista_nomi2 = [ "costlitroacqua","PercEserImp", "Percentimpianti","PercentOpeEd", "SpesAmmGen", "Affitto", "CostiPersonal", "AltriCost", "IVAsualtriCost"]
        self.lista_nomi3 = ["DurPianEcon", "inflazione", "inflazionePrezzoElet", "tassoVAN", "incentpubb",
                            "duratincentpubb", "prezzoindrogeno", "inflazioneIdrog", "prezzoElett", "ContrPubb",
                            "DebitoSenior", "DurDebitoSenior", "tassoDebito", "FreqPagamenti", "DurataPonte",
                            "tassoPonte", "aliquoMedia", "MaxInterssDed", "Perciva"]
        self.lista_nomi4 = ["file_csv", "tipo_file", "p_PV", "batteria", "min_batt", "max_batteria", "tassoDEN", "bar",
                            "eff_batt", "min_elet"]
        self.lista_nomi5 = ["dP_el", "dP_bat"]
        self.lista_nomi = self.lista_nomi1 + self.lista_nomi2 + self.lista_nomi3 + self.lista_nomi4 + self.lista_nomi5
        # nella precendente lista ci sono i modi in cui verranno chiamati gli attributi della simulazione
        for nome_variabile, valore in zip(self.lista_nomi, self.lista_valori):
            if valore == "YES":  # questo fa in modo che il mio file di INPUT mi traduce sempre in SI; così da rendere l'input in italiano funzionante nello stesso modo rispetto all?input in inglese
                valore = "SI"
            setattr(self, nome_variabile, valore)  # qui creo gli attributi della simulazione
            #print(nome_variabile, valore)

        self.attributo = self.df.iloc[29, 2]  # qui salvo il criterio di ottimizzazione della simulazione
        #print(self.attributo)
        self.n_progetti = self.df.iloc[31, 2]  # qui salvo il numero di progetti che voglio visionare
        #print(self.n_progetti)
        self.lingua = self.df.iloc[33, 2]  # qui salvo la lingua nel quale viene scritto l'output
        #print(self.lingua)
        self.relazione = self.df.iloc[28, 8]  # qui salvo se si vuole o meno fare i grafici della relazione tra le variabili
        #print(self.relazione)
        if self.relazione in ["SI", "YES"]:
            self.variable_1_list = self.extract_column_values(self.df, start_row=30,col=5)  # nelle seguenti liste salvo le variabili di cui voglio disegnare i grafici
            self.variable_2_list = self.extract_column_values(self.df, start_row=30, col=8)

        self.si_fa_simulazione = self.df.iloc[
            28, 12]  # qui salvo se si vuole fare la simulazione del prezzo di equilibrio
        if self.si_fa_simulazione in ["SI", "YES"]:
            self.attributo_simulazione = self.df.iloc[29, 11]  # qui salvo o di fare il prezzo di equilibrio o l'incentivo pubblico

        self.si_fa_grafico_SA = self.df.iloc[28, 15]  # qui salvo se si vuole fare la simulazione del prezzo di equilibrio
        if self.si_fa_grafico_SA in ["SI", "YES"]:
            self.SA_variable_list = self.extract_column_values(self.df, start_row=29,col=14)  # nelle seguenti liste salvo le variabili di cui voglio disegnare i grafici

        self.translation_matrix = pd.read_excel("INPUT.xlsx", sheet_name="Translation", usecols="E:N",skiprows=2)# Skip the first 2 rows to start from row 3

    def extract_column_values(self, df, start_row, col):
        values = []
        for index, value in enumerate(df.iloc[start_row:, col], start=start_row):
            if pd.isna(value):
                break
            values.append(value)
        return values # mi serviva per salvare i valori delle relazioni

    def ordina_lista2(self, attributo1, attributo2):
        # Filtro la lista per rimuovere gli elementi con attributo None e memorizzo anche gli indici originali

        lista_filtrata = [(index, item) for index, item in enumerate(self.lista_con_istanze_analisi1) if getattr(item, attributo1) is not None and getattr(item, attributo2) is not None]
        # Controllo se l'attributo è uno di quelli per cui vogliamo ordinare in modo decrescente
        if attributo1 in ["prezzoindrogeno", "costo_medio_operativo", "costo_full_cost", "costo_full_cost_levelized","spegn_giorn"]:
            lista_filtrata.sort(key=lambda x: getattr(x[1], attributo1), reverse=True)
        else:
            lista_filtrata.sort(
                key=lambda x: getattr(x[1], attributo1))  # qui è in modo crescente (tipo nel caso del VAN)
        # Recupero gli indici ordinati
        indici_ordinati = [item[0] for item in lista_filtrata]
        self.lista_con_andamenti_annui_ordinate = [self.lista_con_andamenti_annui[i] for i in indici_ordinati]
        return [item[1] for item in lista_filtrata]

    def ordina_lista(self, attributo):
        # Filtro la lista per rimuovere gli elementi con attributo None e memorizzo anche gli indici originali
        lista_filtrata = [(index, item) for index, item in enumerate(self.lista_con_istanze_analisi1) if getattr(item, attributo) is not None]
        # Controllo se l'attributo è uno di quelli per cui vogliamo ordinare in modo decrescente
        if attributo in ["prezzoindrogeno", "costo_medio_operativo", "costo_full_cost", "costo_full_cost_levelized","spegn_giorn"]:
            lista_filtrata.sort(key=lambda x: getattr(x[1], attributo), reverse=True)
        else:
            lista_filtrata.sort(key=lambda x: getattr(x[1], attributo)) # qui è in modo crescente (tipo nel caso del VAN)
        # Recupero gli indici ordinati
        indici_ordinati = [item[0] for item in lista_filtrata]
        self.lista_con_andamenti_annui_ordinate = [self.lista_con_andamenti_annui[i] for i in indici_ordinati]
        return [item[1] for item in lista_filtrata]

    def tabella_topN_per(self, attributo1, n=10):
        # qui faccio in modo che sia in grado di funzionare sia in caso di INPUT in italiano sia inglese
        if attributo1 == "IRR":
            attributo1 = "TIR"
        if attributo1 == "NPV":
            attributo1 = "VAN"
        if attributo1 in ["Elettrolizzatore", "Electrolyzer"]:
            attributo1 = "PotEle"
        elif attributo1 in ["Batteria","Batteries"]:
            attributo1 = "AccuE"
        elif attributo1 == ["Produzione idrogeno","Hydrogen production"]:
            attributo1 = "ProdAnnuaIdrogkg"
        elif attributo1 in ["Potenza compressore","Compressor power"]:
            attributo1 = "potenza_compressore"
        elif attributo1 in ["Energia immessa","Energy injected into the grid"]:
            attributo1 = "ProdElettVend"
        elif attributo1 in ["Energia autoconsumata","Self-consumed energy"]:
            attributo1 = "EnergiaAutocons"
        elif attributo1 == "Capacity factor":
            attributo1 = "CapFac"
        elif attributo1 in ["Prezzo idrogeno","Hydrogen Price"]:
            attributo1 = "prezzoindrogeno"
        elif attributo1 in ["Costo medio operativo","Average operating cost"]:
            attributo1 = "costo_medio_operativo"
        elif attributo1 == "Full Cost":
            attributo1 = "costo_full_cost"
        elif attributo1 == "Levelized Full Cost":
            attributo1 = "costo_full_cost_levelized"
        elif attributo1 in ["Hydrogen storage cost",]:
            attributo1 = "impianto_stocc"
        self.lista_con_istanze_analisi1_ordinate = self.ordina_lista(attributo1) # ordino le liste in base al criterio di ottimizzazione
        lista_ordinata_top = self.lista_con_istanze_analisi1_ordinate[-n:] # e salvo solo i progettti che voglio visionare
        valori_VAN = [getattr(item, "VAN") for item in lista_ordinata_top] # salvo i VAN in una lista così da poter fare i sommario dei progetti
        valori_TIR = [getattr(item, "TIR") for item in lista_ordinata_top] # lo faccio per tutti i restanti valori che voglio mettere nel sommario
        valori_PAYBACK = [getattr(item, "PAYBACK") for item in lista_ordinata_top]
        valori_CapFac = [getattr(item, "CapFac") for item in lista_ordinata_top]
        valori_ImpPV1 = [getattr(item, "ImpPV1") for item in lista_ordinata_top]
        valori_PotEle = [getattr(item, "PotEle") for item in lista_ordinata_top]
        valori_EnergiaAutocons = [getattr(item, "EnergiaAutocons") for item in lista_ordinata_top]
        valori_EnergiaAutoconsPerc = valori_EnergiaAutocons/self.analisi1.e_pv
        valori_ProdAnnuaIdrogkg = [getattr(item, "ProdAnnuaIdrogkg") for item in lista_ordinata_top]
        valori_ProdElettVend = [getattr(item, "ProdElettVend") for item in lista_ordinata_top]
        valori_PotenzaCompressore = [getattr(item, "potenza_compressore") for item in lista_ordinata_top]
        valori_prezzi = [getattr(item, "prezzoindrogeno") for item in lista_ordinata_top]
        valori_costoStoccaggio = [getattr(item, "impianto_stocc") for item in lista_ordinata_top]
        if attributo1 not in ["Prezzo idrogeno","Hydrogen Price"]: # se devo fare la simulazione del prezzo di equilibrio salvo i valori dei migliori progetti
            self.valori_simulatore.append(valori_CapFac) # lo fa solo se il criterion di ottimizzazione è diverso rispetto al prezzo
            self.valori_simulatore.append(valori_ImpPV1)
            self.valori_simulatore.append(valori_PotEle)
            self.valori_simulatore.append(valori_EnergiaAutocons)
            self.valori_simulatore.append(valori_ProdAnnuaIdrogkg)
            self.valori_simulatore.append(valori_ProdElettVend)
        if self.analisi1.batteria == "SI":
            valori_batteria = [getattr(item, "AccuE") for item in lista_ordinata_top] # se è con batteria salva anche questa
            if attributo1 not in ["Prezzo idrogeno","Hydrogen Price"]:
                self.valori_simulatore.append(valori_batteria)
        lista_ordinata_top_flussi = self.lista_con_andamenti_annui_ordinate[-n:] # salvo i flussi energetici dei migliori n progetti
        self.dfs_flussi_energetici = [] # qui salverò i dataframe dei flussi energetici dei migliori progetti
        for el in lista_ordinata_top_flussi: # qui ci sono in fila i flussi energetici dal miglior n-esimo progetto fino al migliore progetto
            data = {
                'Energia in elettrolizzatore per ogni ora' : el[0],
                'Energia immessa in rete per ogni ora' : el[2],
                'Energia prodotta dal fotovoltaico' : el[3],
                'Potenza minima elettrolizzatore' : el[4],
                'Potenza massima elettrolizzatore' : el[5],
            }
            if self.batteria == "SI":
                data['Energia in batteria ora per ora'] = el[6]
                data['Potenza massima batteria'] = el[7]
                data['Energia disponibile in batteria'] = el[8]
                data['Energia disponibile'] = el[9]
                data['Max erogabile'] = el[10]
            TabellaFlussi = pd.DataFrame(data)
            TabellaFlussi = TabellaFlussi.transpose()
            self.dfs_flussi_energetici.append(TabellaFlussi) # lughezza di essa dipende dai progetti che si vuole visionare
        if self.lingua == "ITA": # questo è ciò che verrà scritto nel sommario dei progetti in italiano
            data = {
                'VAN Progetto': valori_VAN,
                'TIR Progetto': valori_TIR,
                'PAYBACK Progetto [anni]': valori_PAYBACK,
                'Capacity Factor [%]': valori_CapFac,
                'Potenza impianto PV [kW]': valori_ImpPV1,
                'Potenza Elettrolizzatore [kW]': valori_PotEle,
                'Potenza Compressore [kW]': valori_PotenzaCompressore,
                'Energia Autoconsumata [kWh]': valori_EnergiaAutocons,
                'Autoconsumo [%]': valori_EnergiaAutoconsPerc,
                'Produzione idrogeno [kg] (primo anno)': valori_ProdAnnuaIdrogkg,
                'Elettricità immessa in rete [kWh]': valori_ProdElettVend
            }
            if self.analisi1.batteria == "SI":
                data['Dimensione Batteria [kW]'] = valori_batteria
            if self.attributo in ["Prezzo idrogeno","Hydrogen Price"]:
                data['Prezzo idrogeno [€]'] = valori_prezzi
        if self.lingua == "ENG": # questo è in inglese
            data = {
                'Project NPV': valori_VAN,
                'Project IRR': valori_TIR,
                'Project PAYBACK [years]': valori_PAYBACK,
                'Capacity Factor [%]': valori_CapFac,
                'PV Plant Power [kW]': valori_ImpPV1,
                'Electrolyzer Power [kW]': valori_PotEle,
                'Compressor Power [kW]': valori_PotenzaCompressore,
                'Self-Consumed Energy [kWh]': valori_EnergiaAutocons,
                'Self-Consumption [%]': valori_EnergiaAutoconsPerc,
                'Hydrogen Production [kg] (first year)': valori_ProdAnnuaIdrogkg,
                'Electricity Injected into the Grid [kWh]': valori_ProdElettVend
            }
            if self.analisi1.batteria == "SI":
                data['Battery Size [kWh]'] = valori_batteria
            if self.attributo in ["Prezzo idrogeno", "Hydrogen Price"]:
                data['Hydrogen Price [€]'] = valori_prezzi

        self.TabellaMax = pd.DataFrame(data) # questa è la tabella con ciò che entra nel sommario progetti
        self.TabellaMax = self.TabellaMax.transpose()
        column_names = [f'Project {i+1}' for i in range(n)] # qui creo una lista che diveterà il nome delle colonne
        column_names.reverse()
        self.TabellaMax.columns = column_names

        self.dfs_conto_economico = [getattr(item, "dfContoEconomico") for item in lista_ordinata_top] # qui ci sono i dataframe con i conti economici
        self.dfs_indici_finanziari = [getattr(item, "dfIndiciFinanziari") for item in lista_ordinata_top] # qui ci sono i dataframe con gli indici
        self.dfs_flussi_monetari = [getattr(item, "dfFlussiMonetari") for item in lista_ordinata_top] # qui ci sono i flussi di cassa

    def simulatore_simulatore(self,attributo): # qui faccio le simulazioni per prezzo e incentivo di equilibrio
        if attributo in ["Prezzo idrogeno","Hydrogen Price"]: # se è il prezzo ad essere simulato
            for i in range(self.n_progetti):
                if self.lingua == "ENG":
                    self.print_progress_bar(i + 1, self.n_progetti, suffix = "Equilibrium price evaluation")
                elif self.lingua == "ITA":
                    self.print_progress_bar(i+ 1, self.n_progetti, suffix = "Valutazione di prezzo di equilibrio")
                VAN = -1 # come nel caso simile a prima vado a dire che il van è negativo e mi fai i loop fino a quando il van non è positivo
                prezzo = 0.5
                while VAN < 0:
                    if self.batteria == "NO": # self.valori_simulatore è una lista di liste a cui all'interno ci sono gli input dei migliori n progetti
                        self.analisi2 = Analisi_finanziaria(Terr = self.Terr, OpeE = self.OpeE, ImpPV1 = self.valori_simulatore[1][i], ImpPV1eurokW = self.ImpPV1eurokW, EletteuroKW = self.EletteuroKW,
                                                            CompreuroKW = self.CompreuroKW, AccuE = 0, AccuEeurokW = self.AccuEeurokW,idrogstocperc = self.idrogstocperc, StazzRif = self.StazzRif,
                                                            SpeTOpere = self.SpeTOpere, BombSto = self.BombSto, LavoImp = self.LavoImp,CarrEll = self.CarrEll, CapFac = self.valori_simulatore[0][i], PotEle = self.valori_simulatore[2][i],
                                                            tassoDEN = self.tassoDEN, ProdAnnuaIdrogkg = self.valori_simulatore[4][i], bar = self.bar, costlitroacqua = self.costlitroacqua, costounitariostoccaggio = self.costounitariostoccaggio,
                                                            PercEserImp = self.PercEserImp, Percentimpianti = self.Percentimpianti, PercentOpeEd = self.PercentOpeEd, SpesAmmGen = self.SpesAmmGen,
                                                            Affitto = self.Affitto, CostiPersonal = self.CostiPersonal, AltriCost = self.AltriCost, IVAsualtriCost  = self.IVAsualtriCost, DurPianEcon = self.DurPianEcon, inflazione = self.inflazione,
                                                            inflazionePrezzoElet = self.inflazionePrezzoElet, inflazioneIdrog= self.inflazioneIdrog,  tassoVAN = self.tassoVAN,
                                                            incentpubb = self.incentpubb, duratincentpubb = self.duratincentpubb, prezzoindrogeno = prezzo, ProdElettVend = self.valori_simulatore[5][i],
                                                            EnergiaAutocons = self.valori_simulatore[3][i], prezzoElett = self.prezzoElett, ContrPubb = self.ContrPubb, DebitoSenior = self.DebitoSenior,
                                                            DurDebitoSenior = self.DurDebitoSenior, tassoDebito = self.tassoDebito, FreqPagamenti = self.FreqPagamenti, tassoPonte = self.tassoPonte,
                                                            DurataPonte = self.DurataPonte, aliquoMedia = self.aliquoMedia, MaxInterssDed = self.MaxInterssDed, lingua = self.lingua, Perciva = self.Perciva)
                    elif self.batteria == "SI":
                        self.analisi2 = Analisi_finanziaria(Terr = self.Terr, OpeE = self.OpeE, ImpPV1 = self.valori_simulatore[1][i], ImpPV1eurokW = self.ImpPV1eurokW, EletteuroKW = self.EletteuroKW,
                                                            CompreuroKW = self.CompreuroKW, AccuE = self.valori_simulatore[6][i], AccuEeurokW = self.AccuEeurokW,idrogstocperc = self.idrogstocperc, StazzRif = self.StazzRif,
                                                            SpeTOpere = self.SpeTOpere, BombSto = self.BombSto, LavoImp = self.LavoImp,CarrEll = self.CarrEll, CapFac = self.valori_simulatore[0][i], PotEle = self.valori_simulatore[2][i],
                                                            tassoDEN = self.tassoDEN, ProdAnnuaIdrogkg = self.valori_simulatore[4][i], bar = self.bar, costlitroacqua = self.costlitroacqua, costounitariostoccaggio = self.costounitariostoccaggio,
                                                            PercEserImp = self.PercEserImp, Percentimpianti = self.Percentimpianti, PercentOpeEd = self.PercentOpeEd, SpesAmmGen = self.SpesAmmGen,
                                                            Affitto = self.Affitto, CostiPersonal = self.CostiPersonal, AltriCost = self.AltriCost, IVAsualtriCost  = self.IVAsualtriCost, DurPianEcon = self.DurPianEcon, inflazione = self.inflazione,
                                                            inflazionePrezzoElet = self.inflazionePrezzoElet, inflazioneIdrog= self.inflazioneIdrog,  tassoVAN = self.tassoVAN,
                                                            incentpubb = self.incentpubb, duratincentpubb = self.duratincentpubb, prezzoindrogeno = prezzo, ProdElettVend = self.valori_simulatore[5][i],
                                                            EnergiaAutocons = self.valori_simulatore[3][i], prezzoElett = self.prezzoElett, ContrPubb = self.ContrPubb, DebitoSenior = self.DebitoSenior,
                                                            DurDebitoSenior = self.DurDebitoSenior, tassoDebito = self.tassoDebito, FreqPagamenti = self.FreqPagamenti, tassoPonte = self.tassoPonte,
                                                            DurataPonte = self.DurataPonte, aliquoMedia = self.aliquoMedia, MaxInterssDed = self.MaxInterssDed, lingua = self.lingua, Perciva = self.Perciva)
                    self.analisi2.RUN()
                    if self.analisi2.VAN > 0: # fino a quando il VAN non è positivo
                        VAN = self.analisi2.VAN
                        self.prezzi_eq.append(prezzo)
                        self.VAN_eq.append(VAN)
                        break
                    prezzo += 0.5
        elif attributo in ["Incentivo pubblico per kg di idrogeno venduto","Public incentive per kilogram of hydrogen sold"]:
            for i in range(self.n_progetti):
                if self.lingua == "ENG":
                    self.print_progress_bar(i + 1, self.n_progetti, suffix = "Public incentive equilibrium evaluation")
                elif self.lingua == "ITA":
                    self.print_progress_bar(i+ 1, self.n_progetti, suffix = "Valutazione di incentivo pubblico di equilibrio")
                VAN = -1
                prezzo = 0.5
                while VAN < 0:
                    if self.batteria == "NO":
                        self.analisi2 = Analisi_finanziaria(Terr = self.Terr, OpeE = self.OpeE, ImpPV1 = self.valori_simulatore[1][i], ImpPV1eurokW = self.ImpPV1eurokW, EletteuroKW = self.EletteuroKW,
                                                                        CompreuroKW = self.CompreuroKW, AccuE = 0, AccuEeurokW = self.AccuEeurokW,idrogstocperc = self.idrogstocperc, StazzRif = self.StazzRif,
                                                                        SpeTOpere = self.SpeTOpere, BombSto = self.BombSto, LavoImp = self.LavoImp,CarrEll = self.CarrEll, CapFac = self.valori_simulatore[0][i], PotEle = self.valori_simulatore[2][i],
                                                                        tassoDEN = self.tassoDEN, ProdAnnuaIdrogkg = self.valori_simulatore[4][i], bar = self.bar, costlitroacqua = self.costlitroacqua, costounitariostoccaggio = self.costounitariostoccaggio,
                                                                        PercEserImp = self.PercEserImp, Percentimpianti = self.Percentimpianti, PercentOpeEd = self.PercentOpeEd, SpesAmmGen = self.SpesAmmGen,
                                                                        Affitto = self.Affitto, CostiPersonal = self.CostiPersonal, AltriCost = self.AltriCost, IVAsualtriCost  = self.IVAsualtriCost, DurPianEcon = self.DurPianEcon, inflazione = self.inflazione,
                                                                        inflazionePrezzoElet = self.inflazionePrezzoElet, inflazioneIdrog= self.inflazioneIdrog,  tassoVAN = self.tassoVAN,
                                                                        incentpubb = prezzo, duratincentpubb = self.duratincentpubb, prezzoindrogeno = prezzo, ProdElettVend = self.valori_simulatore[5][i],
                                                                        EnergiaAutocons = self.valori_simulatore[3][i], prezzoElett = self.prezzoElett, ContrPubb = self.ContrPubb, DebitoSenior = self.DebitoSenior,
                                                                        DurDebitoSenior = self.DurDebitoSenior, tassoDebito = self.tassoDebito, FreqPagamenti = self.FreqPagamenti, tassoPonte = self.tassoPonte,
                                                                        DurataPonte = self.DurataPonte, aliquoMedia = self.aliquoMedia, MaxInterssDed = self.MaxInterssDed, lingua = self.lingua, Perciva = self.Perciva)
                    elif self.batteria == "SI":
                        self.analisi2 = Analisi_finanziaria(Terr = self.Terr, OpeE = self.OpeE, ImpPV1 = self.valori_simulatore[1][i], ImpPV1eurokW = self.ImpPV1eurokW, EletteuroKW = self.EletteuroKW,
                                                                        CompreuroKW = self.CompreuroKW, AccuE = self.valori_simulatore[6][i], AccuEeurokW = self.AccuEeurokW,idrogstocperc = self.idrogstocperc, StazzRif = self.StazzRif,
                                                                        SpeTOpere = self.SpeTOpere, BombSto = self.BombSto, LavoImp = self.LavoImp,CarrEll = self.CarrEll, CapFac = self.valori_simulatore[0][i], PotEle = self.valori_simulatore[2][i],
                                                                        tassoDEN = self.tassoDEN, ProdAnnuaIdrogkg = self.valori_simulatore[4][i], bar = self.bar, costlitroacqua = self.costlitroacqua, costounitariostoccaggio = self.costounitariostoccaggio,
                                                                        PercEserImp = self.PercEserImp, Percentimpianti = self.Percentimpianti, PercentOpeEd = self.PercentOpeEd, SpesAmmGen = self.SpesAmmGen,
                                                                        Affitto = self.Affitto, CostiPersonal = self.CostiPersonal, AltriCost = self.AltriCost, IVAsualtriCost  = self.IVAsualtriCost, DurPianEcon = self.DurPianEcon, inflazione = self.inflazione,
                                                                        inflazionePrezzoElet = self.inflazionePrezzoElet, inflazioneIdrog= self.inflazioneIdrog,  tassoVAN = self.tassoVAN,
                                                                        incentpubb = prezzo, duratincentpubb = self.duratincentpubb, prezzoindrogeno = self.prezzoindrogeno, ProdElettVend = self.valori_simulatore[5][i],
                                                                        EnergiaAutocons = self.valori_simulatore[3][i], prezzoElett = self.prezzoElett, ContrPubb = self.ContrPubb, DebitoSenior = self.DebitoSenior,
                                                                        DurDebitoSenior = self.DurDebitoSenior, tassoDebito = self.tassoDebito, FreqPagamenti = self.FreqPagamenti, tassoPonte = self.tassoPonte,
                                                                        DurataPonte = self.DurataPonte, aliquoMedia = self.aliquoMedia, MaxInterssDed = self.MaxInterssDed, lingua = self.lingua, Perciva = self.Perciva)
                    self.analisi2.RUN()
                    if self.analisi2.VAN > 0:
                        VAN = self.analisi2.VAN
                        self.prezzi_eq.append(prezzo)
                        self.VAN_eq.append(VAN)
                        break
                    prezzo += 0.5

    def crea_grafico_relazione(self, attributo1, attributo2):
        lista_ordinata = self.ordina_lista(attributo1)
        valori_x = [getattr(item, attributo1) for item in lista_ordinata] # semplicemente ordino le liste
        valori_y = [getattr(item, attributo2) for item in lista_ordinata]
        return valori_x,valori_y

    def crea_grafico_SA(self, attributo1, attributo2, attributo3):

        lista_ordinata = self.ordina_lista2(attributo1, attributo3)
        valori_x = [getattr(item, attributo1) for item in lista_ordinata] # semplicemente ordino le liste
        valori_y = [getattr(item, attributo2) for item in lista_ordinata]
        valori_z = [getattr(item, attributo3) for item in lista_ordinata]
        return valori_x,valori_y, valori_z


    #mi servono per la scrittura dei grafici di sensitivity analysis, così i colori delle serie passano da Gradients 1 a Gradients 2
    def interpolate_color(self, start_color, end_color, factor):
        #interpolate between two colors by a given factor.
        return [int(start_color[i] + (end_color[i] - start_color[i]) * factor) for i in range(3)]
    def rgb_to_hex(self, rgb):
        """Convert an RGB color to hex string for XlsxWriter."""
        return "#{:02x}{:02x}{:02x}".format(*rgb)
    def hex_to_rgb(self, hex_color):
        """Convert HEX color to an RGB tuple."""
        hex_color = hex_color.lstrip("#")
        return tuple(int(hex_color[i:i + 2], 16) for i in (0, 2, 4))


    def traduci_lista_per_il_codice(self, lista): # qui solo cambio il nome così da poter gestire in modo identico l'input in italiano e inglese
        nuova_lista = []
        for el in lista:
            if el == "NPV":
                el = "VAN"
            if el == "IRR":
                el = "TIR"
            if el in ["Elettrolizzatore", "Electrolyzer"]:
                el = "PotEle"
            elif el in ["Batteria","Batteries"]:
                el = "AccuE"
            elif el in ["Produzione idrogeno","Hydrogen production"]:
                el = "ProdAnnuaIdrogkg"
            elif el in ["Energia immessa","Energy injected into the grid"]:
                el = "ProdElettVend"
            elif el in ["Energia autoconsumata","Self-consumed energy"]:
                el = "EnergiaAutocons"
            elif el in ["Prezzo idrogeno","Hydrogen Price"]:
                el = "prezzoindrogeno"
            elif el == "Capacity factor":
                el = "CapFac"
            elif el in ["Potenza compressore","Compressor power"]:
                el = "potenza_compressore"
            elif el in ["Costo medio operativo","Average operating cost"]:
                el = "costo_medio_operativo"
            elif el == "Full Cost":
                el = "costo_full_cost"
            elif el == "Levelized Full Cost":
                el = "costo_full_cost_levelized"
            elif el == "Hydrogen storage cost":
                el = "impianto_stocc"
            elif el == "Shutoff count":
                el = "spegn_giorn"
            nuova_lista.append(el)
        return nuova_lista


    def traduci_nome(self, el):
        if self.lingua == "ENG":
            return el
        elif self.lingua == "ITA":
            el = next((row[1] for row in self.translation_matrix if row[0] == el), "Not found")  # Search for the key
            return el
        elif self.lingua == "SLO":
            el = next((row[2] for row in self.translation_matrix if row[0] == el), "Not found")  # Search for the key
            return el
        elif self.lingua == "DEU":
            el = next((row[3] for row in self.translation_matrix if row[0] == el), "Not found")  # Search for the key
            return el
        elif self.lingua == "FRA":
            el = next((row[4] for row in self.translation_matrix if row[0] == el), "Not found")  # Search for the key
            return el

    def aggiungi_unita_di_misura(self, el):#qui solo cambio il nome così da poter gestire in modo identico l'input in italiano e inglese
        if el == "NPV":
            el = el +" [€]"
        if el == "IRR":
            el = el +" [%]"
        elif el == "Hydrogen production":
            el = el +" [kg]"
        elif el in "Energy injected into the grid":
            el = self.traduci_nome(el) + " [kWh]"
        elif el == "Self-consumed energy":
            el = self.traduci_nome(el) + " [kWh]"
        elif el == "Hydrogen Price":
            el = self.traduci_nome(el) + " [€]"
        elif el == "Capacity factor":
            el = self.traduci_nome(el) + " [%]"
        elif el == "Compressor power":
            el = self.traduci_nome(el) + " kW"
        elif el == "Average operating cost":
            el = self.traduci_nome(el) + " [ €]"
        elif el == "Full Cost":
            el = self.traduci_nome(el) + " [€]"
        elif el == "Levelized Full Cost":
            el = self.traduci_nome(el) + " [€]"
        elif el == "Hydrogen storage cost":
            el = self.traduci_nome(el) + " [€]"
        elif el == "Shutoff count":
            el = self.traduci_nome(el)

        return el

    def scarica_risultati(self, attributo, n = 15): # è la funzione nel cui interno c'è tutto ciò che riguard il file output
        self.tabella_topN_per(attributo, n) # qui ottengo ciò che voglio poi inserire nel file excel
        if attributo not in ["Prezzo idrogeno","Hydrogen Price"] and self.si_fa_simulazione in ["SI","YES"]:
            self.simulatore_simulatore(self.attributo_simulazione) # faccio la simulazione del prezzo/incentivo di eq
        nome_file = "OUTPUT.xlsx"
        workbook = xlsxwriter.Workbook(nome_file, {'nan_inf_to_errors': True}) # creo il file excel e lo chiamo in quel modo in particolare
        if self.lingua == "ENG":
            titolo = "General Summary"
        if self.lingua == "ITA":
            titolo = 'Sommario generale'
        worksheet1 = workbook.add_worksheet(titolo) # al file excel aggiungo un folgio che chiamo sommario generale

        # definisco i colori che poi andrò ad utilizzare nella creazione del file excel. tutte le descrizioni tra parentesi fanno riferimento al file
        # "Y:\Cooperazione\03747_UC_AMETHyST\03747_UC_WP2_Piloting & applications\A.2.3 Financial assessment tool\TOOL_Versioni_PYTHON\GRAFICA\"
        color_palette = {
            "title_bg": "#5F85EA",  # warm blue (elements 2)
            "header_bg": "#dee4f2",  # light blue-gray color (background 1)
            "cell_bg": "#f8f8f8",  # warm gray color (background 2)
            "gridlines": "#666666",  # lighter gray (text 3)
            "chart_fill1": "#d03938",  # light red (illustrations 2)
            "chart_fill2": "#e0b544",  # ocra (illustrations 3)
            "font_color": "#282a2b",  # font black (text 1)
            "white_color": "#FFFFFF", # standard white color (text 4)
            "negative_red": "#aa312e", #colore rosso per i numeri negativi (illustrations 1)
            "header_bg2": "#95b3d7",
            #ALTRI COLORI PRESENTI NELLA PALETTE
            "Elements_1": "#8f60fe",
            "Gradients_1": "#07d5df",
            "Gradients_2": "#f407fe",
            "Text_2": "#4c4c54",
            "illustrations_4": "#e7cc65",

        }
        title_format = workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'bg_color': color_palette['title_bg'],
            'font_color': color_palette["white_color"],
            'font_name': 'Rajdhani Medium',
            'font_size': 22
        }) # questo è invece il formato di come viene definito il titoli
        subtitle_format = workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'bg_color': color_palette['title_bg'],
            'font_color': color_palette["white_color"],
            'font_name': 'Rajdhani Medium',
            'font_size': 16
        })  # questo è invece il formato di come viene definito il titoli
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'align': 'center',
            'valign': 'vcenter',
            'bg_color': color_palette['header_bg'],
            'font_color': color_palette['font_color'],
            'border': 1,
            'border_color': color_palette['gridlines'],
            'font_name': 'Rajdhani Medium',
            'font_size': 11
        }) # questo è invece il formato di come viene definito i nomi nelle prima colonna e riga dei data frame
        # Cell format
        generic_cell_format = workbook.add_format({
            'bg_color': color_palette['cell_bg'],
            'font_color': color_palette['font_color'],
            'border': 1,
            'border_color': color_palette['gridlines']
        }) # qui definisco il formato generale delle celle
        number_format = workbook.add_format({
            'num_format': '#,##0.00;[Red]-#,##0.00;0.00',  # Red color for negative numbers
            'bg_color': color_palette['cell_bg'],
            'font_color': color_palette['font_color'],
            'border': 1,
            'border_color': color_palette['gridlines'],
            'font_name': 'Roboto',
            'font_size': 11

        }) # questo mi serve per dire che ciò che inserisco è un numero; in caso sia negativo lo metti in rosso

        worksheet1.set_column('A:A', 26.5, generic_cell_format) # dico di mettermi le celle nella colonna A nel formato celle con una lunghezze delle celle pari a 26.5
        worksheet1.set_column('B:XFD', 12, generic_cell_format) # qui dico di farle tutte le altre uguali ma con lunghezza pari a 12
        numero_colonne = len(self.TabellaMax.columns)
        worksheet1.merge_range(0, 0, 0, numero_colonne, titolo, title_format) # metto insieme tutte le colonne sopra il dataframe  e ci scrivo il titolo nel suo formato

        # Inseriamo la tabella del sommario dei progetti
        for idx, col in enumerate(self.TabellaMax.columns):
            worksheet1.write(2, idx + 1, col, header_format) # scrivo le prima riga in un formato particolare
        for row_idx, index in enumerate(self.TabellaMax.index):
            worksheet1.write(row_idx + 3, 0, index, header_format) # scrivo le prima colonna in un formato particolare
            for col_idx, value in enumerate(self.TabellaMax.iloc[row_idx]): # dopodiché inserisco i valori all'interno del dataframe come numeri
                if pd.api.types.is_numeric_dtype(type(value)):
                    worksheet1.write_number(row_idx + 3, col_idx + 1, value, number_format)
                else:
                    worksheet1.write(row_idx + 3, col_idx + 1, value, generic_cell_format)

        if self.relazione in ["SI","YES"] and self.attributo not in ["Prezzo idrogeno","Hydrogen Price"]: # se è stato richiesto di disegnare i grafici
            self.variable_1_list_tradotta = self.traduci_lista_per_il_codice(self.variable_1_list) # modifico i nomi in modo che la simulazione sia in grado di leggere quali variabili disegnare
            self.variable_2_list_tradotta = self.traduci_lista_per_il_codice(self.variable_2_list) # modifico i nomi in modo che la simulazione sia in grado di leggere quali variabili disegnare
            self.SA_variable_list_tradotta = self.traduci_lista_per_il_codice(self.SA_variable_list) # modifico i nomi in modo che la simulazione sia in grado di leggere quali variabili disegnare

            start_row = 16 # dalla riga 16 della prima pagine nel file excel


            #ciclo che scrive i grafici x/y
            for el1, el2, el1_, el2_ in zip(self.variable_1_list_tradotta, self.variable_2_list_tradotta, self.variable_1_list, self.variable_2_list):
                valori_x, valori_y = self.crea_grafico_relazione(el1, el2) # sono le due liste che sono una in relazione all'altra

                for col_idx, value in enumerate(valori_x): # inserisco nel file excel (la prima volta nella riga 16) i valori della varibile
                    worksheet1.write(start_row, col_idx, value, generic_cell_format)
                for col_idx, value in enumerate(valori_y): # subito sotto ci inserisco i valori dell'altra variabile
                    worksheet1.write(start_row + 1, col_idx, value, generic_cell_format)

                # qui creo il grafico
                chart = workbook.add_chart({'type': 'scatter', 'subtype': 'straight_with_markers'})
                # gli dico di prendere i valori appena isritti nel file excel e disegnarli nel grafico
                chart.add_series({
                    'name': f'{el1_} vs {el2_}',
                    'categories': [worksheet1.name, start_row, 0, start_row, len(valori_x) - 1],
                    'values': [worksheet1.name, start_row + 1, 0, start_row + 1, len(valori_y) - 1],
                    'marker': {
                        'type': 'circle',
                        'size': 2,
                        'fill': {'color': color_palette['chart_fill1']},  # Colore di riempimento del marker
                        'border': {'color': color_palette['chart_fill1']}  # Colore del bordo del marker
                    },
                    'line': {'none': True}
                })
                chart.set_title({
                    'name': f'Scatter Plot: {el1_} vs {el2_}',
                    'name_font': {
                        'bold': True,
                        'size': 18,  # La stessa dimensione del font
                        'color': color_palette["font_color"],  # Usa il colore del font che hai definito
                        'name': 'Rajdhani Medium'  # Nome del font
                    }
                })
                # Impostazioni per l'asse X
                chart.set_x_axis({
                    'name': el1_,
                    'min': min(valori_x) - abs(min(valori_x) * 0.1),
                    'max': max(valori_x) + abs(max(valori_x) * 0.1),
                    'major_gridlines': {'visible': True, 'line': {'color': color_palette['gridlines']}},
                    # Light gridlines
                    'name_font': {
                        'bold': True,
                        'size': 13,  # Dimensione del font per il titolo dell'asse X
                        'color': color_palette["font_color"],  # Colore del font per l'asse X
                        'name': 'Rajdhani Medium'  # Nome del font per il titolo dell'asse X
                    },
                    'num_font': {  # Imposta il font per i valori dell'asse X
                        'name': 'Roboto',  # Nome del font per i valori dell'asse X
                        'size': 9,  # Dimensione del font per i valori dell'asse X
                        'color': color_palette["font_color"],  # Colore del font per i valori dell'asse X
                    }
                })

                # Impostazioni per l'asse Y
                chart.set_y_axis({
                    'name': el2_,
                    'min': min(valori_y) - abs(min(valori_y) * 0.1),
                    'max': max(valori_y) + abs(max(valori_y) * 0.1),
                    'major_gridlines': {'visible': True, 'line': {'color': color_palette['gridlines']}},
                    'name_font': {
                        'bold': True,
                        'size': 13,  # Dimensione del font per il titolo dell'asse Y
                        'color': color_palette["font_color"],  # Colore del font per l'asse Y
                        'name': 'Rajdhani Medium'  # Nome del font per il titolo dell'asse Y
                    },
                    'num_font': {  # Imposta il font per i valori dell'asse Y
                        'name': 'Roboto',  # Nome del font per i valori dell'asse Y
                        'size': 9,  # Dimensione del font per i valori dell'asse Y
                        'color': color_palette["font_color"],  # Colore del font per i valori dell'asse Y
                    }
                })
                chart.set_legend({'position': 'bottom'}) # inseriscimi la leggenda in basso
                chart_position = f'A{start_row + 1}'  # Inserisci il grafico tre posizioni più in basso
                worksheet1.insert_chart(chart_position, chart, {'x_offset': 5, 'y_offset': 10, 'x_scale': 1.5, 'y_scale': 1.5})
                start_row += 23 # qui li dico che poi il prossimo grafico lo metti una paio di celle più in basso
                # così scrivo tutto ciò che serve nella prima pagina del file excel


            # ciclo che scrive i grafici z/(x,y)
            k=16 #posizione y primo grafico z/(x,y)
            j=8 #posizione x primo grafico z/(x,y)
            for el1, el1_ in zip(self.SA_variable_list_tradotta, self.SA_variable_list):
                #qui genero le tre liste ordinate con i valori di interesse
                valori_x, valori_y, valori_z = self.crea_grafico_SA("PotEle", "AccuE", el1)  # genera 3 liste di variabili in relazione tra loro; in particolare, ordina per dimensione dell'elettrolizzatore e per ogni simulazione associa la relativa dimensione di accumulo e la terza variabile a scelta dello user

                #vecchio snuippet non pivottato
                '''for col_idx, value in enumerate(valori_x):  # inserisco nel file excel (la prima volta nella riga 16) i valori della varibile
                    worksheet1.write(start_row, col_idx, value, generic_cell_format)
                for col_idx, value in enumerate(valori_y):  # subito sotto ci inserisco i valori dell'altra variabile
                    worksheet1.write(start_row + 1, col_idx, value, generic_cell_format)
                for col_idx, value in enumerate(valori_z):  # subito sotto ci inserisco i valori dell'altra variabile
                    worksheet1.write(start_row + 2, col_idx, value, generic_cell_format)'''
                #vecchio snippet che non filtrava i valori per renderli leggibili ma invece li teneva tutti
                '''# Define the unique values for column and row headers
                unique_x = sorted(set(valori_x))  # PotEle values as column headers
                unique_y = sorted(set(valori_y))  # AccuE values as row headers

                # Write the column headers (PotEle values)
                for col_idx, pot_ele_value in enumerate(unique_x, start=1):
                    worksheet1.write(start_row, col_idx, pot_ele_value, generic_cell_format)

                # Write the row headers (AccuE values) and populate el1 values in table cells
                for row_idx, accu_e_value in enumerate(unique_y, start=start_row + 1):
                    worksheet1.write(row_idx, 0, accu_e_value, generic_cell_format)  # AccuE row header

                    # For each PotEle, find the corresponding el1 value
                    for col_idx, pot_ele_value in enumerate(unique_x, start=1):
                        # Find the matching el1 value for the (PotEle, AccuE) pair
                        try:
                            el1_value = next(
                                el1 for (x, y, el1) in zip(valori_x, valori_y, valori_z)
                                if x == pot_ele_value and y == accu_e_value
                            )
                        except StopIteration:
                            el1_value = None  # Use None or an empty cell if no match is found

                        # Write el1 value into the cell
                        worksheet1.write(row_idx, col_idx, el1_value, generic_cell_format)'''

                # Define the unique values for column and row headers
                unique_x = sorted(set(valori_x))  # PotEle values as column headers
                unique_y = sorted(set(valori_y))  # AccuE values as row headers
                #print(valori_z)
                # Find the PotEle with the maximum el1 value
                max_el1_pot_ele = valori_x[valori_z.index(max(valori_z))]

                # Select n other PotEle values at regular intervals
                # Exclude max_el1_pot_ele from regular selection to avoid duplication                                        ↓↓↓
                filtered_pot_ele = [max_el1_pot_ele] + [unique_x[i] for i in range(0, len(unique_x), max(1, len(unique_x) // 16)) if unique_x[i] != max_el1_pot_ele]

                # Write the filtered PotEle values as column headers in Excel
                for col_idx, pot_ele_value in enumerate(filtered_pot_ele, start=1):
                    rounded_value = round(pot_ele_value, 2)  # Round to the tenths place
                    worksheet1.write(start_row, col_idx, rounded_value, generic_cell_format)

                # Write the row headers (AccuE values) and populate el1 values in table cells
                for row_idx, accu_e_value in enumerate(unique_y, start=start_row + 1):
                    rounded_value = round(accu_e_value, 2)  # Round to the tenths place
                    worksheet1.write(row_idx, 0, rounded_value, generic_cell_format)  # AccuE row header

                    # For each filtered PotEle, find the corresponding el1 value
                    for col_idx, pot_ele_value in enumerate(filtered_pot_ele, start=1):
                        try:
                            # Find the matching el1 value for the (PotEle, AccuE) pair
                            el1_value = next(
                                el1 for (x, y, el1) in zip(valori_x, valori_y, valori_z)
                                if x == pot_ele_value and y == accu_e_value
                            )
                        except StopIteration:
                            el1_value = None  # Use None or an empty cell if no match is found

                        # Write el1 value into the cell
                        worksheet1.write(row_idx, col_idx, el1_value, generic_cell_format)

                # qui creo il grafico
                chart = workbook.add_chart({'type': 'scatter', 'subtype': 'straight_with_markers'})
                # gli dico di prendere i valori appena scritti nel file excel e disegnarli nel grafico
                transition_param = 1
                for i, nomecateg in enumerate(filtered_pot_ele):
                    nomecateg =round(nomecateg,1)
                    factor = nomecateg/max(filtered_pot_ele) * transition_param
                    color = self.rgb_to_hex(self.interpolate_color(self.hex_to_rgb(color_palette["Gradients_1"]), self.hex_to_rgb(color_palette["Gradients_2"]), factor))
                    chart.add_series({
                        'name': f'{nomecateg}',
                        'categories': [worksheet1.name, start_row+1, 0, start_row + len(unique_y), 0],  # X-axis range
                        'values': [worksheet1.name, start_row+1, i+1, start_row + len(unique_y), i+1],# Y-axis range
                        'marker': {
                            'type': 'circle',
                            'size': 4,
                            'fill': {'color': color_palette['font_color']},  # Marker fill color
                            'border': {'color': color_palette['font_color']},  # Marker border color
                        },
                        'line': {
                            'width': 2,
                            'color': color,
                        }
                    })

                chart.set_title({
                    'name': f'Sensitivity analysis chart: {el1_} vs Battery size by electrolizer size',
                    'name_font': {
                        'bold': True,
                        'size': 18,  # La stessa dimensione del font
                        'color': color_palette["font_color"],  # Usa il colore del font che hai definito
                        'name': 'Rajdhani Medium'  # Nome del font
                    }
                })

                # Impostazioni per l'asse X
                if self.lingua == "ITA":
                    x_axis_name = "Capacità accumulo [kWh]"
                elif self.lingua == "ENG":
                    x_axis_name = "Battery size [kWh]"

                chart.set_x_axis({
                    'name': x_axis_name,
                    #'min': min(unique_y) - abs(min(unique_y) * 0.1),'max': max(unique_y) + abs(max(unique_y) * 0.1),
                    'major_gridlines': {'visible': True, 'line': {'color': color_palette['gridlines']}},
                    # Light gridlines
                    'name_font': {
                        'bold': True,
                        'size': 13,  # Dimensione del font per il titolo dell'asse X
                        'color': color_palette["font_color"],  # Colore del font per l'asse X
                        'name': 'Rajdhani Medium'  # Nome del font per il titolo dell'asse X
                    },
                    'num_font': {  # Imposta il font per i valori dell'asse X
                        'name': 'Roboto',  # Nome del font per i valori dell'asse X
                        'size': 9,  # Dimensione del font per i valori dell'asse X
                        'color': color_palette["font_color"],  # Colore del font per i valori dell'asse X
                    }
                })
                # Impostazioni per l'asse Y
                chart.set_y_axis({
                    'name': self.aggiungi_unita_di_misura(el1_),
                    ''''min': min(valori_z) - abs(min(valori_z) * 0.1),
                    'max': max(valori_z) + abs(max(valori_z) * 0.1),'''
                    'major_gridlines': {'visible': True, 'line': {'color': color_palette['gridlines']}},
                    'name_font': {
                        'bold': True,
                        'size': 13,  # Dimensione del font per il titolo dell'asse Y
                        'color': color_palette["font_color"],  # Colore del font per l'asse Y
                        'name': 'Rajdhani Medium'  # Nome del font per il titolo dell'asse Y
                    },
                    'num_font': {  # Imposta il font per i valori dell'asse Y
                        'name': 'Roboto',  # Nome del font per i valori dell'asse Y
                        'size': 9,  # Dimensione del font per i valori dell'asse Y
                        'color': color_palette["font_color"],  # Colore del font per i valori dell'asse Y
                    }
                })
                chart.show_blanks_as('span')
                chart.set_legend({
                    'position': 'right',
                    'title':({
                        'name': "Electrolyzer size",
                        'name_font': ({
                            'bold': True,
                            'size': 13,  # Dimensione del font per il titolo dell'asse X
                            'color': color_palette["font_color"],  # Colore del font per l'asse X
                            'name': 'Rajdhani Medium'  # Nome del font per il titolo dell'asse X
                        })
                    })
                })
                worksheet1.insert_chart(k,j, chart,{'x_offset': 5, 'y_offset': 10, 'x_scale': 3, 'y_scale': 2.5})

                start_row += len(unique_y) + 1
                k += 37
                if self.lingua == "ENG":
                    self.print_progress_bar(self.SA_variable_list_tradotta.index(el1) + 1, len(self.SA_variable_list),
                                            "", "Sensitivity analysis charts compilation", 1, 50)
                elif self.lingua == "ITA":
                    self.print_progress_bar(self.SA_variable_list_tradotta.index(el1) + 1, len(self.SA_variable_list),
                                            "", "Compilazione grafici di sensitivity analysis", 1, 50)




        num_projects = len(self.dfs_conto_economico) # il numero di progetti è pari alla lughezza di questo dataframe essendoci all'interno i dataframe dei migliori n progetti

        #questo for gestisce la scrittura dei fogli dei migliori progetti nell'output
        for i in range(num_projects): # per ogni miglior progetto

            project_num = num_projects - i  # ordino al contrario così da avere il progetto migliore come progetto 1
            worksheet = workbook.add_worksheet(f"PA {project_num}") # crea un nuovo floglio e lo chiami Analisi progetto "numero progetto"
            worksheet.freeze_panes(0, 2)
            worksheet.merge_range('A1:B1', f'Project Analysis {project_num}', title_format) # metti insieme le prime celle e scrivi che si tratta del progetto n
            worksheet.set_column('A:A', 8, generic_cell_format) # scrivo come prima che tutte le celle hanno lo stesso formato ma con diverse lunghezze
            worksheet.set_column('B:B', 54, generic_cell_format)
            worksheet.set_column('C:XFD', 13, generic_cell_format)
            worksheet.set_row(1,63)

            # cominciamo inserendo nel file excel il dataframe con gli indici del progetto
            ind_fin = self.dfs_indici_finanziari[i] # prendi il dataframe nella posizione corrispettivo al progetto
            for col_num, value in enumerate(ind_fin.columns): # la primma riga me la metti il formato particolare
                worksheet.write(1, col_num + 1, value, header_format)
            for row_num, row in enumerate(ind_fin.values):
                for col_num, value in enumerate(row):
                    if isinstance(value, np.ndarray):
                        value = value.item() if value.size == 1 else value.tolist()  # Convert single element or full list

                    worksheet.write(row_num + 2, col_num + 1, value, number_format) # inserisci i valori corrispettivi sotto di una posizione

            if self.si_fa_simulazione in ["SI","YES"] and self.attributo not in ["Prezzo idrogeno","Hydrogen Price"]:
                if self.attributo_simulazione in ["Prezzo idrogeno","Hydrogen Price"]:
                    if self.lingua == "ENG":
                        titolo = "Eq price is:"
                    if self.lingua == "ITA":
                        titolo = 'il prezzo di eq:'
                    worksheet.write('V2', titolo, header_format) # inserisci il prezzo di equilibrio nella cella U2
                elif self.attributo_simulazione in ["Incentivo pubblico per kg di idrogeno venduto","Public incentive per kilogram of hydrogen sold"]:
                    if self.lingua == "ENG":
                        titolo = "Eq incentive is:"
                    if self.lingua == "ITA":
                        titolo = "L'incentivo di eq:"
                    worksheet.write('U2', titolo, header_format) # inserisci l'incentivo di equilibrio nella cella U2
                worksheet.write('V3', self.prezzi_eq[i], number_format) # inserisci il suo valore specifico nella cella T3

            # Inseriamo anche il conto economico partendo dalla riga 4
            start_row = 4
            if self.lingua == "ENG":
                titolo = "Income Statement"
            if self.lingua == "ITA":
                titolo = "Conto Economico"
            worksheet.merge_range(start_row, 0, start_row, 1, titolo, subtitle_format) # come prima metti insieme le celle e le metti nel formato richiesto
            cont_eco = self.dfs_conto_economico[i] # prendi il conto economico che ti serve
            worksheet.write(start_row, 1, '', header_format)  # questa cella la voglio vuota
            for col_num, value in enumerate(cont_eco.columns):
                worksheet.write(start_row, col_num + 2, value, header_format) # come sempre la prima lista la voglio in un'altro formato
            for row_idx, (index, row) in enumerate(cont_eco.iterrows()):
                worksheet.write(row_idx + start_row + 1, 1, index, header_format) # faccio lo stesso anche per la prima colonna
                for col_idx, value in enumerate(row):
                    value = float(value)
                    if isinstance(value, (int, float)):
                        worksheet.write_number(row_idx + start_row + 1, col_idx + 2, value, number_format) # inserisci i valori con il formato dei numeri
                    else:
                        worksheet.write(row_idx + start_row + 1, col_idx + 2, value, number_format)

            # Inseriamo anche il flusso di cassa partendo da 2 righe inferiori al conto economico
            start_row += len(cont_eco) + 2
            if self.lingua == "ENG":
                titolo = "Cash Flows"
            elif self.lingua == "ITA":
                titolo = "Flussi di Cassa"
            worksheet.merge_range(start_row, 0, start_row, 1, titolo, subtitle_format) # come prima gli dico di mettere insieme le celle per scrivere il titolo
            flussi_mon = self.dfs_flussi_monetari[i] # prendimi il corrispettivo flusso di cassa
            worksheet.write(start_row, 1, '', header_format)   # questa cella la voglio vuota
            for col_num, value in enumerate(flussi_mon.columns):
                worksheet.write(start_row, col_num + 2, value, header_format) # qui la prima riga nel formato richiesto
            for row_idx, (index, row) in enumerate(flussi_mon.iterrows()):
                worksheet.write(row_idx + start_row + 1, 1, index, header_format) # qui la prima colonna nel formato richiesto
                for col_idx, value in enumerate(row):
                    try:
                        value = float(value) # trasformami il valore in un numero
                    except ValueError:
                        pass
                    if isinstance(value, (int, float)):
                        worksheet.write_number(row_idx + start_row + 1, col_idx + 2, value, number_format) # inseriscimi nella cella come un numero
                    else:
                        worksheet.write(row_idx + start_row + 1, col_idx + 2, value, number_format)

            ore = [el for el in range(8760)] # disegno una riga con i valori delle ore

            # Qui praticamente salvo tutti i flussi energetici del progetto
            data1 = self.dfs_flussi_energetici[i].loc['Energia in elettrolizzatore per ogni ora'].tolist()
            data2 = self.dfs_flussi_energetici[i].loc['Energia immessa in rete per ogni ora'].tolist()
            data3 = self.dfs_flussi_energetici[i].loc['Energia prodotta dal fotovoltaico'].tolist()
            data4 = self.dfs_flussi_energetici[i].loc['Potenza minima elettrolizzatore'].tolist()
            data5 = self.dfs_flussi_energetici[i].loc['Potenza massima elettrolizzatore'].tolist()
            if self.batteria == "SI":
                data6 = self.dfs_flussi_energetici[i].loc['Energia in batteria ora per ora'].tolist()
                data7 = self.dfs_flussi_energetici[i].loc['Potenza massima batteria'].tolist()
                """
                data9 = self.dfs_flussi_energetici[i].loc['Energia disponibile'].tolist()
                data10 = self.dfs_flussi_energetici[i].loc['Max erogabile'].tolist()"""
            worksheet.write_row("C47", ore, header_format)
            if self.lingua == "ENG":
                titolo = "Hourly Energy Flows"
            elif self.lingua == "ITA":
                titolo = "Flussi energetici orari"
            worksheet.merge_range(46, 0, 46, 1, titolo, subtitle_format) # inserisco come prima il titolo nel formato richiesto
            # nelle celle inferiori inserisco i valori e i titoli dei flussi
            worksheet.write("B47", titolo, header_format)
            worksheet.write_row("C48", data1)
            if self.lingua == "ENG":
                titolo = "Energy used by electrolyzer [kWh]"
            elif self.lingua == "ITA":
                titolo = "Energia usata dall'elettrolizzatore [kWh]"
            worksheet.write("B48", titolo, header_format)
            worksheet.write_row("C49", data2)
            if self.lingua == "ENG":
                titolo = "Energy fed into the grid [kWh]"
            elif self.lingua == "ITA":
                titolo = "Energia immessa in rete [kWh]"
            worksheet.write("B49", titolo, header_format)
            worksheet.write_row("C50", data3)
            if self.lingua == "ENG":
                titolo = "PV production [kWh]"
            elif self.lingua == "ITA":
                titolo = "Prodotta FV [kWh]"
            worksheet.write("B50", titolo, header_format)
            worksheet.write_row("C51", data4)
            if self.lingua == "ENG":
                titolo = "Electrolyzer cutoff power [kW]"
            elif self.lingua == "ITA":
                titolo = "Potenza minima elettrolizzatore [kW]"
            worksheet.write("B51", titolo, header_format)
            worksheet.write_row("C52", data5)
            if self.lingua == "ENG":
                titolo = "Electrolyzer nominal power [kW]"
            elif self.lingua == "ITA":
                titolo = "Taglia elettrolizzatore [kW]"
            worksheet.write("B52", titolo, header_format)
            if self.batteria == "SI":
                worksheet.write_row("C53", data6)
                if self.lingua == "ENG":
                    titolo = "Energy in battery net of discharge efficiency [kWh]"
                elif self.lingua == "ITA":
                    titolo = "Energia in batteria al netto dell'efficienza di scarica [kWh]"
                worksheet.write("B53", titolo, header_format)
                worksheet.write_row("C54", data7)
                if self.lingua == "ENG":
                    titolo = "Battery nominal power [kW]"
                elif self.lingua == "ITA":
                    titolo = "Taglia batteria [kW]"
                worksheet.write("B54", titolo, header_format)
                """worksheet.write_row("C55", data9)
                if self.lingua == "ENG":
                    titolo = "Available energy [kWh]"
                elif self.lingua == "ITA":
                    titolo = "Energia disponibile [kWh]"
                worksheet.write("B55", titolo, header_format)
                worksheet.write_row("C56", data10)
                if self.lingua == "ENG":
                    titolo = "Maximum deliverable power [kW]"
                elif self.lingua == "ITA":
                    titolo = "Massima potenza erogabile [kW]"
                worksheet.write("B56", titolo, header_format)"""

            # nella pagina Project Analysis + n vado a creare una grafico i cui dati li prende da ciò che ho appena inserito nel file excel
            sheet_name = f"'PA {project_num}'"
            chart = workbook.add_chart({'type': 'line'})
            chart.add_series({
                'name':       f"={sheet_name}!$B$50",
                'categories': f"={sheet_name}!$C$47:$LXZ$47",
                'values':     f"={sheet_name}!$C$50:$LXZ$50",
                'line':       {'color': '#92D050'} #verde prodotta
            })
            chart.add_series({
                'name':       f"={sheet_name}!$B$49",
                'categories': f"={sheet_name}!$C$47:$LXZ$47",
                'values':     f"={sheet_name}!$C$49:$LXZ$49",
                'line':       {'color': '#FFC000'} #giallo immessa
            })
            chart.add_series({
                'name':       f"={sheet_name}!$B$48",
                'categories': f"={sheet_name}!$C$47:$LXZ$47",
                'values':     f"={sheet_name}!$C$48:$LXZ$48",
                'line':       {'color': '#00AF50'} #verde autoconsumo
            })
            chart.add_series({
                'name':       f"={sheet_name}!$B$52",
                'categories': f"={sheet_name}!$C$47:$LXZ$47",
                'values':     f"={sheet_name}!$C$52:$LXZ$52",
                'line':       {'color': '#767171','dash_type': 'dash','transparency': 50}
            })
            chart.add_series({
                'name':       f"={sheet_name}!$B$51",
                'categories': f"={sheet_name}!$C$47:$LXZ$47",
                'values':     f"={sheet_name}!$C$51:$LXZ$51",
                'line':       {'color': '#767171','dash_type': 'dash','transparency': 50}
            })
            if self.batteria == "SI":
                chart.add_series({
                    'name':       f"={sheet_name}!$B$53",
                    'categories': f"={sheet_name}!$C$47:$LXZ$47",
                    'values':     f"={sheet_name}!$C$53:$LXZ$53",
                    'line':       {'color': 'orange'}
                })
                chart.add_series({
                    'name':       f"={sheet_name}!$B$54",
                    'categories': f"={sheet_name}!$C$47:$LXZ$47",
                    'values':     f"={sheet_name}!$C$54:$LXZ$54",
                    'line':       {'color': 'orange','dash_type': 'dash','transparency': 50}
                })
                """chart.add_series({
                    'name':       f"={sheet_name}!$B$55",
                    'categories': f"={sheet_name}!$C$47:$LXZ$47",
                    'values':     f"={sheet_name}!$C$55:$LXZ$55",
                    'line':       {'color': 'magenta','dash_type': 'dash','transparency': 50}
                })
                chart.add_series({
                    'name':       f"={sheet_name}!$B$56",
                    'categories': f"={sheet_name}!$C$47:$LXZ$47",
                    'values':     f"={sheet_name}!$C$56:$LXZ$56",
                    'line':       {'color': 'magenta','dash_type': 'dash','transparency': 50}
                })"""
            # Qui vado a definire la legenda del grafico
            if self.lingua == "ENG":
                titolo1 = "Energy Flow Chart"
                titolo2 = "Hour"
                titolo3 = "Quantity"
            elif self.lingua == "ITA":
                titolo1 = 'Grafico dei Flussi Energetici'
                titolo2 = 'Ora'
                titolo3 = 'Quantità'
            chart.set_title({
                'name': titolo1,
                'align': 'left',
                'name_font': {
                    'name': 'Rajdhani Medium',  # Font per il titolo del grafico
                    'size': 18,  # Dimensione del font
                    'bold': True,
                    'color': color_palette["font_color"]  # Colore del titolo (nero di default)
                }
            })

            # Definizione dell'asse X
            chart.set_x_axis({
                'name': titolo2,
                'name_font': {
                    'name': 'Rajdhani Medium',  # Font per il titolo dell'asse X
                    'size': 13,  # Dimensione del font
                    'bold': True,
                    'color': color_palette["font_color"]
                },
                'num_font': {
                    'name': 'Roboto',  # Font per i valori dell'asse X
                    'size': 9, # Dimensione del font per i valori
                    'color': color_palette["font_color"]
                }
            })

            # Definizione dell'asse Y
            chart.set_y_axis({
                'name': titolo3,
                'name_font': {
                    'name': 'Rajdhani Medium',  # Font per il titolo dell'asse Y
                    'size': 13,  # Dimensione del font
                    'bold': True,
                    'color': color_palette["font_color"]
                },
                'num_font': {
                    'name': 'Roboto',  # Font per i valori dell'asse Y
                    'size': 9,  # Dimensione del font per i valori
                    'color': color_palette["font_color"]
                }
            })

            # Impostazioni della legenda e della dimensione del grafico
            chart.set_size({'width': 100000, 'height': 700})
            chart.set_legend({
                'position': 'left',
                'font': {
                    'name': 'Rajdhani Medium',  # Font per la leggenda
                    'size': 12,
                    'color': color_palette["font_color"]
                }
            })
            # in base a quanto è lungo (quindi in base al fatto se ho meno c'è la batteria) disegno il grafico in un punto diverso
            if self.batteria == "NO":
                worksheet.insert_chart('B54', chart)
            elif self.batteria == "SI":
                worksheet.insert_chart('B58', chart)

            if self.lingua == "ENG":
                self.print_progress_bar(i+1,num_projects, "", "Output file compilation",1,50)
            elif self.lingua == "ITA":
                self.print_progress_bar(i+1,num_projects, "", "Compilazione file output",1,50)

        if self.lingua == "ENG":
            print("Saving...")
        elif self.lingua == "ITA":
            print("Salvataggio file...")
        workbook.close()


    #da qui in poi ci sono le funzioni per la barra di completamento
    def interpolate_color(self, start_rgb, end_rgb, factor):
        return tuple(int(start + (end - start) * factor) for start, end in zip(start_rgb, end_rgb))
    def rgb_to_ansi(self, rgb):
        return f'\033[38;2;{rgb[0]};{rgb[1]};{rgb[2]}m'
    def print_progress_bar(self, iteration, total, prefix='', suffix="Completamento", decimals=1, length=50, fill='█'):
        start_color = (244, 7, 254)  # #F407FE (Magenta)
        end_color = (7, 213, 223)  # #07D5DF (Cyan)

        percent = iteration / float(total)
        filled_length = int(length * percent)

        bar = ""
        for i in range(length):
            color = self.interpolate_color(start_color, end_color, i / length)
            bar += self.rgb_to_ansi(color) + (fill if i < filled_length else '-')

        sys.stdout.write(f'\r{prefix} |{bar}\033[0m| {percent * 100:.{decimals}f}% {suffix}')
        sys.stdout.flush()

        if iteration == total:
            print()

file_path = "INPUT.xlsx"
dati = Analisi_combinata(file_path)
