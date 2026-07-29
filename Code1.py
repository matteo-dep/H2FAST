import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import xlsxwriter
import warnings
import sys
import io

class Analisi_tecnica:
    def __init__(self, file_csv, tipo_file, p_PV, dP_el, batteria, dP_bat, min_batt, max_batteria, lingua, eff_batt, min_elet, progress_bar=None, status_text=None):
        self.batteria = batteria 
        self.file_csv = file_csv 
        self.tipo_file = tipo_file 
        self.p_PV = p_PV  
        self.dP_el = dP_el 
        self.dP_bat = dP_bat 
        self.lingua = lingua 
        self.progress_bar = progress_bar
        self.status_text = status_text

        self.E_PV = self.load_data()  
        self.min_batt = min_batt  
        self.max_batteria = max_batteria  
        self.min_elet = min_elet  
        self.e_pv = np.sum(self.E_PV)  
        
        gran_elc = 0
        if self.p_PV <= 100: gran_elc = 10
        elif 100 < self.p_PV <= 500: gran_elc = 20
        elif 500 < self.p_PV <= 1000: gran_elc = 50
        elif self.p_PV > 1000: gran_elc = 100
        self.P_elc = np.linspace(self.p_PV / gran_elc, self.p_PV*dP_el, gran_elc)

        self.eff_batt = eff_batt  
        
        if self.batteria == "NO":
            self.qt_progetti = len(self.P_elc)  
            self.andamenti = np.zeros((self.qt_progetti, 6,len(self.E_PV)))  
            self.E_H2 = np.zeros(self.qt_progetti)  
            self.E_im = np.zeros(self.qt_progetti)  
            self.M_H2 = np.zeros(self.qt_progetti)  
            self.Auto = np.zeros(self.qt_progetti)  
            self.CF = np.zeros(self.qt_progetti)  
            self.OFFgiorn = np.zeros(self.qt_progetti)  
            self.OFF = np.zeros(self.qt_progetti)  
            self.spegn_giorn = np.zeros(self.qt_progetti)
            
        if self.batteria == "SI":
            len_batt = 0  
            for el in self.P_elc:  
                gran_bat = 0
                if el <= 100: gran_bat = 10
                elif 100 < el <= 500: gran_bat = 20
                elif 500 < el <= 1000: gran_bat = 50
                elif el > 1000: gran_bat = 100
                P_batt = np.linspace(el * self.dP_bat / gran_bat, el * self.dP_bat, gran_bat)
                len_batt += len(P_batt)  

            self.qt_progetti = len_batt
            self.andamenti = np.zeros((self.qt_progetti, 11,len(self.E_PV)))  
            self.E_H2 = np.zeros(self.qt_progetti)  
            self.E_im = np.zeros(self.qt_progetti)  
            self.M_H2 = np.zeros(self.qt_progetti)  
            self.Auto = np.zeros(self.qt_progetti)  
            self.CF = np.zeros(self.qt_progetti)  
            self.OFFgiorn = np.zeros(self.qt_progetti)  
            self.OFF = np.zeros(self.qt_progetti)  
            self.potenza_batt = np.zeros(self.qt_progetti)  
            self.potenza_elett = np.zeros(self.qt_progetti )  
            self.spegn_giorn = np.zeros(self.qt_progetti)  

        self.off = 0
        self.flag = 0
        self.count_to_24 = 0
        self.count2 = 0

    def load_data(self):
        if self.tipo_file == "SI":
            lines = [line.decode("utf-8") for line in self.file_csv.readlines()]
            self.file_csv.seek(0)
            header = "time,P,G(i),H_sun,T2m,WS10m,Int"  
            header_idx = next((i for i, line in enumerate(lines) if line.strip() == header),None)  
            if header_idx is None:
                raise ValueError("Errore file PVGIS" if self.lingua=="ITA" else "Error PVGIS file")
            rows = []
            for line in lines[header_idx + 1:]:  
                columns = line.strip().split(',')
                if len(columns) == 7: rows.append(columns)
                else: break  
            data = pd.DataFrame(rows, columns=header.split(','))  
            E_PV = pd.to_numeric(data['P'], errors='coerce').dropna()  
            E_PV /= 1000  
            return E_PV.to_numpy()
        if self.tipo_file == "NO":  
            data = pd.read_csv(self.file_csv, sep=';',decimal=',')  
            E_PV = pd.to_numeric(data.iloc[:, 0],errors='coerce').dropna()  
            return E_PV.to_numpy()

    @staticmethod
    def eff_elc(x):  
        y = (-6.1371 * x ** 6 + 24.394 * x ** 5 - 39.663 * x ** 4 + 33.988 * x ** 3 - 16.412 * x ** 2 + 4.2929 * x + 0.1022)
        return y

    def run_analysis_nobattery(self):  
        for index, p_elc in enumerate(self.P_elc):  
            p_elc_min = self.min_elet * p_elc  
            E_H2_h = np.zeros(len(self.E_PV))  
            M_H2_h = np.zeros(len(self.E_PV))  
            E_im_h = np.zeros(len(self.E_PV))  
            P_pv_h = np.zeros(len(self.E_PV))  
            Pot_min_elett = np.full(len(self.E_PV),p_elc_min)  
            Pot_max_elett = np.full(len(self.E_PV),p_elc)  

            self.off = 0  
            self.flag = 1  
            self.count_to_24 = 0  
            self.count2 = 0  

            for i, p_pv in enumerate(self.E_PV):  
                if p_pv > p_elc:  
                    e_H2 = p_elc  
                    e_im = p_pv - e_H2  
                    m_H2 = e_H2 * 0.565 / (120 / 3.6)  
                    self.flag = 0  
                    self.count_to_24 = 0  
                elif p_elc_min <= p_pv < p_elc:  
                    e_H2 = p_pv  
                    e_im = 0  
                    var = p_pv / p_elc  
                    eta = self.eff_elc(var)
                    m_H2 = e_H2 * eta / (120 / 3.6)  
                    self.flag = 0  
                    self.count_to_24 = 0  
                elif p_pv <= p_elc_min:  
                    e_H2 = 0  
                    e_im = p_pv  
                    m_H2 = 0  
                    if self.flag == 0:  
                        self.off += 1  
                        self.flag = 1  
                    else:  
                        self.flag = 1  
                    self.count_to_24 += 1  
                    if self.count_to_24 == 24:  
                        self.count2 += 1  
                        self.count_to_24 = 0  

                E_H2_h[i] = e_H2  
                M_H2_h[i] = m_H2  
                E_im_h[i] = e_im  
                P_pv_h[i] = p_pv  

            e_H2 = np.sum(E_H2_h)  
            e_im = np.sum(E_im_h)  
            m_H2 = np.sum(M_H2_h)  
            spegn_giorn_proj = self.count2 + self.off  
            e_TOT = len(self.E_PV) * p_elc  
            cf = e_H2 / e_TOT * 100  
            autoconsumo = e_H2 / self.e_pv * 100  

            self.OFF[index] = self.off  
            self.spegn_giorn[index] = spegn_giorn_proj  
            self.CF[index] = cf  
            self.Auto[index] = autoconsumo  
            self.E_H2[index] = e_H2  
            self.E_im[index] = e_im  
            self.M_H2[index] = m_H2  

            self.andamenti[index, 0, :] = E_H2_h
            self.andamenti[index, 1, :] = M_H2_h
            self.andamenti[index, 2, :] = E_im_h
            self.andamenti[index, 3, :] = P_pv_h
            self.andamenti[index, 4, :] = Pot_min_elett
            self.andamenti[index, 5, :] = Pot_max_elett

            if self.progress_bar is not None and self.status_text is not None:
                pct = (index + 1) / len(self.P_elc)
                self.progress_bar.progress(pct)
                msg = "Technical Analysis without Battery" if self.lingua == "ENG" else "Analisi Tecnica senza batteria"
                self.status_text.text(f"{pct * 100:.1f}% - {msg}")

    def run_analysis_battery_static_min(self):  
        index = 0  
        for p_elc in self.P_elc:  
            p_elc_min = self.min_elet * p_elc  
            gran_bat = 0
            if p_elc <= 100: gran_bat = 10
            elif 100 < p_elc <= 500: gran_bat = 20
            elif 500 < p_elc <= 1000: gran_bat = 50
            elif p_elc > 1000: gran_bat = 100
            P_batt = np.linspace(0, self.dP_bat * p_elc, gran_bat) 

            for p_batt in P_batt: 
                e_batt_max = p_batt*self.max_batteria  
                E_H2_h = np.zeros(len(self.E_PV))  
                M_H2_h = np.zeros(len(self.E_PV))  
                E_im_h = np.zeros(len(self.E_PV))  
                E_batt_disponibile_h = np.zeros(len(self.E_PV)) 
                E_batt = np.zeros(len(self.E_PV)) 
                P_pv_h = np.zeros(len(self.E_PV)) 
                Energia_disp = np.zeros(len(self.E_PV)) 
                Pot_min_elett = np.full(len(self.E_PV), p_elc_min) 
                Pot_max_elett = np.full(len(self.E_PV), p_elc) 
                Pot_max_batt = np.full(len(self.E_PV), p_batt) 
                max_erogabile = np.zeros(len(self.E_PV)) 
                
                self.off = 0 
                self.flag = 1  
                self.count_to_24 = 0 
                self.count2 = 0  
                j = 0 

                for i,p_pv in enumerate(self.E_PV): 
                    e_im = 0
                    if i == 0: energia_disp_batt = 0
                    elif i != 0: energia_disp_batt = max(0, E_batt[i - 1] - self.min_batt*p_batt) 
                    
                    energia_disp = p_pv + self.eff_batt*energia_disp_batt 

                    if p_pv == 0: 
                        if energia_disp < p_elc_min: 
                            e_H2 = 0; m_H2 = 0 
                            if i == 0: E_batt[0] = 0; j += 1
                            elif i != 0: E_batt[i] = E_batt[i-1]; j += 1
                            self.count_to_24 += 1 
                            if self.flag == 0: self.off += 1
                            self.flag = 1 
                            if self.count_to_24 == 24: 
                                self.count2 += 1; self.count_to_24 = 0 
                        elif p_elc_min <= energia_disp < p_elc: 
                            e_H2 = min(energia_disp, e_batt_max * self.eff_batt) 
                            eta = self.eff_elc(e_H2 / p_elc)
                            m_H2 = e_H2 * eta / (120 / 3.6)
                            E_batt[i] = E_batt[i - 1] - e_H2 / self.eff_batt
                            j += 1; self.flag = 0; self.count_to_24 = 0 
                        elif p_elc <= energia_disp: 
                            e_H2 = min(p_elc, e_batt_max, energia_disp_batt)*self.eff_batt 
                            eta = self.eff_elc(e_H2 / p_elc)
                            m_H2 = e_H2 * eta / (120 / 3.6)  
                            E_batt[i] = E_batt[i - 1] - e_H2 / self.eff_batt
                            j += 1; self.flag = 0; self.count_to_24 = 0 

                    elif 0 < p_pv < p_elc_min: 
                        if energia_disp < p_elc_min: 
                            e_H2 = 0; m_H2 = 0; j += 1
                            if i == 0: E_batt[i] = min(p_pv, e_batt_max) * self.eff_batt
                            elif i != 0: E_batt[i] = min((min(p_pv, e_batt_max)*self.eff_batt + E_batt[i-1]), p_batt )
                            e_im = max(p_pv - (E_batt[i] - E_batt[i-1])/self.eff_batt, 0)
                            self.count_to_24 += 1  
                            if self.flag == 0: self.off += 1
                            self.flag = 1  
                            if self.count_to_24 == 24: self.count2 += 1; self.count_to_24 = 0  
                        elif p_elc_min <= energia_disp < p_elc: 
                            e_H2 = min(energia_disp, e_batt_max*self.eff_batt + p_pv) 
                            eta = self.eff_elc(e_H2 / p_elc)
                            m_H2 = e_H2 * eta / (120 / 3.6)
                            if i == j:
                                if e_H2 == energia_disp: E_batt[i] = self.min_batt*e_batt_max
                                elif e_H2 == e_batt_max*self.eff_batt + p_pv: E_batt[i] = E_batt[i-1] - e_batt_max
                            j += 1; self.flag = 0; self.count_to_24 = 0 
                        elif p_elc <= energia_disp: 
                            e_H2 = min(e_batt_max + p_pv, p_elc) 
                            if i==j:
                                if e_H2 == e_batt_max + p_pv: 
                                    eta = self.eff_elc(e_H2 / p_elc)
                                    m_H2 = e_H2 * eta / (120 / 3.6)
                                    E_batt[i] = E_batt[i-1] - e_batt_max
                                elif e_H2 == p_elc:
                                    m_H2 = e_H2 * 0.565 / (120 / 3.6)
                                    E_batt[i] = E_batt[i-1] - (p_elc - p_pv)/self.eff_batt
                            j += 1; self.flag = 0; self.count_to_24 = 0 

                    elif p_elc_min <= p_pv < p_elc: 
                        if p_elc_min <= energia_disp < p_elc:
                            e_H2 = min(energia_disp, e_batt_max*self.eff_batt + p_pv)
                            eta = self.eff_elc(e_H2 / p_elc)
                            m_H2 = e_H2 * eta / (120 / 3.6)
                            if i == j:
                                if e_H2 == energia_disp: E_batt[i] = E_batt[i-1] - (energia_disp  - p_pv) / self.eff_batt
                                elif e_H2 == e_batt_max*self.eff_batt + p_pv: E_batt[i] = E_batt[i-1] - e_batt_max
                            j += 1; self.flag = 0; self.count_to_24 = 0 
                        elif p_elc <= energia_disp:
                            e_H2 = min(e_batt_max * self.eff_batt + p_pv, p_elc)
                            if i == j:
                                if e_H2 == e_batt_max * self.eff_batt + p_pv:
                                    eta = self.eff_elc(e_H2 / p_elc)
                                    m_H2 = e_H2 * eta / (120 / 3.6)
                                    E_batt[i] = E_batt[i-1] - e_batt_max
                                elif e_H2 == p_elc:
                                    m_H2 = e_H2 * 0.565 / (120 / 3.6)
                                    E_batt[i] = E_batt[i-1] - (p_elc - p_pv)/self.eff_batt
                            j += 1; self.flag = 0; self.count_to_24 = 0 

                    elif p_elc <= p_pv:
                        e_H2 = p_elc
                        m_H2 = e_H2 * 0.565 / (120 / 3.6)
                        E_batt[i] = E_batt[i-1] + min((p_pv - p_elc) * self.eff_batt, p_batt - E_batt[i - 1], e_batt_max)
                        e_im = p_pv - p_elc - (E_batt[i] - E_batt[i-1]) / self.eff_batt
                        j += 1; self.flag = 0; self.count_to_24 = 0 

                    E_H2_h[i] = e_H2; M_H2_h[i] = m_H2; E_im_h[i] = e_im  
                    E_batt_disponibile_h[i] = energia_disp_batt
                    Energia_disp[i] = energia_disp; P_pv_h[i] = p_pv
                    max_erogabile[i] = e_batt_max*self.eff_batt + p_pv

                e_H2_ = np.sum(E_H2_h); e_im_ = np.sum(E_im_h); m_H2_ = np.sum(M_H2_h)
                spegn_giorn_proj = self.off + self.count2 
                e_TOT = len(self.E_PV) * p_elc
                cf = e_H2_ / e_TOT * 100
                autoconsumo = e_H2_ / self.e_pv * 100

                self.spegn_giorn[index] = spegn_giorn_proj; self.OFF[index] = self.off
                self.potenza_elett[index] = p_elc; self.potenza_batt[index] = p_batt
                self.CF[index] = cf; self.Auto[index] = autoconsumo
                self.E_H2[index] = e_H2_; self.E_im[index] = e_im_; self.M_H2[index] = m_H2_
                self.andamenti[index, 0, :] = E_H2_h; self.andamenti[index, 1, :] = M_H2_h
                self.andamenti[index, 2, :] = E_im_h; self.andamenti[index, 3, :] = P_pv_h
                self.andamenti[index, 4, :] = Pot_min_elett; self.andamenti[index, 5, :] = Pot_max_elett
                self.andamenti[index, 6, :] = E_batt; self.andamenti[index, 7, :] = Pot_max_batt
                self.andamenti[index, 8, :] = E_batt_disponibile_h; self.andamenti[index, 9, :] = Energia_disp
                self.andamenti[index, 10, :] = max_erogabile
                
                if self.progress_bar is not None and self.status_text is not None:
                    pct = (index + 1) / self.qt_progetti
                    self.progress_bar.progress(pct)
                    msg = "Technical Analysis with Battery" if self.lingua == "ENG" else "Analisi Tecnica con batteria"
                    self.status_text.text(f"{pct * 100:.1f}% - {msg}")
                index += 1

    def run_analysis(self):
        if self.batteria == "SI": self.run_analysis_battery_static_min()
        elif self.batteria == "NO": self.run_analysis_nobattery()

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

class Analisi_combinata_Streamlit:
    def __init__(self, parametri, progress_bar, status_text): 
        # Carichiamo i parametri inseriti dall'utente nella UI
        for key, value in parametri.items():
            setattr(self, key, value)
            
        self.analisi1 = Analisi_tecnica(
            file_csv=self.file_csv, tipo_file=self.tipo_file, p_PV=self.p_PV, 
            dP_el=self.dP_el, batteria=self.batteria, dP_bat=self.dP_bat,
            min_batt=self.min_batt, max_batteria=self.max_batteria, lingua=self.lingua, 
            eff_batt=self.eff_batt, min_elet=self.min_elet, 
            progress_bar=progress_bar, status_text=status_text
        )
        # 1. Eseguiamo l'analisi tecnica
        self.analisi1.run_analysis() 
        self.istanze_finanziarie = []
        
        status_text.text("Avvio Analisi Finanziaria...")
        
        # 2. Generiamo l'analisi finanziaria per ogni scenario tecnico calcolato
        tot_progetti = len(self.analisi1.P_elc) if self.batteria == "NO" else len(self.analisi1.potenza_batt)
        
        if self.batteria == "NO":
            for index, (el1, el2, el3, el4, el5, el6, el7) in enumerate(zip(self.analisi1.CF, self.analisi1.P_elc, self.analisi1.E_H2, self.analisi1.E_im, self.analisi1.M_H2, self.analisi1.andamenti, self.analisi1.spegn_giorn)):
                analisi2 = Analisi_finanziaria(Terr=self.Terr, OpeE=self.OpeE, ImpPV1=self.p_PV, ImpPV1eurokW=self.ImpPV1eurokW, EletteuroKW=self.EletteuroKW, CompreuroKW=self.CompreuroKW, AccuE=0, AccuEeurokW=self.AccuEeurokW, idrogstocperc=self.idrogstocperc, StazzRif=self.StazzRif, SpeTOpere=self.SpeTOpere, BombSto=self.BombSto, LavoImp=self.LavoImp, CarrEll=self.CarrEll, CapFac=el1, PotEle=el2, tassoDEN=self.tassoDEN, ProdAnnuaIdrogkg=el5, bar=self.bar, costlitroacqua=self.costlitroacqua, costounitariostoccaggio=self.costounitariostoccaggio, PercEserImp=self.PercEserImp, Percentimpianti=self.Percentimpianti, PercentOpeEd=self.PercentOpeEd, SpesAmmGen=self.SpesAmmGen, Affitto=self.Affitto, CostiPersonal=self.CostiPersonal, AltriCost=self.AltriCost, IVAsualtriCost=self.IVAsualtriCost, DurPianEcon=self.DurPianEcon, inflazione=self.inflazione, inflazionePrezzoElet=self.inflazionePrezzoElet, inflazioneIdrog=self.inflazioneIdrog, tassoVAN=self.tassoVAN, incentpubb=self.incentpubb, duratincentpubb=self.duratincentpubb, prezzoindrogeno=self.prezzoindrogeno, ProdElettVend=el4, EnergiaAutocons=el3, prezzoElett=self.prezzoElett, ContrPubb=self.ContrPubb, DebitoSenior=self.DebitoSenior, DurDebitoSenior=self.DurDebitoSenior, tassoDebito=self.tassoDebito, FreqPagamenti=self.FreqPagamenti, tassoPonte=self.tassoPonte, DurataPonte=self.DurataPonte, aliquoMedia=self.aliquoMedia, MaxInterssDed=self.MaxInterssDed, lingua=self.lingua, Perciva=self.Perciva, spegn_giorn=el7)
                analisi2.RUN()
                self.istanze_finanziarie.append(analisi2)
                progress_bar.progress((index + 1) / tot_progetti)
                
        elif self.batteria == "SI":
            for index, (el1, el2, el3, el4, el5, el6, el7, el8) in enumerate(zip(self.analisi1.potenza_batt, self.analisi1.CF, self.analisi1.potenza_elett, self.analisi1.E_H2, self.analisi1.E_im, self.analisi1.M_H2, self.analisi1.andamenti, self.analisi1.spegn_giorn)):
                analisi2 = Analisi_finanziaria(Terr=self.Terr, OpeE=self.OpeE, ImpPV1=self.p_PV, ImpPV1eurokW=self.ImpPV1eurokW, EletteuroKW=self.EletteuroKW, CompreuroKW=self.CompreuroKW, AccuE=el1, AccuEeurokW=self.AccuEeurokW, idrogstocperc=self.idrogstocperc, StazzRif=self.StazzRif, SpeTOpere=self.SpeTOpere, BombSto=self.BombSto, LavoImp=self.LavoImp, CarrEll=self.CarrEll, CapFac=el2, PotEle=el3, tassoDEN=self.tassoDEN, ProdAnnuaIdrogkg=el6, bar=self.bar, costlitroacqua=self.costlitroacqua, costounitariostoccaggio=self.costounitariostoccaggio, PercEserImp=self.PercEserImp, Percentimpianti=self.Percentimpianti, PercentOpeEd=self.PercentOpeEd, SpesAmmGen=self.SpesAmmGen, Affitto=self.Affitto, CostiPersonal=self.CostiPersonal, AltriCost=self.AltriCost, IVAsualtriCost=self.IVAsualtriCost, DurPianEcon=self.DurPianEcon, inflazione=self.inflazione, inflazionePrezzoElet=self.inflazionePrezzoElet, inflazioneIdrog=self.inflazioneIdrog, tassoVAN=self.tassoVAN, incentpubb=self.incentpubb, duratincentpubb=self.duratincentpubb, prezzoindrogeno=self.prezzoindrogeno, ProdElettVend=el5, EnergiaAutocons=el4, prezzoElett=self.prezzoElett, ContrPubb=self.ContrPubb, DebitoSenior=self.DebitoSenior, DurDebitoSenior=self.DurDebitoSenior, tassoDebito=self.tassoDebito, FreqPagamenti=self.FreqPagamenti, tassoPonte=self.tassoPonte, DurataPonte=self.DurataPonte, aliquoMedia=self.aliquoMedia, MaxInterssDed=self.MaxInterssDed, lingua=self.lingua, Perciva=self.Perciva, spegn_giorn=el8)
                analisi2.RUN()
                self.istanze_finanziarie.append(analisi2)
                progress_bar.progress((index + 1) / tot_progetti)

        # 3. Ordiniamo i progetti dal VAN più alto al più basso
        self.istanze_finanziarie.sort(key=lambda x: x.VAN if x.VAN is not None else -float('inf'), reverse=True)
        # Teniamo solo i "top N" richiesti dall'utente
        self.top_progetti = self.istanze_finanziarie[:self.n_progetti]

        # 4. Prepariamo le liste per le tabelle di Streamlit
        self.dfs_conto_economico = [proj.dfContoEconomico for proj in self.top_progetti]
        self.dfs_flussi_monetari = [proj.dfFlussiMonetari for proj in self.top_progetti]

        # 5. Generiamo la Tabella Max riassuntiva
        data_sommario = {
            'VAN Progetto [€]': [p.VAN for p in self.top_progetti],
            'TIR Progetto': [p.TIR for p in self.top_progetti],
            'PAYBACK [anni]': [p.PAYBACK for p in self.top_progetti],
            'Potenza Elettrolizzatore [kW]': [p.PotEle for p in self.top_progetti],
            'Produzione H2 [kg/anno]': [p.ProdAnnuaIdrogkg for p in self.top_progetti],
            'Investimento iniziale [€]': [p.investimento for p in self.top_progetti]
        }
        if self.batteria == "SI":
            data_sommario['Dimensione Batteria [kWh]'] = [p.AccuE for p in self.top_progetti]

        self.TabellaMax = pd.DataFrame(data_sommario).T
        self.TabellaMax.columns = [f'Progetto {i+1}' for i in range(len(self.top_progetti))]

# ==========================================
# INTERFACCIA WEB (UI)
# ==========================================
st.set_page_config(page_title="H2FAsT Simulator", layout="wide")
st.title("Simulatore H2FAsT - Impianti Idrogeno 🏭")

# --- BARRA LATERALE PER GLI INPUT ---
st.sidebar.header("📁 Dati Input")
lingua = st.sidebar.selectbox("Lingua (Language)", ["ITA", "ENG"])
file_csv = st.sidebar.file_uploader("Carica file serie oraria CSV", type=["csv"])
tipo_file = st.sidebar.selectbox("Il file proviene da PVGIS?", ["SI", "NO"])

st.sidebar.header("⚙️ Parametri Tecnici")
p_PV = st.sidebar.number_input("PV system peak power [kWp]", value=1000.0)
dP_el = st.sidebar.number_input("Electrolyzer range limit", value=1.0)
min_elet = st.sidebar.number_input("% Electrolyzer power cutoff", value=0.2)
batteria = st.sidebar.selectbox("Presence of battery in the system", ["SI", "NO"])
dP_bat = st.sidebar.number_input("Battery range limit", value=2.0) if batteria == "SI" else 0.0
min_batt = st.sidebar.number_input("Cutoff SoC [%]", value=0.2)
max_batteria = st.sidebar.number_input("Maximum energy % deliverable in 1h", value=0.5)
eff_batt = st.sidebar.number_input("% Charge and discharge battery efficiency", value=0.95)

st.sidebar.header("💶 CAPEX & Costi Impianto")
Terr = st.sidebar.number_input("Land [€]", value=0.0)
OpeE = st.sidebar.number_input("Building works [€]", value=0.0)
StazzRif = st.sidebar.number_input("Filling station [€]", value=0.0)
ImpPV1eurokW = st.sidebar.number_input("PV system [€/kWp]", value=800.0)
EletteuroKW = st.sidebar.number_input("Electrolyzer [€/kW]", value=1650.0)
CompreuroKW = st.sidebar.number_input("Compressor [€/kW]", value=4000.0)
AccuEeurokW = st.sidebar.number_input("Batteries [€/kWh]", value=200.0)
costounitariostoccaggio = st.sidebar.number_input("Storage [€/kg]", value=1200.0)

st.sidebar.header("📊 Parametri Finanziari")
prezzoindrogeno = st.sidebar.number_input("Hydrogen sale price [€/kg]", value=20.0)
tassoVAN = st.sidebar.number_input("Discount rate for NPV calculation", value=0.07)
inflazione = st.sidebar.number_input("General inflation rate", value=0.01)
DurPianEcon = st.sidebar.number_input("Economic plan duration [years]", value=20, step=1)
n_progetti = st.sidebar.number_input("Numero di Top Progetti da visualizzare", value=5, min_value=1)

parametri_simulazione = {
    "lingua": lingua, "file_csv": file_csv, "tipo_file": tipo_file, "p_PV": p_PV,
    "dP_el": dP_el, "min_elet": min_elet, "batteria": batteria, "dP_bat": dP_bat,
    "min_batt": min_batt, "max_batteria": max_batteria, "eff_batt": eff_batt,
    "Terr": Terr, "OpeE": OpeE, "StazzRif": StazzRif, "SpeTOpere": 0, "BombSto": 0,
    "LavoImp": 0, "CarrEll": 0, "ImpPV1eurokW": ImpPV1eurokW, "EletteuroKW": EletteuroKW, 
    "CompreuroKW": CompreuroKW, "AccuEeurokW": AccuEeurokW, 
    "costounitariostoccaggio": costounitariostoccaggio, "idrogstocperc": 0.1,
    "costlitroacqua": 0.025, "PercEserImp": 0.005, "Percentimpianti": 0.0025,
    "PercentOpeEd": 0.0005, "SpesAmmGen": 2000, "Affitto": 0, "CostiPersonal": 0,
    "AltriCost": 0, "IVAsualtriCost": "NO", "DurPianEcon": DurPianEcon,
    "inflazione": inflazione, "inflazionePrezzoElet": 0.015, "tassoVAN": tassoVAN, 
    "incentpubb": 2.0, "duratincentpubb": 20, "prezzoindrogeno": prezzoindrogeno,
    "inflazioneIdrog": 0.01, "prezzoElett": 0.1, "ContrPubb": 0.9, "DebitoSenior": 0.0, 
    "DurDebitoSenior": 10, "tassoDebito": 0.05, "FreqPagamenti": 1, "DurataPonte": 0,
    "tassoPonte": 0, "aliquoMedia": 0.275, "MaxInterssDed": 0.3, "Perciva": 0.22,
    "tassoDEN": 0.005, "bar": 500, "n_progetti": n_progetti
}

# --- AVVIO PROCESSO ---
if st.button("🚀 Avvia Simulazione", use_container_width=True):
    if file_csv is None:
        st.warning("⚠️ Carica prima un file CSV con la Serie oraria di produzione PV!")
    else:
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        try:
            # Calcolo
            app_combinata = Analisi_combinata_Streamlit(parametri_simulazione, progress_bar, status_text)
            
            progress_bar.empty()
            status_text.empty()
            st.success("✅ Simulazione completata!")
            
            # DASHBOARD INTERATTIVA (TABS)
            tab1, tab2, tab3 = st.tabs(["🏆 Sommario Progetti", "📈 Conto Economico", "💶 Flussi di Cassa"])
            
            with tab1:
                st.subheader("Sommario delle migliori configurazioni")
                st.dataframe(app_combinata.TabellaMax, use_container_width=True)
                
            with tab2:
                st.subheader("Conto Economico")
                scelta_eco = st.selectbox("Seleziona il progetto:", options=range(n_progetti), format_func=lambda x: f"Progetto {x+1}", key="se")
                st.dataframe(app_combinata.dfs_conto_economico[scelta_eco], use_container_width=True)

            with tab3:
                st.subheader("Flussi di Cassa")
                scelta_cassa = st.selectbox("Seleziona il progetto:", options=range(n_progetti), format_func=lambda x: f"Progetto {x+1}", key="sc")
                st.dataframe(app_combinata.dfs_flussi_monetari[scelta_cassa], use_container_width=True)
                
        except Exception as e:
            st.error(f"❌ Errore durante l'esecuzione: {e}")
