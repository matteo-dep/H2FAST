"""
Profili orari delle fonti rinnovabili per H2FAST.

- Fotovoltaico da PVGIS (JRC, Commissione europea) - API seriescalc
- Eolico da Open-Meteo (archivio ERA5, vento a 100 m) + curva di potenza semplificata
- Profili caricati dall'utente (CSV PVGIS o CSV generico a una colonna in kW)

Tutte le funzioni restituiscono array numpy di 8760 valori orari in kW (anno non bisestile).
"""
import io

import numpy as np
import pandas as pd
import requests

URL_PVGIS = ["https://re.jrc.ec.europa.eu/api/v5_3/seriescalc",
             "https://re.jrc.ec.europa.eu/api/v5_2/seriescalc"]
URL_OPENMETEO = "https://archive-api.open-meteo.com/v1/archive"
ORE_ANNO = 8760
TIMEOUT = 60


class ErroreProfilo(Exception):
    """Errore leggibile dall'utente sul profilo orario."""


def a_8760(valori, anno=None):
    """Riporta una serie oraria a 8760 valori (toglie il 29 febbraio negli anni bisestili)."""
    v = np.asarray(valori, dtype=float)
    v = np.nan_to_num(v, nan=0.0)
    if len(v) == 8784:
        v = np.concatenate([v[:59 * 24], v[60 * 24:]])  # 29 febbraio = 60° giorno
    if len(v) != ORE_ANNO:
        raise ErroreProfilo(f"Il profilo ha {len(v)} valori orari: ne servono 8760 (un anno, ora per ora).")
    return np.clip(v, 0, None)


# ---------------------------------------------------------------------------------------------
# FOTOVOLTAICO
# ---------------------------------------------------------------------------------------------
def scarica_pvgis(lat, lon, kwp, anno=2019, perdite=14.0, angoli_ottimali=True, inclinazione=35.0, azimut=0.0):
    """Produzione FV oraria (kW) da PVGIS. azimut: 0 = sud, -90 = est, 90 = ovest (convenzione PVGIS)."""
    if kwp <= 0:
        return np.zeros(ORE_ANNO)
    params = {"lat": lat, "lon": lon, "startyear": int(anno), "endyear": int(anno), "pvcalculation": 1,
              "peakpower": kwp, "loss": perdite, "outputformat": "json"}
    if angoli_ottimali:
        params["optimalangles"] = 1
    else:
        params["angle"] = inclinazione
        params["aspect"] = azimut
    ultimo_errore = None
    for url in URL_PVGIS:
        try:
            r = requests.get(url, params=params, timeout=TIMEOUT)
            if r.status_code != 200:
                try:
                    ultimo_errore = r.json().get("message", r.text[:200])
                except ValueError:
                    ultimo_errore = r.text[:200]
                continue
            dati = r.json()
            orari = pd.DataFrame(dati["outputs"]["hourly"])
            return a_8760(orari["P"].to_numpy(dtype=float) / 1000.0)  # W -> kW
        except (requests.RequestException, KeyError, ValueError) as e:
            ultimo_errore = str(e)
    raise ErroreProfilo(f"PVGIS non ha restituito dati: {ultimo_errore}")


# ---------------------------------------------------------------------------------------------
# EOLICO
# ---------------------------------------------------------------------------------------------
def scarica_vento(lat, lon, anno=2019):
    """Velocità del vento oraria a 100 m (m/s) dall'archivio Open-Meteo (ERA5)."""
    params = {"latitude": lat, "longitude": lon, "start_date": f"{int(anno)}-01-01",
              "end_date": f"{int(anno)}-12-31", "hourly": "wind_speed_100m", "wind_speed_unit": "ms",
              "timezone": "UTC"}
    try:
        r = requests.get(URL_OPENMETEO, params=params, timeout=TIMEOUT)
        r.raise_for_status()
        return a_8760(r.json()["hourly"]["wind_speed_100m"])
    except (requests.RequestException, KeyError, ValueError) as e:
        raise ErroreProfilo(f"Open-Meteo non ha restituito dati sul vento: {e}")


def potenza_eolica(v_vento, p_nominale_kw, v_cut_in=3.0, v_rated=12.0, v_cut_out=25.0):
    """Curva di potenza semplificata (cubica tra cut-in e nominale), stessa dell'app precedente, vettorizzata."""
    v = np.asarray(v_vento, dtype=float)
    p = np.zeros_like(v)
    if p_nominale_kw <= 0:
        return p
    salita = (v >= v_cut_in) & (v < v_rated)
    p[salita] = p_nominale_kw * (v[salita] ** 3 - v_cut_in ** 3) / (v_rated ** 3 - v_cut_in ** 3)
    p[(v >= v_rated) & (v <= v_cut_out)] = p_nominale_kw
    return p


# ---------------------------------------------------------------------------------------------
# FILE CARICATI
# ---------------------------------------------------------------------------------------------
def leggi_csv_pvgis(contenuto):
    """CSV orario scaricato dal sito PVGIS (colonna P in W), come Analisi_tecnica.load_data."""
    testo = contenuto.decode("utf-8", errors="ignore") if isinstance(contenuto, bytes) else contenuto
    lines = testo.splitlines()
    header = "time,P,G(i),H_sun,T2m,WS10m,Int"
    idx = next((i for i, line in enumerate(lines) if line.strip() == header), None)
    if idx is None:
        raise ErroreProfilo("Non trovo l'intestazione PVGIS 'time,P,G(i),H_sun,T2m,WS10m,Int': "
                            "scarica di nuovo il file orario (Hourly data) da PVGIS in formato CSV.")
    righe = []
    for line in lines[idx + 1:]:
        col = line.strip().split(",")
        if len(col) != 7:
            break
        righe.append(col)
    df = pd.DataFrame(righe, columns=header.split(","))
    return pd.to_numeric(df["P"], errors="coerce").fillna(0).to_numpy() / 1000.0


def leggi_csv_generico(contenuto):
    """CSV con i valori orari in kW nella prima colonna; separatore e decimale riconosciuti in automatico."""
    testo = contenuto.decode("utf-8", errors="ignore") if isinstance(contenuto, bytes) else contenuto
    for sep, dec in ((";", ","), (",", "."), ("\t", ","), ("\t", "."), (";", ".")):
        try:
            df = pd.read_csv(io.StringIO(testo), sep=sep, decimal=dec, header=None)
        except Exception:
            continue
        testo_col = df.iloc[:, 0].astype(str).str.strip().str.replace(" ", "")
        if dec == ",":  # con l'intestazione la colonna resta testo: la virgola decimale va convertita a mano
            testo_col = testo_col.str.replace(",", ".", regex=False)
        col = pd.to_numeric(testo_col, errors="coerce")
        valori = col.dropna().to_numpy()
        if len(valori) >= ORE_ANNO:
            return valori
    raise ErroreProfilo("Non riesco a leggere il CSV: serve una colonna con 8760 valori orari in kW "
                        "(una riga di intestazione è ammessa).")


def leggi_profilo(contenuto, da_pvgis):
    valori = leggi_csv_pvgis(contenuto) if da_pvgis else leggi_csv_generico(contenuto)
    return a_8760(valori)


def resa_specifica(profilo_kw, potenza_kw):
    """kWh/kW annui: serve a controllare che potenza dichiarata e profilo siano coerenti."""
    return float(np.sum(profilo_kw)) / potenza_kw if potenza_kw > 0 else 0.0
