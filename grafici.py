"""Grafici Plotly per la dashboard H2FAST.

Colori per ruolo (palette categoriale validata per daltonismo, in ordine fisso):
    elettrolizzatore = blu, immessa/non utilizzata = arancio, batteria = acqua,
    fotovoltaico = giallo, eolico = blu, fonte extra = acqua; produzione totale = grigio scuro.
I grafici non impostano un template: Streamlit applica il proprio tema (chiaro/scuro).
"""
import numpy as np
import pandas as pd
import plotly.graph_objects as go

BLU = "#2a78d6"
ARANCIO = "#eb6834"
ACQUA = "#1baf7a"
GIALLO = "#eda100"
VIOLA = "#4a3aa7"
GRIGIO = "#7a7974"
SCALA_SEQ = "Blues"

ORE = pd.date_range("2019-01-01", periods=8760, freq="h")   # anno tipo non bisestile
MESI = ["Gen", "Feb", "Mar", "Apr", "Mag", "Giu", "Lug", "Ago", "Set", "Ott", "Nov", "Dic"]


def _stile(fig, titolo=None, altezza=380, unita_y=None, legenda=True):
    # titolo in alto, legenda sotto il grafico: così non si sovrappongono mai
    fig.update_layout(
        title=dict(text=titolo, x=0, xanchor="left", font=dict(size=15)) if titolo else None,
        height=altezza + (40 if legenda else 0), margin=dict(l=10, r=10, t=45 if titolo else 15, b=10),
        hovermode="x unified", showlegend=legenda, bargap=0.25,
        legend=dict(orientation="h", yanchor="top", y=-0.12, xanchor="left", x=0),
    )
    fig.update_xaxes(showgrid=False)
    fig.update_yaxes(gridcolor="rgba(128,128,128,0.18)", zeroline=False, title_text=unita_y)
    return fig


def mensile(serie):
    """Somma mensile (kWh -> MWh) di una serie oraria di 8760 valori."""
    return pd.Series(np.asarray(serie), index=ORE).groupby(ORE.month).sum().to_numpy() / 1000.0


# ---------------------------------------------------------------------------------------------
# PROFILI E FLUSSI ENERGETICI
# ---------------------------------------------------------------------------------------------
def fig_profilo_fonti(componenti):
    """Produzione mensile per fonte (barre impilate). componenti: {nome: array 8760 kW}."""
    colori = {"Fotovoltaico": GIALLO, "Eolico": BLU, "Fonte extra": ACQUA}
    fig = go.Figure()
    for nome, serie in componenti.items():
        if serie is None or np.sum(serie) <= 0:
            continue
        fig.add_bar(x=MESI, y=mensile(serie), name=nome, marker_color=colori.get(nome, VIOLA),
                    hovertemplate="%{y:,.1f} MWh")
    fig.update_layout(barmode="stack")
    return _stile(fig, "Produzione rinnovabile mensile", 320, "MWh")


def fig_energia_mensile(andamenti, batteria):
    """Destinazione dell'energia mese per mese per il progetto selezionato."""
    fig = go.Figure()
    fig.add_bar(x=MESI, y=mensile(andamenti[0]), name="All'elettrolizzatore", marker_color=BLU,
                hovertemplate="%{y:,.1f} MWh")
    fig.add_bar(x=MESI, y=mensile(andamenti[2]), name="Immessa in rete / non utilizzata", marker_color=ARANCIO,
                hovertemplate="%{y:,.1f} MWh")
    fig.add_scatter(x=MESI, y=mensile(andamenti[3]), name="Produzione rinnovabile", mode="lines+markers",
                    line=dict(color=GRIGIO, width=2), marker=dict(size=8), hovertemplate="%{y:,.1f} MWh")
    fig.update_layout(barmode="stack")
    return _stile(fig, "Dove va l'energia, mese per mese", 380, "MWh")


def fig_idrogeno_mensile(andamenti):
    fig = go.Figure()
    fig.add_bar(x=MESI, y=mensile(andamenti[1]) * 1000, name="Idrogeno", marker_color=BLU,
                hovertemplate="%{y:,.0f} kg")
    return _stile(fig, "Idrogeno prodotto per mese", 300, "kg", legenda=False)


def fig_heatmap(serie, titolo, unita, scala=SCALA_SEQ):
    """Mappa giorno x ora: si leggono insieme stagionalità e ciclo giornaliero."""
    z = np.asarray(serie).reshape(365, 24).T
    giorni = pd.date_range("2019-01-01", periods=365, freq="D")
    fig = go.Figure(go.Heatmap(z=z, x=giorni, y=list(range(24)), colorscale=scala,
                               colorbar=dict(title=unita, thickness=12),
                               hovertemplate="%{x|%d %b}, ore %{y}:00<br>%{z:,.1f} " + unita + "<extra></extra>"))
    fig.update_yaxes(title_text="Ora del giorno", dtick=6, autorange="reversed")
    fig.update_xaxes(tickformat="%b", dtick="M1")
    fig.update_layout(height=320, margin=dict(l=10, r=10, t=50, b=10),
                      title=dict(text=titolo, x=0, font=dict(size=15)))
    return fig


def fig_periodo(andamenti, batteria, inizio, giorni):
    """Andamento orario in un periodo scelto (default una settimana)."""
    a = inizio * 24
    b = min(a + giorni * 24, 8760)
    x = ORE[a:b]
    fig = go.Figure()
    fig.add_scatter(x=x, y=andamenti[3][a:b], name="Produzione rinnovabile", fill="tozeroy",
                    line=dict(color=GRIGIO, width=1), fillcolor="rgba(122,121,116,0.15)",
                    hovertemplate="%{y:,.0f} kW")
    fig.add_scatter(x=x, y=andamenti[0][a:b], name="All'elettrolizzatore", line=dict(color=BLU, width=2),
                    hovertemplate="%{y:,.0f} kW")
    fig.add_scatter(x=x, y=andamenti[2][a:b], name="Immessa / non utilizzata", line=dict(color=ARANCIO, width=2),
                    hovertemplate="%{y:,.0f} kW")
    fig.add_scatter(x=x, y=andamenti[5][a:b], name="Taglia elettrolizzatore",
                    line=dict(color=BLU, width=1, dash="dash"), hoverinfo="skip")
    fig.add_scatter(x=x, y=andamenti[4][a:b], name="Minimo tecnico",
                    line=dict(color=BLU, width=1, dash="dot"), hoverinfo="skip")
    if batteria:
        fig.add_scatter(x=x, y=andamenti[6][a:b], name="Energia in batteria", line=dict(color=ACQUA, width=2),
                        hovertemplate="%{y:,.0f} kWh")
    fig = _stile(fig, "Funzionamento ora per ora", 420, "kW / kWh")
    fig.update_xaxes(rangeslider=dict(visible=True, thickness=0.06))
    fig.update_layout(legend_y=-0.32)
    return fig


def fig_durata(andamenti):
    """Curva di durata del carico dell'elettrolizzatore (ore ordinate per potenza decrescente)."""
    p = np.sort(andamenti[0])[::-1]
    nom = andamenti[5][0]
    fig = go.Figure()
    fig.add_scatter(x=np.arange(1, 8761), y=p / nom * 100 if nom > 0 else p, line=dict(color=BLU, width=2),
                    fill="tozeroy", fillcolor="rgba(42,120,214,0.15)", name="Carico",
                    hovertemplate="%{x} ore sopra il %{y:.0f}%<extra></extra>")
    fig = _stile(fig, "Curva di durata: per quante ore l'elettrolizzatore lavora a un certo carico", 320,
                 "% della potenza nominale", legenda=False)
    fig.update_xaxes(title_text="ore/anno")
    fig.update_layout(hovermode="closest")
    return fig


# ---------------------------------------------------------------------------------------------
# ECONOMIA
# ---------------------------------------------------------------------------------------------
def voci_capex(a):
    """Scomposizione dell'investimento (netto IVA) del progetto."""
    voci = {
        "Fotovoltaico": a.ImpPV1 * a.ImpPV1eurokW,
        "Eolico": a.ImpWind * a.ImpWindeurokW,
        "Fonte extra": a.ImpExtra * a.ImpExtraeurokW,
        "Elettrolizzatore": a.PotEle * a.EletteuroKW,
        "Compressore": a.potenza_compressore * a.CompreuroKW,
        "Batteria": a.AccuE * a.AccuEeurokW,
        "Stoccaggio H2": a.impianto_stocc,
        "Bombole stoccaggio": a.BombSto,
        "Stazione di rifornimento": a.StazzRif,
        "Spese tecniche": a.SpeTOpere,
        "Lavori impiantistici": a.LavoImp,
        "Carrello elevatore": a.CarrEll,
        "Opere edili": a.OpeE,
        "Terreno": a.Terr,
    }
    return {k: v for k, v in voci.items() if v > 0}


def fig_capex(a):
    voci = voci_capex(a)
    s = pd.Series(voci).sort_values()
    tot = s.sum()
    fig = go.Figure(go.Bar(x=s.values, y=s.index, orientation="h", marker_color=BLU,
                           text=[f"{v / tot * 100:.0f}%" for v in s.values], textposition="outside",
                           hovertemplate="%{y}: € %{x:,.0f}<extra></extra>"))
    fig = _stile(fig, f"Investimento: € {tot:,.0f} (netto IVA)".replace(",", "."), max(260, 34 * len(s) + 80),
                 None, legenda=False)
    fig.update_xaxes(title_text="€", showgrid=True, gridcolor="rgba(128,128,128,0.18)")
    fig.update_layout(hovermode="closest")
    return fig


def fig_flussi_cassa(a):
    """Flusso di cassa netto annuo e cumulato, con l'anno di rientro."""
    anni = np.arange(len(a.FlussoNettoCassa))
    cum = np.cumsum(a.FlussoNettoCassa)
    fig = go.Figure()
    fig.add_bar(x=anni, y=a.FlussoNettoCassa, name="Flusso di cassa netto dell'anno",
                marker_color=[ARANCIO if v < 0 else BLU for v in a.FlussoNettoCassa], opacity=0.55,
                hovertemplate="€ %{y:,.0f}")
    fig.add_scatter(x=anni, y=cum, name="Cumulato", mode="lines+markers", line=dict(color=BLU, width=2),
                    marker=dict(size=8), hovertemplate="€ %{y:,.0f}")
    fig.add_hline(y=0, line=dict(color=GRIGIO, width=1))
    if a.PAYBACK is not None:
        fig.add_vline(x=a.PAYBACK, line=dict(color=GRIGIO, width=1, dash="dash"))
        fig.add_annotation(x=a.PAYBACK, y=cum[a.PAYBACK], text=f"rientro: anno {a.PAYBACK}", showarrow=True,
                           arrowhead=0, ax=40, ay=-30)
    fig = _stile(fig, "Flussi di cassa del progetto", 380, "€")
    fig.update_xaxes(title_text="anno", dtick=1)
    fig.update_layout(legend_y=-0.22)
    return fig


def fig_costo_kg(a):
    """Composizione del costo per kg (indicatori del motore originale) e LCOH attualizzato."""
    voci = pd.Series({
        "Costo medio operativo": a.costo_medio_operativo,
        "Costo medio investimenti": a.costo_medio_investimenti,
    })
    fig = go.Figure()
    fig.add_bar(x=["Full cost medio"], y=[voci.iloc[0]], name=voci.index[0], marker_color=ARANCIO,
                hovertemplate="%{y:.2f} €/kg")
    fig.add_bar(x=["Full cost medio"], y=[voci.iloc[1]], name=voci.index[1], marker_color=BLU,
                hovertemplate="%{y:.2f} €/kg")
    fig.add_bar(x=["LCOH attualizzato"], y=[a.LCOH], name="LCOH attualizzato", marker_color=VIOLA,
                hovertemplate="%{y:.2f} €/kg")
    fig.add_hline(y=a.prezzoindrogeno, line=dict(color=GRIGIO, dash="dash"),
                  annotation_text=f"prezzo di vendita {a.prezzoindrogeno:.2f} €/kg", annotation_position="top right")
    fig.update_layout(barmode="stack")
    return _stile(fig, "Costo dell'idrogeno", 340, "€/kg")


# ---------------------------------------------------------------------------------------------
# ESPLORAZIONE DELLE CONFIGURAZIONI
# ---------------------------------------------------------------------------------------------
def fronte_pareto(df, x, y, max_x, max_y):
    """ID delle configurazioni non dominate per gli obiettivi (x, y) con il verso indicato."""
    d = df[["ID", x, y]].replace([np.inf, -np.inf], np.nan).dropna()
    if d.empty:
        return d
    sx = d[x] if max_x else -d[x]
    sy = d[y] if max_y else -d[y]
    ordine = np.lexsort((-sy.to_numpy(), -sx.to_numpy()))
    tenuti, migliore = [], -np.inf
    for i in ordine:
        if sy.iloc[i] > migliore:
            tenuti.append(i)
            migliore = sy.iloc[i]
    return d.iloc[tenuti].sort_values(x)


def fig_esplora(df, x, y, colore, nome_x, nome_y, nome_col, pareto=None, top_ids=None, sel_id=None, ok_mask=None):
    fig = go.Figure()
    if ok_mask is not None and not ok_mask.all():
        escluse = df[~ok_mask]
        fig.add_scatter(x=escluse[x], y=escluse[y], mode="markers", name="Fuori dai vincoli",
                        customdata=escluse["ID"], marker=dict(size=6, color="rgba(128,128,128,0.35)"),
                        hovertemplate=f"{nome_x}: %{{x:,.2f}}<br>{nome_y}: %{{y:,.2f}}"
                                      f"<extra>config %{{customdata}} · fuori dai vincoli</extra>")
        df = df[ok_mask]
    fig.add_scatter(x=df[x], y=df[y], mode="markers", name="Configurazioni",
                      customdata=df["ID"],
                      marker=dict(size=8, color=df[colore], colorscale=SCALA_SEQ, showscale=True,
                                  colorbar=dict(title=nome_col, thickness=12), opacity=0.8,
                                  line=dict(width=0)),
                      hovertemplate=f"{nome_x}: %{{x:,.2f}}<br>{nome_y}: %{{y:,.2f}}<br>"
                                    f"{nome_col}: %{{marker.color:,.2f}}<extra>config %{{customdata}}</extra>")
    if pareto is not None and len(pareto) > 1:
        fig.add_scatter(x=pareto[x], y=pareto[y], mode="lines", name="Frontiera di Pareto",
                        line=dict(color=ARANCIO, width=2), hoverinfo="skip")
    if top_ids:
        t = df[df["ID"].isin(top_ids)]
        fig.add_scatter(x=t[x], y=t[y], mode="markers", name="Migliori progetti",
                        marker=dict(size=13, color="rgba(0,0,0,0)", line=dict(color=ARANCIO, width=2)),
                        hoverinfo="skip")
    if sel_id is not None:
        s = df[df["ID"] == sel_id]
        fig.add_scatter(x=s[x], y=s[y], mode="markers", name="Selezionata",
                        marker=dict(size=16, symbol="x", color=VIOLA), hoverinfo="skip")
    fig = _stile(fig, None, 520, nome_y)
    fig.update_xaxes(title_text=nome_x, showgrid=True, gridcolor="rgba(128,128,128,0.18)")
    fig.update_layout(hovermode="closest", dragmode="zoom", clickmode="event+select", legend_y=-0.16)
    return fig


def fig_sensitivity(df, variabile, nome_var, max_linee=8):
    """Variabile in funzione della taglia batteria, una linea per taglia di elettrolizzatore."""
    d = df.dropna(subset=[variabile])
    taglie = np.sort(d["PotEle"].unique())
    if len(taglie) > max_linee:
        scelte = taglie[np.linspace(0, len(taglie) - 1, max_linee).round().astype(int)]
    else:
        scelte = taglie
    fig = go.Figure()
    colori = np.linspace(0.35, 1.0, len(scelte))
    for t, c in zip(scelte, colori):
        g = d[d["PotEle"] == t].sort_values("AccuE")
        blu = f"rgba(42,120,214,{c:.2f})"
        fig.add_scatter(x=g["AccuE"], y=g[variabile], mode="lines", name=f"{t:,.0f} kW",
                        line=dict(color=blu, width=2),
                        hovertemplate=f"elettrolizzatore {t:,.0f} kW<br>batteria %{{x:,.0f}} kWh<br>"
                                      f"{nome_var}: %{{y:,.2f}}<extra></extra>")
    fig = _stile(fig, f"{nome_var} al variare della batteria, per taglia di elettrolizzatore", 380, nome_var)
    fig.update_xaxes(title_text="Batteria [kWh]")
    fig.update_layout(hovermode="closest", legend=dict(title="Elettrolizzatore", orientation="v", x=1.02, y=1,
                                                       xanchor="left", yanchor="top"))
    return fig
