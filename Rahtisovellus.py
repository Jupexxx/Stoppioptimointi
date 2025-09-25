import streamlit as st
import pandas as pd
import numpy as np
import pyomo.environ as pyo
from io import BytesIO
import time
import pydeck as pdk
from pyproj import Transformer
from pyomo.common.errors import ApplicationError # Tuodaan virheluokka

# =============================================================================
# VAKIOT JA KONFIGURAATIO
# =============================================================================
# Excel-välilehtien nimet
SHEET_AUTOT = 'AutojenYhteenveto'
SHEET_JAKO = 'Jakokeikat'
SHEET_NOUTO = 'Noutokeikat'
SHEET_PNRO = 'Postinumerot'
SHEET_TARIFFI = 'Tariffitaulukko'

# Yleiset sarakkeiden nimet
COL_AUTOTUNNUS = 'Autotunnus'
COL_LIIKENNOITSIJA = 'Liikennöitsijän nimi'
COL_VANHAT_KULUT = 'Kulut entisellä mallilla'
COL_RAHTIKIRJA = 'Rahtikirjanumero'
COL_POSTINUMERO = 'Postinumero'
COL_POSTITOIMIPAIKKA = 'Postitoimipaikka'
COL_KILOT = 'Kilot'
COL_NIPPUNUMERO = 'Nippunumero'
COL_VYOHYKE = 'Vyöhyke'
COL_X_KOORD = 'X-Koordinaatti'
COL_Y_KOORD = 'Y-Koordinaatti'
COL_PAINO_ALKU = 'Painoluokka ALKAA (kg)'
COL_PAINO_LOPPU = 'Painoluokka LOPPUU (kg)'
COL_LASKENTATAPA = 'Laskentatapa'
PREFIX_VYOHYKE_COL = 'VYÖHYKE'
LASKENTATAPA_KG = '€/kg'


# =============================================================================
# APUFUNKTIOT JA DATAN KÄSITTELY
# =============================================================================
def luo_mallipohja_exceliin():
    """Luo ja palauttaa Excel-mallipohjan BytesIO-muodossa."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        pd.DataFrame({
            COL_AUTOTUNNUS: ['AUTO-001', 'AUTO-002'],
            COL_LIIKENNOITSIJA: ['Kuljetus A Oy', 'Siirto B Tmi'],
            COL_VANHAT_KULUT: [15500.50, 22340.00]
        }).to_excel(writer, sheet_name=SHEET_AUTOT, index=False)
        pd.DataFrame({
            COL_AUTOTUNNUS: ['AUTO-001', 'AUTO-001', 'AUTO-002'], COL_RAHTIKIRJA: ['RK123', 'RK124', 'RK125'],
            'Vastaanottaja': ['Yritys X', 'Yritys Y', 'Yritys Z'], COL_POSTINUMERO: ['60100', '61800', '61330'],
            COL_KILOT: [850, 120, 25000], COL_NIPPUNUMERO: ['NIPPU-01', 'NIPPU-01', 'NIPPU-02']
        }).to_excel(writer, sheet_name=SHEET_JAKO, index=False)
        pd.DataFrame({
            COL_AUTOTUNNUS: ['AUTO-001', 'AUTO-002'], COL_RAHTIKIRJA: ['RK501', 'RK502'],
            'Noutopaikka': ['Varasto A', 'Tehdas B'], COL_POSTINUMERO: ['60220', '60800'],
            COL_KILOT: [1500, 3450], COL_NIPPUNUMERO: ['NIPPU-03', 'NIPPU-04']
        }).to_excel(writer, sheet_name=SHEET_NOUTO, index=False)
        pd.DataFrame({
            COL_POSTINUMERO: ['60100', '60101', '60220', '61800', '61330', '60800', '60801'],
            COL_POSTITOIMIPAIKKA: ['Seinäjoki Keskus', 'Seinäjoki PL', 'Pajuluoma', 'Kauhajoki', 'Jalasjärvi kk', 'Ilmajoki', 'Ilmajoki PL'],
            COL_VYOHYKE: [1, 1, 1, 3, 5, 2, 2],
            COL_X_KOORD: [288812, 'EILÖYDY', 286625, 274433, 271234, 273247, 'EILÖYDY'],
            COL_Y_KOORD: [6967917, 'EILÖYDY', 6967422, 6940011, 6932111, 6964498, 'EILÖYDY']
        }).to_excel(writer, sheet_name=SHEET_PNRO, index=False)
        pd.DataFrame({
            COL_PAINO_ALKU: [0, 501, 10001, 20001], COL_PAINO_LOPPU: [500, 10000, 20000, np.nan],
            COL_LASKENTATAPA: ['€/nippu', '€/nippu', '€/nippu', LASKENTATAPA_KG],
            f'{PREFIX_VYOHYKE_COL} 1': [50.0, 120.0, 200.0, 0.008], f'{PREFIX_VYOHYKE_COL} 2': [65.0, 150.0, 220.0, 0.010],
            f'{PREFIX_VYOHYKE_COL} 3': [80.0, 180.0, 250.0, 0.012], f'{PREFIX_VYOHYKE_COL} 4': [90.0, 200.0, 280.0, 0.014],
            f'{PREFIX_VYOHYKE_COL} 5': [100.0, 220.0, 310.0, 0.016], f'{PREFIX_VYOHYKE_COL} 6': [120.0, 250.0, 350.0, 0.018],
        }).to_excel(writer, sheet_name=SHEET_TARIFFI, index=False)
    return output.getvalue()

def validoi_syotetiedosto(sheets):
    """Tarkistaa, että ladattu Excel sisältää kaikki tarvittavat välilehdet ja sarakkeet."""
    vaatimukset = {
        SHEET_AUTOT: [COL_AUTOTUNNUS, COL_LIIKENNOITSIJA, COL_VANHAT_KULUT],
        SHEET_JAKO: [COL_AUTOTUNNUS, COL_RAHTIKIRJA, COL_POSTINUMERO, COL_KILOT, COL_NIPPUNUMERO],
        SHEET_NOUTO: [COL_AUTOTUNNUS, COL_RAHTIKIRJA, COL_POSTINUMERO, COL_KILOT, COL_NIPPUNUMERO],
        SHEET_PNRO: [COL_POSTINUMERO, COL_VYOHYKE, COL_X_KOORD, COL_Y_KOORD],
        SHEET_TARIFFI: [COL_PAINO_ALKU, COL_PAINO_LOPPU, COL_LASKENTATAPA]
    }
    virheet = []
    for sheet_name, required_cols in vaatimukset.items():
        if sheet_name not in sheets:
            virheet.append(f"Puuttuva välilehti: '{sheet_name}'")
            continue
        df_cols = sheets[sheet_name].columns
        for col in required_cols:
            if col not in df_cols:
                virheet.append(f"Puuttuva sarake välilehdellä '{sheet_name}': '{col}'")
    
    if f'{PREFIX_VYOHYKE_COL} 1' not in sheets[SHEET_TARIFFI].columns:
         virheet.append(f"Välilehdeltä '{SHEET_TARIFFI}' puuttuu vähintään yksi vyöhykesarake (esim. '{PREFIX_VYOHYKE_COL} 1')")
         
    return virheet if virheet else None

def get_painoluokka_str(idx, df_tariff):
    rivi = df_tariff.loc[idx]
    alku, loppu = rivi[COL_PAINO_ALKU], rivi[COL_PAINO_LOPPU]
    if pd.isna(loppu): return f"> {int(alku) - 1} kg"
    return f"{int(alku)}-{int(loppu)} kg"

def get_painoluokka_rivi_idx(paino, df_tariff):
    df_tariff_copy = df_tariff.copy()
    df_tariff_copy[COL_PAINO_LOPPU] = df_tariff_copy[COL_PAINO_LOPPU].fillna(np.inf)
    sopivat_rivit = df_tariff_copy[(df_tariff_copy[COL_PAINO_ALKU] <= paino) & (df_tariff_copy[COL_PAINO_LOPPU] >= paino)]
    return sopivat_rivit.index[0] if not sopivat_rivit.empty else None

def _valmistele_data(sheets, autot_mukana):
    df_autot_orig = sheets[SHEET_AUTOT]
    df_autot = df_autot_orig[df_autot_orig[COL_AUTOTUNNUS].isin(autot_mukana)].copy()
    df_keikat = pd.concat([sheets[SHEET_JAKO], sheets[SHEET_NOUTO]], ignore_index=True)
    df_keikat = df_keikat[df_keikat[COL_AUTOTUNNUS].isin(autot_mukana)]
    df_niput = df_keikat.groupby(COL_NIPPUNUMERO).agg(
        nippu_paino=(COL_KILOT, 'sum'),
        Autotunnus=(COL_AUTOTUNNUS, 'first'),
        Postinumero=(COL_POSTINUMERO, 'first')
    ).reset_index()
    df_niput[COL_POSTINUMERO] = df_niput[COL_POSTINUMERO].astype(str)
    return df_autot, df_niput, sheets[SHEET_TARIFFI]

# --- KORJATTU JA VANKENNETTU FUNKTIO ---
def laske_oikea_nippu_hinta(row, df_tariff_current, df_tariff_orig):
    """Laskee nipulle hinnan ja varmistaa monotonisuuden painoluokkien välillä."""
    paino = row['nippu_paino']
    vyohyke = row['Vyöhyke']
    vyohyke_sarake = f"{PREFIX_VYOHYKE_COL} {vyohyke}"
    if vyohyke_sarake not in df_tariff_current.columns:
        return 0.0

    rivi_idx = get_painoluokka_rivi_idx(paino, df_tariff_orig)
    if rivi_idx is None:
        return 0.0
    
    try:
        rivi_sijainti = df_tariff_orig.index.get_loc(rivi_idx)
    except KeyError:
        return 0.0

    nykyinen_rivi = df_tariff_orig.iloc[rivi_sijainti]
    laskentatapa = nykyinen_rivi[COL_LASKENTATAPA]
    hinta = pd.to_numeric(df_tariff_current.iloc[rivi_sijainti][vyohyke_sarake], errors='coerce')
    if pd.isna(hinta): return 0.0
    
    raakahinta = paino * hinta if laskentatapa == LASKENTATAPA_KG else hinta

    if rivi_sijainti == 0:
        return raakahinta

    # Vertailu tehdään vain, jos nykyinen laskentatapa on €/kg, koska silloin hinta voi "pudota"
    if laskentatapa == LASKENTATAPA_KG:
        edellinen_rivi = df_tariff_orig.iloc[rivi_sijainti - 1]
        edellinen_loppu_kg = pd.to_numeric(edellinen_rivi[COL_PAINO_LOPPU], errors='coerce')
        
        if pd.notna(edellinen_loppu_kg):
            edellinen_laskentatapa = edellinen_rivi[COL_LASKENTATAPA]
            edellinen_hinta = pd.to_numeric(df_tariff_current.iloc[rivi_sijainti - 1][vyohyke_sarake], errors='coerce')
            if pd.isna(edellinen_hinta): return raakahinta

            lattiahinta = edellinen_loppu_kg * edellinen_hinta if edellinen_laskentatapa == LASKENTATAPA_KG else edellinen_hinta
            return max(raakahinta, lattiahinta)
    
    return raakahinta
# --- KORJAUKSEN LOPPU ---

# =============================================================================
# OPTIMOINTI: TARIFFIEN LASKENTA
# =============================================================================
# ... (tämä osa pysyy täysin samana) ...
def _luo_tariffimallin_perusrakenne(df_autot, df_tariff_input):
    model = pyo.ConcreteModel(name="TariffiOptimointi")
    model.AUTOT = pyo.Set(initialize=list(df_autot[COL_AUTOTUNNUS]))
    vyohyke_sarakkeet = [c for c in df_tariff_input.columns if PREFIX_VYOHYKE_COL in c]
    model.VYOHYKKEET = pyo.Set(initialize=sorted([int(c.split(' ')[1]) for c in vyohyke_sarakkeet]))
    model.TARIFFI_RIVIT = pyo.Set(initialize=list(df_tariff_input.index))
    model.tariffi = pyo.Var(model.TARIFFI_RIVIT, model.VYOHYKKEET, within=pyo.NonNegativeReals)
    return model

def _lisaa_tariffimallin_kustannuslausekkeet_ja_tavoite(model, df_niput, df_tariff_input, vanhat_kulut_dict):
    @model.Expression(model.AUTOT)
    def uusi_kustannus_per_auto(m, auto):
        auton_niput = df_niput[df_niput[COL_AUTOTUNNUS] == auto]
        total_cost = 0
        for _, nippu in auton_niput.iterrows():
            rivi_idx, vyohyke = int(nippu['tariffi_rivi_idx']), int(nippu[COL_VYOHYKE])
            if vyohyke not in m.VYOHYKKEET: continue
            laskentatapa = df_tariff_input.at[rivi_idx, COL_LASKENTATAPA]
            cost = m.tariffi[rivi_idx, vyohyke]
            total_cost += nippu['nippu_paino'] * cost if laskentatapa == LASKENTATAPA_KG else cost
        return total_cost

    model.pos_ero = pyo.Var(model.AUTOT, within=pyo.NonNegativeReals)
    model.neg_ero = pyo.Var(model.AUTOT, within=pyo.NonNegativeReals)
    model.tavoite = pyo.Objective(rule=lambda m: sum(m.pos_ero[a] + m.neg_ero[a] for a in m.AUTOT), sense=pyo.minimize)

    @model.Constraint(model.AUTOT)
    def erotus_rajoite(m, a):
        return m.uusi_kustannus_per_auto[a] - vanhat_kulut_dict.get(a, 0) == m.pos_ero[a] - m.neg_ero[a]
    return model

def _lisaa_tariffimallin_rajoitteet(model, df_autot, df_tariff_input, vanhat_kulut_dict, params):
    heitto = params['heitto'] / 100.0
    if params['taso'] == 'Auto':
        @model.Constraint(model.AUTOT)
        def kustannusrajoite(m, a):
            vanha = vanhat_kulut_dict.get(a, 0)
            return pyo.inequality(vanha * (1 - heitto), m.uusi_kustannus_per_auto[a], vanha * (1 + heitto))
    elif params['taso'] == 'Liikennöitsijä':
        liikennöitsijat = df_autot[COL_LIIKENNOITSIJA].unique()
        model.LIIKENNOITSIJAT = pyo.Set(initialize=liikennöitsijat)
        vanhat_kulut_liikennöitsijä = df_autot.groupby(COL_LIIKENNOITSIJA)[COL_VANHAT_KULUT].sum()
        @model.Expression(model.LIIKENNOITSIJAT)
        def uusi_kustannus_per_liikennöitsijä(m, l):
            return sum(m.uusi_kustannus_per_auto[a] for a in df_autot[df_autot[COL_LIIKENNOITSIJA] == l][COL_AUTOTUNNUS])
        @model.Constraint(model.LIIKENNOITSIJAT)
        def kustannusrajoite(m, l):
            vanha = vanhat_kulut_liikennöitsijä.get(l, 0)
            return pyo.inequality(vanha * (1 - heitto), m.uusi_kustannus_per_liikennöitsijä[l], vanha * (1 + heitto))
    else: # Kokonaisuus
        vanhat_yht = df_autot[COL_VANHAT_KULUT].sum()
        uudet_yht = sum(model.uusi_kustannus_per_auto[a] for a in model.AUTOT)
        model.kustannusrajoite = pyo.Constraint(rule=pyo.inequality(vanhat_yht * (1 - heitto), uudet_yht, vanhat_yht * (1 + heitto)))

    model.lukitut_rajoitteet = pyo.ConstraintList()
    for (r, v_str), arvo in params['lukitut_tariffit'].items():
        v = int(v_str.split(' ')[1])
        if v in model.VYOHYKKEET:
            model.lukitut_rajoitteet.add(model.tariffi[r, v] == arvo)

    MIN_KERROIN, MAX_KERROIN = 1.0 + (params['min_korotus'] / 100.0), 1.0 + (params['max_korotus'] / 100.0)
    
    @model.Constraint(model.TARIFFI_RIVIT, model.VYOHYKKEET)
    def monotonisuus_paino_MIN(m, r, v):
        if r + 1 in m.TARIFFI_RIVIT and df_tariff_input.at[r, COL_LASKENTATAPA] == df_tariff_input.at[r+1, COL_LASKENTATAPA]:
            return m.tariffi[r+1, v] >= m.tariffi[r, v] * MIN_KERROIN
        return pyo.Constraint.Skip
    @model.Constraint(model.TARIFFI_RIVIT, model.VYOHYKKEET)
    def monotonisuus_paino_MAX(m, r, v):
        if r + 1 in m.TARIFFI_RIVIT and df_tariff_input.at[r, COL_LASKENTATAPA] == df_tariff_input.at[r+1, COL_LASKENTATAPA]:
            return m.tariffi[r+1, v] <= m.tariffi[r, v] * MAX_KERROIN
        return pyo.Constraint.Skip

    @model.Constraint(model.TARIFFI_RIVIT, model.VYOHYKKEET)
    def monotonisuus_vyohyke_MIN(m, r, v):
        if v + 1 in m.VYOHYKKEET:
            return m.tariffi[r, v+1] >= m.tariffi[r, v] * MIN_KERROIN
        return pyo.Constraint.Skip
    @model.Constraint(model.TARIFFI_RIVIT, model.VYOHYKKEET)
    def monotonisuus_vyohyke_MAX(m, r, v):
        if v + 1 in m.VYOHYKKEET:
            return m.tariffi[r, v+1] <= m.tariffi[r, v] * MAX_KERROIN
        return pyo.Constraint.Skip

    return model

def suorita_tariffi_optimointi(sheets, df_zones_current, autot_mukana, params):
    df_autot, df_niput_base, df_tariff_input = _valmistele_data(sheets, autot_mukana)
    df_niput_base['tariffi_rivi_idx'] = df_niput_base['nippu_paino'].apply(lambda p: get_painoluokka_rivi_idx(p, df_tariff_input))
    df_niput_base.dropna(subset=['tariffi_rivi_idx'], inplace=True)
    df_zones_current[COL_POSTINUMERO] = df_zones_current[COL_POSTINUMERO].astype(str)
    df_niput = pd.merge(df_niput_base, df_zones_current, on=COL_POSTINUMERO, how='inner')
    df_niput[COL_VYOHYKE] = pd.to_numeric(df_niput[COL_VYOHYKE], errors='coerce').fillna(0).astype(int)
    if df_niput.empty: return "virhe", "Datan yhdistelyn jälkeen ei jäänyt yhtään käsiteltävää nippua.", None

    model = _luo_tariffimallin_perusrakenne(df_autot, df_tariff_input)
    vanhat_kulut_dict = df_autot.set_index(COL_AUTOTUNNUS)[COL_VANHAT_KULUT].to_dict()
    model = _lisaa_tariffimallin_kustannuslausekkeet_ja_tavoite(model, df_niput, df_tariff_input, vanhat_kulut_dict)
    model = _lisaa_tariffimallin_rajoitteet(model, df_autot, df_tariff_input, vanhat_kulut_dict, params)

    solver = pyo.SolverFactory('cbc')
    try:
        results = solver.solve(model, tee=False)
    except ApplicationError:
        error_msg = ("Ratkaisija kaatui. Tämä johtuu Streamlit Cloudissa yleensä resurssirajoituksista (RAM). "
                     "Kokeile pienemmällä datasetillä tai löysemmillä parametreilla.")
        return "virhe", error_msg, None

    if (results.solver.status == pyo.SolverStatus.ok) and (results.solver.termination_condition == pyo.TerminationCondition.optimal):
        df_tulos = df_tariff_input.copy()
        vyohyke_sarakkeet = [c for c in df_tariff_input.columns if PREFIX_VYOHYKE_COL in c]
        for r in model.TARIFFI_RIVIT:
            for v_idx, col in enumerate(vyohyke_sarakkeet, 1):
                if v_idx in model.VYOHYKKEET:
                    df_tulos.at[r, col] = round(pyo.value(model.tariffi[r, v_idx]), 4)
        df_vertailu = pd.DataFrame([{
            COL_AUTOTUNNUS: a, 'Vanha kustannus (€)': vanhat_kulut_dict.get(a, 0),
            'Uusi kustannus (€)': pyo.value(model.uusi_kustannus_per_auto[a])
        } for a in model.AUTOT])
        return "ok", df_tulos, df_vertailu
    else:
        return "virhe", "Ratkaisua ei löytynyt. Kokeile löysempiä parametreja tai poista lukituksia.", None

# =============================================================================
# OPTIMOINTI: VYÖHYKKEIDEN MÄÄRITYS
# =============================================================================
def suorita_vyohyke_optimointi(sheets, df_tariff_current, autot_mukana, params):
    df_autot, df_niput_base, df_tariff_input = _valmistele_data(sheets, autot_mukana)
    df_niput_base['tariffi_rivi_idx'] = df_niput_base['nippu_paino'].apply(lambda p: get_painoluokka_rivi_idx(p, df_tariff_input))
    df_niput = df_niput_base.dropna(subset=['tariffi_rivi_idx']).copy()
    if df_niput.empty: return "virhe", "Datan valmistelun jälkeen ei jäänyt yhtään käsiteltävää nippua.", None
    model = pyo.ConcreteModel()
    autot = list(df_autot[COL_AUTOTUNNUS]); model.AUTOT = pyo.Set(initialize=autot)
    postinumerot = df_niput[COL_POSTINUMERO].unique(); model.POSTINUMEROT = pyo.Set(initialize=postinumerot)
    vyohyke_sarakkeet = [c for c in df_tariff_current.columns if PREFIX_VYOHYKE_COL in c]
    vyohykkeet = sorted([int(c.split(' ')[1]) for c in vyohyke_sarakkeet]); model.VYOHYKKEET = pyo.Set(initialize=vyohykkeet)
    model.y = pyo.Var(model.POSTINUMEROT, model.VYOHYKKEET, within=pyo.Binary)
    @model.Constraint(model.POSTINUMEROT)
    def vain_yksi_vyohyke_per_pnro(m, p): return sum(m.y[p, v] for v in m.VYOHYKKEET) == 1
    model.lukitut_vyohykkeet = pyo.ConstraintList()
    for pnro, vyohyke in params['lukitut_vyohykkeet'].items():
        if pnro in model.POSTINUMEROT and pd.notna(vyohyke): 
            if int(vyohyke) in model.VYOHYKKEET:
                model.lukitut_vyohykkeet.add(model.y[pnro, int(vyohyke)] == 1)
    tariffi_dict = {(r_idx, v_idx): row[v_col] for r_idx, row in df_tariff_current.iterrows() for v_idx, v_col in zip(vyohykkeet, vyohyke_sarakkeet)}
    @model.Expression(model.AUTOT)
    def uusi_kustannus_per_auto(m, auto):
        auton_niput = df_niput[df_niput[COL_AUTOTUNNUS] == auto]; total_cost = 0
        for _, nippu in auton_niput.iterrows():
            pnro, rivi_idx = nippu[COL_POSTINUMERO], int(nippu['tariffi_rivi_idx']); laskentatapa = df_tariff_input.at[rivi_idx, COL_LASKENTATAPA]
            nippu_hinta = sum(m.y[pnro, v] * tariffi_dict[(rivi_idx, v)] for v in m.VYOHYKKEET)
            total_cost += nippu['nippu_paino'] * nippu_hinta if laskentatapa == LASKENTATAPA_KG else nippu_hinta
        return total_cost
    vanhat_kulut_dict_auto = df_autot.set_index(COL_AUTOTUNNUS)[COL_VANHAT_KULUT].to_dict()
    heitto_kerroin = params['heitto'] / 100.0
    model.pos_ero = pyo.Var(model.AUTOT, within=pyo.NonNegativeReals); model.neg_ero = pyo.Var(model.AUTOT, within=pyo.NonNegativeReals)
    model.tavoite = pyo.Objective(rule=lambda m: sum(m.pos_ero[a] + m.neg_ero[a] for a in m.AUTOT), sense=pyo.minimize)
    @model.Constraint(model.AUTOT)
    def erotus_rajoite(m, a): return m.uusi_kustannus_per_auto[a] - vanhat_kulut_dict_auto.get(a, 0) == m.pos_ero[a] - m.neg_ero[a]
    if params['taso'] == 'Auto':
        @model.Constraint(model.AUTOT)
        def kustannusrajoite(m, a): vanha = vanhat_kulut_dict_auto.get(a,0); return pyo.inequality(vanha * (1-heitto_kerroin), m.uusi_kustannus_per_auto[a], vanha * (1+heitto_kerroin))
    elif params['taso'] == 'Liikennöitsijä':
        liikennöitsijat = df_autot[COL_LIIKENNOITSIJA].unique(); model.LIIKENNOITSIJAT = pyo.Set(initialize=liikennöitsijat)
        vanhat_kulut_liikennöitsijä = df_autot.groupby(COL_LIIKENNOITSIJA)[COL_VANHAT_KULUT].sum()
        @model.Expression(model.LIIKENNOITSIJAT)
        def uusi_kustannus_per_liikennöitsijä(m, l): return sum(m.uusi_kustannus_per_auto[a] for a in df_autot[df_autot[COL_LIIKENNOITSIJA] == l][COL_AUTOTUNNUS])
        @model.Constraint(model.LIIKENNOITSIJAT)
        def kustannusrajoite(m, l): vanha = vanhat_kulut_liikennöitsijä.get(l, 0); return pyo.inequality(vanha * (1 - heitto_kerroin), m.uusi_kustannus_per_liikennöitsijä[l], vanha * (1 + heitto_kerroin))
    else:
        vanhat_kulut_yhteensä = df_autot[COL_VANHAT_KULUT].sum(); uudet_kulut_yhteensä = sum(model.uusi_kustannus_per_auto[a] for a in model.AUTOT)
        model.kustannusrajoite = pyo.Constraint(rule=pyo.inequality(vanhat_kulut_yhteensä * (1 - heitto_kerroin), uudet_kulut_yhteensä, vanhat_kulut_yhteensä * (1 + heitto_kerroin)))
    
    solver = pyo.SolverFactory('cbc')
    solver.options['threads'] = 1
    solver.options['ratio'] = params['vaje'] / 100.0
    try:
        results = solver.solve(model, tee=False)
    except ApplicationError:
        error_msg = ("Ratkaisija kaatui. Tämä johtuu Streamlit Cloudissa yleensä resurssirajoituksista (RAM), "
                     "erityisesti suurilla datamäärillä. Kokeile pienemmällä datasetillä tai suurenna sallittua optimointivajetta.")
        return "virhe", error_msg, None

    if (results.solver.status == pyo.SolverStatus.ok) and (results.solver.termination_condition == pyo.TerminationCondition.optimal):
        tulokset = [{COL_POSTINUMERO: p, COL_VYOHYKE: v} for p in model.POSTINUMEROT for v in model.VYOHYKKEET if pyo.value(model.y[p,v]) > 0.9]
        df_tulos = pd.DataFrame(tulokset)
        df_vertailu_auto = pd.DataFrame([{COL_AUTOTUNNUS: a, 'Vanha kustannus (€)': vanhat_kulut_dict_auto.get(a, 0), 'Uusi kustannus (€)': pyo.value(model.uusi_kustannus_per_auto[a])} for a in model.AUTOT])
        return "ok", df_tulos, df_vertailu_auto
    else: return "virhe", "Ratkaisua ei löytynyt. Kokeile löysempiä parametreja tai poista lukituksia.", None

def laske_vyohykkeet_automaattisesti(df_keikat, df_pnro, paakeskus_pnro='60100'):
    df_keikat[COL_POSTINUMERO] = df_keikat[COL_POSTINUMERO].astype(str)
    df_volyymit = df_keikat.groupby(COL_POSTINUMERO).agg(Kilot_sum=(COL_KILOT, 'sum'), Rahtikirjojen_lkm=(COL_RAHTIKIRJA, 'nunique'), Nippujen_lkm=(COL_NIPPUNUMERO, 'nunique')).reset_index()
    df_pnro[COL_POSTINUMERO] = df_pnro[COL_POSTINUMERO].astype(str)
    df_pnro_valmis = pd.merge(df_pnro, df_volyymit, on=COL_POSTINUMERO, how='left').fillna({'Rahtikirjojen_lkm': 0, 'Nippujen_lkm': 0, 'Kilot_sum': 0})
    df_pnro_valmis.replace('EILÖYDY', np.nan, inplace=True)
    coords_map = df_pnro_valmis.dropna(subset=[COL_X_KOORD, COL_Y_KOORD]).set_index(COL_POSTINUMERO)
    for idx, row in df_pnro_valmis[df_pnro_valmis[COL_X_KOORD].isna()].iterrows():
        try:
            base_pnro = str(int(row[COL_POSTINUMERO]) // 100 * 100)
            if base_pnro in coords_map.index:
                df_pnro_valmis.loc[idx, COL_X_KOORD] = coords_map.loc[base_pnro, COL_X_KOORD]; df_pnro_valmis.loc[idx, COL_Y_KOORD] = coords_map.loc[base_pnro, COL_Y_KOORD]
        except (ValueError, TypeError): continue
    df_pnro_valmis.dropna(subset=[COL_X_KOORD, COL_Y_KOORD], inplace=True)
    if df_pnro_valmis.empty: raise ValueError("Koordinaattidataa ei löytynyt tai sitä ei voitu käsitellä.")
    df_pnro_valmis[COL_X_KOORD] = pd.to_numeric(df_pnro_valmis[COL_X_KOORD]); df_pnro_valmis[COL_Y_KOORD] = pd.to_numeric(df_pnro_valmis[COL_Y_KOORD])
    if paakeskus_pnro not in df_pnro_valmis[COL_POSTINUMERO].values: raise ValueError(f"Pääkeskuksen postinumeroa {paakeskus_pnro} ei löytynyt datasta.")
    paakeskus_coords = df_pnro_valmis[df_pnro_valmis[COL_POSTINUMERO] == paakeskus_pnro][[COL_X_KOORD, COL_Y_KOORD]].iloc[0]
    hub_threshold = df_pnro_valmis['Rahtikirjojen_lkm'].quantile(0.95)
    df_pnro_valmis['Onko_Hub'] = (df_pnro_valmis['Rahtikirjojen_lkm'] >= hub_threshold) & (df_pnro_valmis['Rahtikirjojen_lkm'] >= 5)
    df_pnro_valmis['Etaisyys_paakeskuksesta_km'] = np.sqrt((df_pnro_valmis[COL_X_KOORD] - paakeskus_coords[COL_X_KOORD])**2 + (df_pnro_valmis[COL_Y_KOORD] - paakeskus_coords[COL_Y_KOORD])**2) / 1000
    df_pnro_valmis['Syrjaisyyspisteet'] = df_pnro_valmis['Etaisyys_paakeskuksesta_km'] / (df_pnro_valmis['Nippujen_lkm'] + 1)
    def maarita_vyohyke(row):
        if row[COL_POSTINUMERO] == paakeskus_pnro: return 1
        if row['Etaisyys_paakeskuksesta_km'] < 15: return 2
        if 15 <= row['Etaisyys_paakeskuksesta_km'] < 40 and row['Onko_Hub']: return 3
        if row['Etaisyys_paakeskuksesta_km'] >= 40 and row['Onko_Hub']: return 4
        maaseutu = df_pnro_valmis[(df_pnro_valmis[COL_POSTINUMERO] != paakeskus_pnro) & (df_pnro_valmis['Etaisyys_paakeskuksesta_km'] >= 15) & (~df_pnro_valmis['Onko_Hub'])]
        if maaseutu.empty: return 5
        median_syrjaisyys = maaseutu['Syrjaisyyspisteet'].median()
        if row['Syrjaisyyspisteet'] <= median_syrjaisyys: return 5
        else: return 6
    df_pnro_valmis['Uusi_Vyohyke'] = df_pnro_valmis.apply(maarita_vyohyke, axis=1)
    zone_map = df_pnro_valmis.set_index(COL_POSTINUMERO)['Uusi_Vyohyke']
    def korjaa_postilokerot(row):
        p_str = row[COL_POSTINUMERO]; current_zone = row['Uusi_Vyohyke']
        try:
            p_int = int(p_str)
            if p_int % 10 != 0:
                base_p_str = str(p_int - (p_int % 10))
                if base_p_str in zone_map: return zone_map[base_p_str]
        except (ValueError, TypeError): pass
        return current_zone
    df_pnro_valmis['Uusi_Vyohyke'] = df_pnro_valmis.apply(korjaa_postilokerot, axis=1)
    output_cols = [COL_POSTINUMERO, 'Uusi_Vyohyke', COL_X_KOORD, COL_Y_KOORD, 'Rahtikirjojen_lkm']
    return df_pnro_valmis[output_cols].rename(columns={'Uusi_Vyohyke': COL_VYOHYKE})

# --- APUFUNKTIO PORAUTUMISNÄYTÖLLE ---
def nayta_porautumisanalyysi(data_valinnalle, vanha_kustannus, otsikko):
    """Näyttää standardoidun analyysinäkymän annetulle datalle."""
    st.subheader(f"Porautumisanalyysi: {otsikko}")
    if data_valinnalle.empty:
        st.warning("Valinnalle ei löytynyt kustannusdataa.")
        return

    painoluokka_jarjestys = [get_painoluokka_str(i, st.session_state.sheets[SHEET_TARIFFI]) for i in st.session_state.sheets[SHEET_TARIFFI].index]
    data_valinnalle['Painoluokka'] = data_valinnalle['tariffi_rivi_idx'].apply(lambda idx: get_painoluokka_str(idx, st.session_state.sheets[SHEET_TARIFFI]))
    data_valinnalle['Painoluokka'] = pd.Categorical(data_valinnalle['Painoluokka'], categories=painoluokka_jarjestys, ordered=True)
    
    total_cost = data_valinnalle['Uusi_nippu_hinta'].sum()
    st.metric("Lasketut kokonaiskustannukset", f"{total_cost:,.2f} €", delta=f"{(total_cost - vanha_kustannus):,.2f} €")
    
    pivot_table_abs = pd.pivot_table(data_valinnalle, values='Uusi_nippu_hinta', index='Painoluokka', columns=COL_VYOHYKE, aggfunc='sum', fill_value=0)
    
    st.write("**Kustannusten jakautuminen (€)**")
    st.dataframe(pivot_table_abs.style.background_gradient(cmap='Greens').format("{:,.2f} €"), use_container_width=True)
    
    if total_cost > 0:
        pivot_table_perc = (pivot_table_abs / total_cost * 100)
        st.write("**Kustannusten jakautuminen (%)**")
        st.dataframe(pivot_table_perc.style.background_gradient(cmap='Blues').format("{:.2f}%"), use_container_width=True)

# --- MUUTETTU KOHTA: FUNKTIOT EXCEL-VIENTIÄ VARTEN ---
def laske_erittely_data(sheets, df_tulokset_niput, autot_mukana):
    """Valmistelee yksittäisten lähetysten datan Excel-vientiä varten."""
    df_jako = sheets[SHEET_JAKO].copy()
    df_nouto = sheets[SHEET_NOUTO].copy()
    
    df_jako['Tyyppi'] = 'Jakelu'
    df_nouto['Tyyppi'] = 'Nouto'
    
    df_kaikki_keikat = pd.concat([df_jako, df_nouto], ignore_index=True)
    df_kaikki_keikat = df_kaikki_keikat[df_kaikki_keikat[COL_AUTOTUNNUS].isin(autot_mukana)]

    # Yhdistetään nippujen hinnat JA PAINOT yksittäisiin keikkoihin
    df_erittely = pd.merge(
        df_kaikki_keikat,
        df_tulokset_niput[[COL_NIPPUNUMERO, COL_LIIKENNOITSIJA, 'nippu_paino', 'Uusi_nippu_hinta']], # LISÄTTY 'nippu_paino'
        on=COL_NIPPUNUMERO,
        how='left'
    )
    df_erittely.rename(columns={'Uusi_nippu_hinta': 'Hinta', 'nippu_paino': 'Paino'}, inplace=True) # LISÄTTY PAINON UUDELLEEN NIMEYS
    return df_erittely

def luo_tulos_exceliin(df_tariffi, df_vyohykkeet, df_erittely):
    """Luo ja palauttaa tulos-Excelin, joka sisältää myös yksittäiset lähetykset."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_tariffi.to_excel(writer, sheet_name='Lasketut_Tariffit', index=False)
        df_vyohykkeet.to_excel(writer, sheet_name='Lasketut_Vyohykkeet', index=False)
        
        # MUOKATTU SARAKKEIDEN JÄRJESTYS
        cols_to_show = [COL_AUTOTUNNUS, COL_LIIKENNOITSIJA, COL_NIPPUNUMERO, 'Paino', 'Hinta']
        
        df_erittely[df_erittely['Tyyppi'] == 'Jakelu'][cols_to_show].to_excel(
            writer, sheet_name='Jakelut_eritelty', index=False
        )
        df_erittely[df_erittely['Tyyppi'] == 'Nouto'][cols_to_show].to_excel(
            writer, sheet_name='Noudot_eritelty', index=False
        )
    return output.getvalue()
# --- MUUTOKSEN LOPPU ---

# =============================================================================
# STREAMLIT-KÄYTTÖLIITTYMÄ
# =============================================================================
st.set_page_config(layout="wide", page_title="Rahtioptimointi")

# Session State alustus
if 'app_loaded' not in st.session_state:
    st.session_state.app_loaded = True
    st.session_state.sheets = {}
    st.session_state.df_tariff_current = pd.DataFrame()
    st.session_state.df_zones_current = pd.DataFrame()
    st.session_state.df_autot_current = pd.DataFrame()
    st.session_state.vertailu_auto = pd.DataFrame()
    st.session_state.lukitut_tariffit = {}
    st.session_state.lukitut_vyohykkeet = {}
    st.session_state.last_error = ""
    st.session_state.erittely_data = pd.DataFrame() # Uusi session state Excel-vientiä varten

st.title("🚛 Rahtikustannusten optimointityökalu")

# --- SIVUPALKKI ---
with st.sidebar:
    st.header("1. Data")
    st.download_button("📥 Lataa mallipohja", luo_mallipohja_exceliin(), 'syotetiedot_malli.xlsx')
    uploaded_file = st.file_uploader("Lataa Excel-pohja", type="xlsx")

    if uploaded_file:
        if st.button("Lataa data", type="primary"):
            try:
                sheets = pd.read_excel(uploaded_file, sheet_name=None)
                virheet = validoi_syotetiedosto(sheets)
                if virheet:
                    st.session_state.last_error = "Virhe syötetiedostossa: " + ", ".join(virheet)
                else:
                    st.session_state.sheets = sheets
                    st.session_state.df_tariff_current = sheets[SHEET_TARIFFI].copy()
                    st.session_state.df_zones_current = sheets[SHEET_PNRO].copy()
                    st.session_state.df_autot_current = sheets[SHEET_AUTOT].copy()
                    st.session_state.vertailu_auto = pd.DataFrame()
                    st.session_state.lukitut_tariffit = {}
                    st.session_state.lukitut_vyohykkeet = {}
                    st.session_state.last_error = ""
                    st.session_state.erittely_data = pd.DataFrame() # Nollataan myös tämä
                    st.toast("Data ladattu!", icon="✅")
            except Exception as e:
                st.session_state.last_error = f"Tiedoston lukemisessa tapahtui odottamaton virhe: {e}"
            st.rerun()

    if not st.session_state.get('sheets'):
        st.info("Lataa data Excel-tiedostosta aloittaaksesi.")
        st.stop()

    if st.button("Nollaa muutokset ja tulokset"):
        st.session_state.df_tariff_current = st.session_state.sheets[SHEET_TARIFFI].copy()
        st.session_state.df_zones_current = st.session_state.sheets[SHEET_PNRO].copy()
        st.session_state.df_autot_current = st.session_state.sheets[SHEET_AUTOT].copy()
        st.session_state.vertailu_auto = pd.DataFrame()
        st.session_state.lukitut_tariffit = {}
        st.session_state.lukitut_vyohykkeet = {}
        st.session_state.erittely_data = pd.DataFrame() # Nollataan myös tämä
        st.toast("Kaikki muutokset ja tulokset nollattu.", icon="🔄")
        st.rerun()

    st.header("2. Yleiset parametrit")
    tasmaystaso = st.radio("Mihin hintaa täsmätään?", ('Kokonaisuus', 'Liikennöitsijä', 'Auto'), index=0, key="taso_radio")
    sallittu_heitto = st.slider("Sallittu heitto (%)", 0.5, 30.0, 5.0, 0.5, key="heitto_slider")
    
    st.header("3. Toiminnot")
    with st.expander("Vyöhykkeiden määritys", expanded=False):
        vyohyke_tapa = st.radio("Valitse toiminto:", ("Käytä alkuperäisiä (Excel)", "Generoi älykkäästi (Heuristiikka)", "Optimoi matemaattisesti (Hienosäätö)"), key="vyohyke_tapa_radio", index=0)
        if vyohyke_tapa == "Käytä alkuperäisiä (Excel)":
             if st.button("Palauta alkuperäiset vyöhykkeet"):
                st.session_state.df_zones_current = st.session_state.sheets[SHEET_PNRO].copy()
                st.session_state.lukitut_vyohykkeet = {}
                st.toast("Alkuperäiset vyöhykkeet palautettu ja lukitukset poistettu.", icon="↩️"); st.rerun()
        elif vyohyke_tapa == "Generoi älykkäästi (Heuristiikka)":
            paakeskus_pnro = st.text_input("Pääkeskuksen postinumero", "60100", key="paakeskus_input")
            if st.button("Suorita älykäs generointi"):
                try:
                    with st.spinner("Analysoidaan dataa..."):
                        df_keikat = pd.concat([st.session_state.sheets[SHEET_JAKO], st.session_state.sheets[SHEET_NOUTO]], ignore_index=True)
                        tulos = laske_vyohykkeet_automaattisesti(df_keikat, st.session_state.sheets[SHEET_PNRO], paakeskus_pnro)
                        st.session_state.df_zones_current = tulos
                        st.session_state.lukitut_vyohykkeet = {}
                    st.toast("Uusi vyöhykemalli generoitu ja lukitukset poistettu!", icon="🤖")
                except ValueError as e:
                    st.session_state.last_error = f"Vyöhykkeiden generointi epäonnistui: {e}"
                st.rerun()
        elif vyohyke_tapa == "Optimoi matemaattisesti (Hienosäätö)":
            sallittu_optimointivaje = st.slider("Sallittu optimointivaje (%)", 0.05, 10.0, 1.0, 0.1, key="vaje_slider")
            if st.button("Suorita matemaattinen optimointi"):
                with st.spinner("Optimoidaan vyöhykkeitä..."):
                    params = {'taso': tasmaystaso, 'heitto': sallittu_heitto, 'vaje': sallittu_optimointivaje, 'lukitut_vyohykkeet': st.session_state.lukitut_vyohykkeet}
                    status, tulos, vertailu = suorita_vyohyke_optimointi(st.session_state.sheets, st.session_state.df_tariff_current, list(st.session_state.df_autot_current[COL_AUTOTUNNUS]), params)
                    if status == "ok":
                        original_zones = st.session_state.sheets[SHEET_PNRO].copy().drop(columns=COL_VYOHYKE, errors='ignore')
                        original_zones[COL_POSTINUMERO] = original_zones[COL_POSTINUMERO].astype(str)
                        tulos[COL_POSTINUMERO] = tulos[COL_POSTINUMERO].astype(str)
                        new_zones_complete = pd.merge(original_zones, tulos, on=COL_POSTINUMERO, how='left')
                        for pnro, vyohyke in st.session_state.lukitut_vyohykkeet.items():
                             new_zones_complete.loc[new_zones_complete[COL_POSTINUMERO] == pnro, COL_VYOHYKE] = vyohyke
                        st.session_state.df_zones_current = new_zones_complete
                        st.session_state.vertailu_auto = vertailu
                        st.toast("Vyöhykkeet optimoitu!", icon="🎯")
                    else: st.session_state.last_error = tulos
                st.rerun()

    with st.expander("Tariffien laskenta", expanded=True):
        minimi_korotus = st.slider("MINIMIKOROTUS (%)", 0.0, 5.0, 0.1, 0.01, key="min_korotus_slider")
        max_korotus = st.slider("MAKSIMIKOROTUS (%)", 1.0, 20.0, 5.0, 0.1, key="max_korotus_slider")
        if st.button("Laske uudet tariffit", type="primary"):
            with st.spinner("Lasketaan tariffeja..."):
                params = {'taso': tasmaystaso, 'heitto': sallittu_heitto, 'min_korotus': minimi_korotus, 'max_korotus': max_korotus, 'lukitut_tariffit': st.session_state.lukitut_tariffit}
                status, tulos, vertailu = suorita_tariffi_optimointi(st.session_state.sheets, st.session_state.df_zones_current, list(st.session_state.df_autot_current[COL_AUTOTUNNUS]), params)
                if status == "ok": 
                    st.session_state.df_tariff_current = tulos
                    st.session_state.vertailu_auto = vertailu
                    st.toast("Uudet tariffit laskettu!", icon="💰")
                else: 
                    st.session_state.last_error = tulos
            st.rerun()
            
    st.header("4. Tallenna & Vie")
    # --- MUUTETTU KOHTA: Kaksivaiheinen vienti ---
    if st.button("Valmistele Excel-raportti"):
        if st.session_state.get("df_tulokset_yksiloity") is not None and not st.session_state.df_tulokset_yksiloity.empty:
            st.session_state.erittely_data = laske_erittely_data(
                st.session_state.sheets,
                st.session_state.df_tulokset_yksiloity,
                list(st.session_state.df_autot_current[COL_AUTOTUNNUS])
            )
            st.toast("Raportin data valmis ladattavaksi!", icon="📊")
        else:
            st.warning("Aja ensin laskenta ja valitse autoja, jotta raportti voidaan luoda.")

    if not st.session_state.erittely_data.empty:
        st.download_button(
            label="💾 Lataa Excel-raportti",
            data=luo_tulos_exceliin(st.session_state.df_tariff_current, st.session_state.df_zones_current, st.session_state.erittely_data),
            file_name="optimoinnin_raportti.xlsx",
            mime="application/vnd.ms-excel"
        )
    # --- MUUTOKSEN LOPPU ---

# --- PÄÄNÄYTTÖ ---
# ... (loput koodista pysyy täysin samana, lukuunottamatta yhtä kohtaa) ...
if st.session_state.last_error: 
    st.error(st.session_state.last_error)
    st.session_state.last_error = ""

st.subheader("Nykyinen tariffitaulukko")
edited_tariff = st.data_editor(st.session_state.df_tariff_current, key="tariff_editor", use_container_width=True)

if not edited_tariff.equals(st.session_state.df_tariff_current):
    muutokset = {}
    vyohyke_cols = [c for c in edited_tariff.columns if PREFIX_VYOHYKE_COL in c]
    for r_idx, row in edited_tariff.iterrows():
        for c_name in vyohyke_cols:
            orig_val = st.session_state.df_tariff_current.at[r_idx, c_name]
            new_val = row[c_name]
            if pd.notna(new_val) and (pd.isna(orig_val) or not np.isclose(float(new_val), float(orig_val))):
                muutokset[(r_idx, c_name)] = float(new_val)
    
    st.session_state.df_tariff_current = edited_tariff.copy()
    st.session_state.lukitut_tariffit.update(muutokset)
    st.info("Tariffimuutokset tallennettu. Ne lukitaan seuraavassa laskennassa.")
    st.rerun()

if st.session_state.lukitut_tariffit:
    with st.expander(f"Aktiiviset tariffilukitukset ({len(st.session_state.lukitut_tariffit)} kpl)"):
        for (r, c) in list(st.session_state.lukitut_tariffit.keys()):
            col1, col2 = st.columns([0.8, 0.2])
            with col1:
                v = st.session_state.lukitut_tariffit[(r,c)]
                st.markdown(f"- **Rivi {r}** ({get_painoluokka_str(r, edited_tariff)}), **Sarake {c}**: **{v:.4f}**")
            with col2:
                if st.button("Poista", key=f"del_tariff_{r}_{c.replace(' ', '_')}", use_container_width=True):
                    del st.session_state.lukitut_tariffit[(r,c)]
                    st.toast(f"Lukitus poistettu: Rivi {r}, Sarake {c}", icon="🔓")
                    st.rerun()

st.subheader("Nykyinen vyöhykemalli")
col1, col2 = st.columns([0.4, 0.6])
with col1:
    df_zones_display = st.session_state.df_zones_current.copy()
    
    if COL_POSTITOIMIPAIKKA not in df_zones_display.columns:
        df_pnro_orig = st.session_state.sheets[SHEET_PNRO].copy()
        df_pnro_orig[COL_POSTINUMERO] = df_pnro_orig[COL_POSTINUMERO].astype(str)
        df_zones_display[COL_POSTINUMERO] = df_zones_display[COL_POSTINUMERO].astype(str)
        df_zones_display = pd.merge(df_zones_display, df_pnro_orig[[COL_POSTINUMERO, COL_POSTITOIMIPAIKKA]], on=COL_POSTINUMERO, how='left')

    df_keikat_temp = pd.concat([st.session_state.sheets[SHEET_JAKO], st.session_state.sheets[SHEET_NOUTO]], ignore_index=True)
    
    df_keikat_temp[COL_POSTINUMERO] = df_keikat_temp[COL_POSTINUMERO].astype(str)
    df_zones_display[COL_POSTINUMERO] = df_zones_display[COL_POSTINUMERO].astype(str)

    rk_lkm = df_keikat_temp.groupby(COL_POSTINUMERO)[COL_RAHTIKIRJA].nunique().reset_index(name='Rahtikirjojen_lkm')
    
    if 'Rahtikirjojen_lkm' in df_zones_display.columns:
        df_zones_display = df_zones_display.drop(columns=['Rahtikirjojen_lkm'])
        
    df_zones_display = pd.merge(df_zones_display, rk_lkm, on=COL_POSTINUMERO, how='left').fillna({'Rahtikirjojen_lkm': 0})
    
    display_cols = [COL_POSTINUMERO, COL_POSTITOIMIPAIKKA, COL_VYOHYKE, 'Rahtikirjojen_lkm']
    for col in display_cols:
        if col not in df_zones_display.columns: df_zones_display[col] = np.nan
    df_zones_display['Rahtikirjojen_lkm'] = df_zones_display['Rahtikirjojen_lkm'].astype(int)

    cols_to_show = [c for c in display_cols if c in df_zones_display.columns]
    edited_zones = st.data_editor(df_zones_display[cols_to_show], key="zones_editor", use_container_width=True, height=400, disabled=[COL_POSTITOIMIPAIKKA, 'Rahtikirjojen_lkm'])
    
    if not edited_zones.equals(df_zones_display[cols_to_show]):
        merged_df = pd.merge(
            st.session_state.df_zones_current[[COL_POSTINUMERO, COL_VYOHYKE]].rename(columns={COL_VYOHYKE: 'vanha_vyohyke'}),
            edited_zones[[COL_POSTINUMERO, COL_VYOHYKE]].rename(columns={COL_VYOHYKE: 'uusi_vyohyke'}),
            on=COL_POSTINUMERO, how='inner'
        )
        muuttuneet_rivit = merged_df[merged_df['vanha_vyohyke'].astype(str) != merged_df['uusi_vyohyke'].astype(str)]
        
        for _, row in muuttuneet_rivit.iterrows():
            st.session_state.lukitut_vyohykkeet[row[COL_POSTINUMERO]] = row['uusi_vyohyke']
        
        vyohyke_updates = edited_zones.set_index(COL_POSTINUMERO)[COL_VYOHYKE]
        df_to_update = st.session_state.df_zones_current.set_index(COL_POSTINUMERO)
        df_to_update.update(vyohyke_updates)
        st.session_state.df_zones_current = df_to_update.reset_index()
        
        st.info("Vyöhykemuutokset tallennettu. Ne lukitaan seuraavassa optimoinnissa.")
        st.rerun()

    if st.session_state.lukitut_vyohykkeet:
        with st.expander(f"Aktiiviset vyöhykelukitukset ({len(st.session_state.lukitut_vyohykkeet)} kpl)"):
            for pnro in list(st.session_state.lukitut_vyohykkeet.keys()):
                col_text, col_button = st.columns([4, 1])
                with col_text:
                    v = st.session_state.lukitut_vyohykkeet[pnro]
                    st.markdown(f"- **{pnro}** → Vyöhyke **{int(v) if pd.notna(v) else 'Tyhjä'}**")
                with col_button:
                    if st.button("Poista", key=f"del_zone_{pnro}", use_container_width=True):
                        del st.session_state.lukitut_vyohykkeet[pnro]
                        orig_pnro_df = st.session_state.sheets[SHEET_PNRO].astype({COL_POSTINUMERO: str})
                        orig_row = orig_pnro_df[orig_pnro_df[COL_POSTINUMERO] == str(pnro)]
                        if not orig_row.empty:
                            orig_value = orig_row[COL_VYOHYKE].iloc[0]
                            st.session_state.df_zones_current.loc[st.session_state.df_zones_current[COL_POSTINUMERO].astype(str) == str(pnro), COL_VYOHYKE] = orig_value
                        st.toast(f"Lukitus poistettu: {pnro}", icon="🔓")
                        st.rerun()

with col2:
    df_map = st.session_state.df_zones_current.copy()
    df_map.replace('EILÖYDY', np.nan, inplace=True); df_map.dropna(subset=[COL_X_KOORD, COL_Y_KOORD, COL_VYOHYKE], inplace=True)
    if not df_map.empty:
        try:
            transformer = Transformer.from_crs("EPSG:3067", "EPSG:4326", always_xy=True)
            df_map['lon'], df_map['lat'] = transformer.transform(df_map[COL_X_KOORD].values, df_map[COL_Y_KOORD].values)
            df_map[COL_VYOHYKE] = df_map[COL_VYOHYKE].astype(int)
            colors = [[33, 150, 243, 160], [100, 181, 246, 160], [255, 235, 59, 160], [255, 193, 7, 160], [255, 87, 34, 160], [213, 0, 0, 160]]
            df_map['color'] = df_map[COL_VYOHYKE].apply(lambda z: colors[min(z - 1, len(colors) - 1)])
            st.pydeck_chart(pdk.Deck(
                map_provider="carto", map_style="light",
                initial_view_state=pdk.ViewState(latitude=df_map['lat'].mean(), longitude=df_map['lon'].mean(), zoom=7, pitch=0),
                layers=[pdk.Layer('ScatterplotLayer', data=df_map, get_position='[lon, lat]', get_fill_color='color', get_radius=1500, pickable=True)],
                tooltip={"text": f"{COL_POSTINUMERO}: {{{COL_POSTINUMERO}}}\n{COL_VYOHYKE}: {{{COL_VYOHYKE}}}"}
            ))
        except Exception as e: st.warning(f"Karttavisualisoinnin luonti epäonnistui: {e}")
    else: st.info("Ei näytettävää dataa kartalla.")

if not st.session_state.vertailu_auto.empty:
    st.header("Laskennan tulokset")
     # --- LISÄTTY KOHTA: Laske ja päivitä oikeat kustannukset päätaulukkoon ---
    # 1. Laske ensin kaikkien yksittäisten nippujen oikeat hinnat
    df_tariff_orig = st.session_state.sheets[SHEET_TARIFFI]
    _, df_niput_base, _ = _valmistele_data(st.session_state.sheets, st.session_state.sheets[SHEET_AUTOT][COL_AUTOTUNNUS])
    df_tulokset_yksiloity_temp = pd.merge(df_niput_base, st.session_state.df_zones_current[[COL_POSTINUMERO, COL_VYOHYKE]], on=COL_POSTINUMERO, how='inner')
    df_tulokset_yksiloity_temp[COL_VYOHYKE] = pd.to_numeric(df_tulokset_yksiloity_temp[COL_VYOHYKE], errors='coerce').fillna(0).astype(int)
    if not df_tulokset_yksiloity_temp.empty:
        df_tulokset_yksiloity_temp['Uusi_nippu_hinta'] = df_tulokset_yksiloity_temp.apply(
            laske_oikea_nippu_hinta, axis=1,
            df_tariff_current=st.session_state.df_tariff_current,
            df_tariff_orig=df_tariff_orig
        )
        # 2. Laske oikeat kokonaissummat per auto
        korjatut_summat = df_tulokset_yksiloity_temp.groupby(COL_AUTOTUNNUS)['Uusi_nippu_hinta'].sum()
        
        # 3. Päivitä nämä oikeat summat session stateen tallennettuun vertailutaulukkoon
        st.session_state.vertailu_auto['Uusi kustannus (€)'] = st.session_state.vertailu_auto[COL_AUTOTUNNUS].map(korjatut_summat).fillna(0)
    # --- LISÄtyn LOHKON LOPPU ---
    df_vertailu = st.session_state.vertailu_auto.copy()
    df_orig_autot = st.session_state.sheets[SHEET_AUTOT][[COL_AUTOTUNNUS, COL_LIIKENNOITSIJA]]
    df_vertailu = pd.merge(df_vertailu, df_orig_autot, on=COL_AUTOTUNNUS, how='left')
    df_vertailu['Erotus (€)'] = df_vertailu['Uusi kustannus (€)'] - df_vertailu['Vanha kustannus (€)']
    df_vertailu['Erotus (%)'] = (df_vertailu['Vanha kustannus (€)'].replace(0, np.nan))
    df_vertailu['Erotus (%)'] = df_vertailu['Erotus (€)'] / df_vertailu['Erotus (%)'] * 100
    df_vertailu.fillna(0, inplace=True)

    st.subheader("Autojen valinta ja vertailu")
    df_vertailu['Mukana'] = df_vertailu[COL_AUTOTUNNUS].isin(list(st.session_state.df_autot_current[COL_AUTOTUNNUS]))
    
    display_cols_autot = ['Mukana', COL_AUTOTUNNUS, COL_LIIKENNOITSIJA, 'Vanha kustannus (€)', 'Uusi kustannus (€)', 'Erotus (€)', 'Erotus (%)']
    edited_autot = st.data_editor(df_vertailu[display_cols_autot], key="autot_editor", hide_index=True, use_container_width=True)

    autot_nyt_mukana = set(edited_autot[edited_autot['Mukana']][COL_AUTOTUNNUS])
    autot_ennen = set(st.session_state.df_autot_current[COL_AUTOTUNNUS])
    if autot_nyt_mukana != autot_ennen:
        st.session_state.df_autot_current = st.session_state.sheets[SHEET_AUTOT][st.session_state.sheets[SHEET_AUTOT][COL_AUTOTUNNUS].isin(autot_nyt_mukana)].copy()
        st.info("Autojen valinta on muuttunut. Aja haluamasi laskenta uudelleen päivittääksesi tulokset.")
    
    df_naytettava = edited_autot[edited_autot['Mukana']]
    if not df_naytettava.empty:
               # --- KORJATTU KOHTA: Tulosten laskenta ja tallennus session stateen ---
        df_tariff_orig = st.session_state.sheets[SHEET_TARIFFI]
        _, df_niput_base, _ = _valmistele_data(st.session_state.sheets, list(autot_nyt_mukana))
        
        df_tulokset_yksiloity = pd.merge(df_niput_base, st.session_state.df_zones_current[[COL_POSTINUMERO, COL_VYOHYKE]], on=COL_POSTINUMERO, how='inner')
        df_tulokset_yksiloity = pd.merge(df_tulokset_yksiloity, df_naytettava[[COL_AUTOTUNNUS, COL_LIIKENNOITSIJA]], on=COL_AUTOTUNNUS, how='left')
        df_tulokset_yksiloity.dropna(subset=[COL_LIIKENNOITSIJA], inplace=True)

        # Käytetään ensin alkuperäistä nimeä 'nippu_paino' laskentaan
        df_tulokset_yksiloity['tariffi_rivi_idx'] = df_tulokset_yksiloity['nippu_paino'].apply(lambda p: get_painoluokka_rivi_idx(p, df_tariff_orig))
        
        # Nimetään sarake uudelleen 'Paino'-nimiseksi vasta kun sitä ei enää tarvita laskennassa
        # --- kommataan virheellinen rivi df_tulokset_yksiloity.rename(columns={'nippu_paino': 'Paino'}, inplace=True)
        # --- KORJAUKSEN LOPPU ---

        
        df_tulokset_yksiloity.dropna(subset=[COL_LIIKENNOITSIJA], inplace=True) # Varmistetaan, että vain valitut autot ovat mukana
        df_tulokset_yksiloity['tariffi_rivi_idx'] = df_tulokset_yksiloity['nippu_paino'].apply(lambda p: get_painoluokka_rivi_idx(p, df_tariff_orig))
        df_tulokset_yksiloity.dropna(subset=['tariffi_rivi_idx'], inplace=True)
        df_tulokset_yksiloity[COL_VYOHYKE] = pd.to_numeric(df_tulokset_yksiloity[COL_VYOHYKE], errors='coerce').fillna(0).astype(int)
        valid_vyohykkeet = {int(c.split(' ')[1]) for c in st.session_state.df_tariff_current.columns if PREFIX_VYOHYKE_COL in c}
        df_tulokset_yksiloity = df_tulokset_yksiloity[df_tulokset_yksiloity[COL_VYOHYKE].isin(valid_vyohykkeet)]

        if not df_tulokset_yksiloity.empty:
            df_tulokset_yksiloity['Uusi_nippu_hinta'] = df_tulokset_yksiloity.apply(
                laske_oikea_nippu_hinta, axis=1,
                df_tariff_current=st.session_state.df_tariff_current,
                df_tariff_orig=df_tariff_orig
            )
        st.session_state.df_tulokset_yksiloity = df_tulokset_yksiloity
        # --- MUUTOKSEN LOPPU ---
        
        st.write("**Yhteenvedot (perustuen valittuihin autoihin):**")
        summa_auto = pd.DataFrame(df_naytettava[['Vanha kustannus (€)', 'Uusi kustannus (€)', 'Erotus (€)']].sum()).T
        summa_auto[COL_AUTOTUNNUS] = 'YHTEENSÄ'
        st.dataframe(summa_auto.set_index(COL_AUTOTUNNUS).style.format("{:,.2f} €"), use_container_width=True)
        st.subheader("Liikennöitsijäkohtainen yhteenveto")
        df_liikenne = df_naytettava.groupby(COL_LIIKENNOITSIJA)[['Vanha kustannus (€)', 'Uusi kustannus (€)', 'Erotus (€)']].sum().reset_index()
        df_liikenne['Erotus (%)'] = (df_liikenne['Vanha kustannus (€)'].replace(0, np.nan))
        df_liikenne['Erotus (%)'] = df_liikenne['Erotus (€)'] / df_liikenne['Erotus (%)'] * 100
        df_liikenne.fillna(0, inplace=True)
        st.dataframe(df_liikenne, hide_index=True, use_container_width=True)
        
        st.markdown("---")
        st.header("Porautumisanalyysi")

        col1_drill, col2_drill = st.columns(2)
        with col1_drill:
            liik_lista = ["Valitse..."] + sorted(df_naytettava[COL_LIIKENNOITSIJA].unique())
            valittu_liik = st.selectbox("Valitse liikennöitsijä", liik_lista)
        with col2_drill:
            auto_lista = ["Valitse..."] + sorted(df_naytettava[COL_AUTOTUNNUS].unique())
            valittu_auto = st.selectbox("Valitse auto", auto_lista)

        if valittu_liik != "Valitse...":
            data_filt = st.session_state.df_tulokset_yksiloity[st.session_state.df_tulokset_yksiloity[COL_LIIKENNOITSIJA] == valittu_liik]
            vanha_kustannus = df_liikenne[df_liikenne[COL_LIIKENNOITSIJA] == valittu_liik]['Vanha kustannus (€)'].sum()
            nayta_porautumisanalyysi(data_filt, vanha_kustannus, valittu_liik)
        
        elif valittu_auto != "Valitse...":
            data_filt = st.session_state.df_tulokset_yksiloity[st.session_state.df_tulokset_yksiloity[COL_AUTOTUNNUS] == valittu_auto]
            vanha_kustannus = df_naytettava[df_naytettava[COL_AUTOTUNNUS] == valittu_auto]['Vanha kustannus (€)'].sum()
            nayta_porautumisanalyysi(data_filt, vanha_kustannus, valittu_auto)
