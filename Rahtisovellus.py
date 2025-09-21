import streamlit as st
import pandas as pd
import numpy as np
import pyomo.environ as pyo
from io import BytesIO
import time
import pydeck as pdk
from pyproj import Transformer

# =============================================================================
# APUFUNKTIOT JA DATALÄHTEET
# = "==========================================================================="
def luo_mallipohja_exceliin():
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df1 = pd.DataFrame({'Autotunnus': ['AUTO-001', 'AUTO-002'],'Liikennöitsijän nimi': ['Kuljetus A Oy', 'Siirto B Tmi'],'Kulut entisellä mallilla': [15500.50, 22340.00]})
        df1.to_excel(writer, sheet_name='AutojenYhteenveto', index=False)
        df2 = pd.DataFrame({'Autotunnus': ['AUTO-001', 'AUTO-001', 'AUTO-002'],'Rahtikirjanumero': ['RK123', 'RK124', 'RK125'],'Vastaanottaja': ['Yritys X', 'Yritys Y', 'Yritys Z'],'Postinumero': ['60100', '61800', '61330'],'Kilot': [850, 120, 25000],'Nippunumero': ['NIPPU-01', 'NIPPU-01', 'NIPPU-02']})
        df2.to_excel(writer, sheet_name='Jakokeikat', index=False)
        df3 = pd.DataFrame({'Autotunnus': ['AUTO-001', 'AUTO-002'],'Rahtikirjanumero': ['RK501', 'RK502'],'Noutopaikka': ['Varasto A', 'Tehdas B'],'Postinumero': ['60220', '60800'],'Kilot': [1500, 3450],'Nippunumero': ['NIPPU-03', 'NIPPU-04']})
        df3.to_excel(writer, sheet_name='Noutokeikat', index=False)
        
        # UUTTA: Lisätty Postitoimipaikka-sarake mallipohjaan
        df4_data = {
            'Postinumero': ['60100', '60101', '60220', '61800', '61330', '60800', '60801'],
            'Postitoimipaikka': ['Seinäjoki Keskus', 'Seinäjoki PL', 'Pajuluoma', 'Kauhajoki', 'Jalasjärvi kk', 'Ilmajoki', 'Ilmajoki PL'],
            'Vyöhyke': [1, 1, 1, 3, 5, 2, 2],
            'X-Koordinaatti': [288812, 'EILÖYDY', 286625, 274433, 271234, 273247, 'EILÖYDY'],
            'Y-Koordinaatti': [6967917, 'EILÖYDY', 6967422, 6940011, 6932111, 6964498, 'EILÖYDY']
        }
        df4 = pd.DataFrame(df4_data)
        df4.to_excel(writer, sheet_name='Postinumerot', index=False)
        
        df5_data = {'Painoluokka ALKAA (kg)': [0, 501, 10001, 20001],'Painoluokka LOPPUU (kg)': [500, 10000, 20000, np.nan],'Laskentatapa': ['€/nippu', '€/nippu', '€/nippu', '€/kg'],'VYÖHYKE 1': [50.0, 120.0, 200.0, 0.008],'VYÖHYKE 2': [65.0, 150.0, 220.0, 0.010],'VYÖHYKE 3': [80.0, 180.0, 250.0, 0.012],'VYÖHYKE 4': [90.0, 200.0, 280.0, 0.014],'VYÖHYKE 5': [100.0, 220.0, 310.0, 0.016],'VYÖHYKE 6': [120.0, 250.0, 350.0, 0.018],}
        df5 = pd.DataFrame(df5_data); df5.to_excel(writer, sheet_name='Tariffitaulukko', index=False)
    return output.getvalue()

@st.cache_data
def muunna_df_exceliksi(df):
    output = BytesIO()
    if isinstance(df, pd.io.formats.style.Styler): df = df.data
    with pd.ExcelWriter(output, engine='openpyxl') as writer: df.to_excel(writer, index=False, sheet_name='Tulokset')
    return output.getvalue()

def get_painoluokka_str(idx, df_tariff):
    rivi = df_tariff.loc[idx]; alku = rivi['Painoluokka ALKAA (kg)']; loppu = rivi['Painoluokka LOPPUU (kg)']
    if pd.isna(loppu): return f"> {int(alku) - 1} kg"
    return f"{int(alku)}-{int(loppu)} kg"

def get_painoluokka_rivi_idx(paino, df_tariff):
    df_tariff_copy = df_tariff.copy(); df_tariff_copy['Painoluokka LOPPUU (kg)'] = df_tariff_copy['Painoluokka LOPPUU (kg)'].fillna(np.inf)
    sopivat_rivit = df_tariff_copy[(df_tariff_copy['Painoluokka ALKAA (kg)'] <= paino) & (df_tariff_copy['Painoluokka LOPPUU (kg)'] >= paino)]
    if not sopivat_rivit.empty: return sopivat_rivit.index[0]
    return None

# =============================================================================
# OPTIMOINTI JA ANALYTIIKKA -FUNKTIOT
# (Nämä pysyvät ennallaan)
# =============================================================================
def _valmistele_data(sheets, autot_mukana):
    df_autot_orig = sheets['AutojenYhteenveto']
    df_autot = df_autot_orig[df_autot_orig['Autotunnus'].isin(autot_mukana)].copy()
    df_keikat = pd.concat([sheets['Jakokeikat'], sheets['Noutokeikat']], ignore_index=True)
    df_keikat = df_keikat[df_keikat['Autotunnus'].isin(autot_mukana)]
    df_niput = df_keikat.groupby('Nippunumero').agg(nippu_paino=('Kilot', 'sum'), Autotunnus=('Autotunnus', 'first'), Postinumero=('Postinumero', 'first')).reset_index()
    df_niput['Postinumero'] = df_niput['Postinumero'].astype(str)
    return df_autot, df_niput, sheets['Tariffitaulukko']

def suorita_tariffi_optimointi(sheets, df_zones_current, autot_mukana, params):
    df_autot, df_niput_base, df_tariff_input = _valmistele_data(sheets, autot_mukana)
    df_niput_base['tariffi_rivi_idx'] = df_niput_base['nippu_paino'].apply(lambda p: get_painoluokka_rivi_idx(p, df_tariff_input))
    df_niput_base.dropna(subset=['tariffi_rivi_idx'], inplace=True)
    df_zones_current['Postinumero'] = df_zones_current['Postinumero'].astype(str)
    df_niput = pd.merge(df_niput_base, df_zones_current, on='Postinumero', how='inner')
    df_niput['Vyöhyke'] = pd.to_numeric(df_niput['Vyöhyke'], errors='coerce').fillna(0).astype(int)
    if df_niput.empty: return "virhe", "Datan yhdistelyn jälkeen ei jäänyt yhtään käsiteltävää nippua.", None
    model = pyo.ConcreteModel()
    autot = list(df_autot['Autotunnus']); model.AUTOT = pyo.Set(initialize=autot)
    vyohyke_sarakkeet = [c for c in df_tariff_input.columns if 'VYÖHYKE' in c]
    vyohykkeet = sorted([int(c.split(' ')[1]) for c in vyohyke_sarakkeet]); model.VYOHYKKEET = pyo.Set(initialize=vyohykkeet)
    tariffi_rivit = list(df_tariff_input.index); model.TARIFFI_RIVIT = pyo.Set(initialize=tariffi_rivit)
    model.tariffi = pyo.Var(model.TARIFFI_RIVIT, model.VYOHYKKEET, within=pyo.NonNegativeReals)
    @model.Expression(model.AUTOT)
    def uusi_kustannus_per_auto(m, auto):
        auton_niput = df_niput[df_niput['Autotunnus'] == auto]; total_cost = 0
        for _, nippu in auton_niput.iterrows():
            rivi_idx, vyohyke = int(nippu['tariffi_rivi_idx']), int(nippu['Vyöhyke'])
            if vyohyke not in m.VYOHYKKEET: continue
            laskentatapa = df_tariff_input.at[rivi_idx, 'Laskentatapa']
            cost = m.tariffi[rivi_idx, vyohyke]; total_cost += nippu['nippu_paino'] * cost if laskentatapa == '€/kg' else cost
        return total_cost
    vanhat_kulut_dict_auto = df_autot.set_index('Autotunnus')['Kulut entisellä mallilla'].to_dict()
    heitto_kerroin = params['heitto'] / 100.0
    model.pos_ero = pyo.Var(model.AUTOT, within=pyo.NonNegativeReals); model.neg_ero = pyo.Var(model.AUTOT, within=pyo.NonNegativeReals)
    model.tavoite = pyo.Objective(rule=lambda m: sum(m.pos_ero[a] + m.neg_ero[a] for a in m.AUTOT), sense=pyo.minimize)
    @model.Constraint(model.AUTOT)
    def erotus_rajoite(m, a): return m.uusi_kustannus_per_auto[a] - vanhat_kulut_dict_auto.get(a, 0) == m.pos_ero[a] - m.neg_ero[a]
    if params['taso'] == 'Auto':
        @model.Constraint(model.AUTOT)
        def kustannusrajoite(m, a): vanha = vanhat_kulut_dict_auto.get(a,0); return pyo.inequality(vanha * (1-heitto_kerroin), m.uusi_kustannus_per_auto[a], vanha * (1+heitto_kerroin))
    elif params['taso'] == 'Liikennöitsijä':
        liikennöitsijat = df_autot['Liikennöitsijän nimi'].unique(); model.LIIKENNOITSIJAT = pyo.Set(initialize=liikennöitsijat)
        vanhat_kulut_liikennöitsijä = df_autot.groupby('Liikennöitsijän nimi')['Kulut entisellä mallilla'].sum()
        @model.Expression(model.LIIKENNOITSIJAT)
        def uusi_kustannus_per_liikennöitsijä(m, l): return sum(m.uusi_kustannus_per_auto[a] for a in df_autot[df_autot['Liikennöitsijän nimi'] == l]['Autotunnus'])
        @model.Constraint(model.LIIKENNOITSIJAT)
        def kustannusrajoite(m, l): vanha = vanhat_kulut_liikennöitsijä.get(l, 0); return pyo.inequality(vanha * (1 - heitto_kerroin), m.uusi_kustannus_per_liikennöitsijä[l], vanha * (1 + heitto_kerroin))
    else:
        vanhat_kulut_yhteensä = df_autot['Kulut entisellä mallilla'].sum(); uudet_kulut_yhteensä = sum(model.uusi_kustannus_per_auto[a] for a in model.AUTOT)
        model.kustannusrajoite = pyo.Constraint(rule=pyo.inequality(vanhat_kulut_yhteensä * (1 - heitto_kerroin), uudet_kulut_yhteensä, vanhat_kulut_yhteensä * (1 + heitto_kerroin)))
    model.lukitut_rajoitteet = pyo.ConstraintList()
    for (r, v_str), arvo in params['lukitut_tariffit'].items():
        v = int(v_str.split(' ')[1]); model.lukitut_rajoitteet.add(model.tariffi[r, v] == arvo)
    MIN_KERROIN = 1.0 + (params['min_korotus'] / 100.0); MAX_KERROIN = 1.0 + (params['max_korotus'] / 100.0)
    @model.Constraint(model.TARIFFI_RIVIT, model.VYOHYKKEET)
    def monotonisuus_paino_MIN(m,r,v):
        if r + 1 in m.TARIFFI_RIVIT and df_tariff_input.at[r,'Laskentatapa'] == df_tariff_input.at[r+1,'Laskentatapa']: return m.tariffi[r+1,v] >= m.tariffi[r,v] * MIN_KERROIN
        return pyo.Constraint.Skip
    @model.Constraint(model.TARIFFI_RIVIT, model.VYOHYKKEET)
    def monotonisuus_paino_MAX(m,r,v):
        if r + 1 in m.TARIFFI_RIVIT and df_tariff_input.at[r,'Laskentatapa'] == df_tariff_input.at[r+1,'Laskentatapa']: return m.tariffi[r+1,v] <= m.tariffi[r,v] * MAX_KERROIN
        return pyo.Constraint.Skip
    @model.Constraint(model.TARIFFI_RIVIT, model.VYOHYKKEET)
    def monotonisuus_vyohyke_MIN(m,r,v):
        if v + 1 in m.VYOHYKKEET: return m.tariffi[r,v+1] >= m.tariffi[r,v] * MIN_KERROIN
        return pyo.Constraint.Skip
    @model.Constraint(model.TARIFFI_RIVIT, model.VYOHYKKEET)
    def monotonisuus_vyohyke_MAX(m,r,v):
        if v + 1 in m.VYOHYKKEET: return m.tariffi[r,v+1] <= m.tariffi[r,v] * MAX_KERROIN
        return pyo.Constraint.Skip
    solver = pyo.SolverFactory('cbc'); results = solver.solve(model, tee=False)
    if (results.solver.status == pyo.SolverStatus.ok) and (results.solver.termination_condition == pyo.TerminationCondition.optimal):
        df_tulos = df_tariff_input.copy()
        for r in model.TARIFFI_RIVIT:
            for v, col in zip(model.VYOHYKKEET, vyohyke_sarakkeet): df_tulos.at[r, col] = round(pyo.value(model.tariffi[r, v]), 4)
        df_vertailu_auto = pd.DataFrame([{'Autotunnus': a, 'Vanha kustannus (€)': vanhat_kulut_dict_auto.get(a, 0), 'Uusi kustannus (€)': pyo.value(model.uusi_kustannus_per_auto[a])} for a in model.AUTOT])
        return "ok", df_tulos, df_vertailu_auto
    else: return "virhe", "Ratkaisua ei löytynyt. Kokeile löysempiä parametreja tai poista lukituksia.", None

def suorita_vyohyke_optimointi(sheets, df_tariff_current, autot_mukana, params):
    df_autot, df_niput_base, df_tariff_input = _valmistele_data(sheets, autot_mukana)
    df_niput_base['tariffi_rivi_idx'] = df_niput_base['nippu_paino'].apply(lambda p: get_painoluokka_rivi_idx(p, df_tariff_input))
    df_niput = df_niput_base.dropna(subset=['tariffi_rivi_idx']).copy()
    if df_niput.empty: return "virhe", "Datan valmistelun jälkeen ei jäänyt yhtään käsiteltävää nippua.", None
    model = pyo.ConcreteModel()
    autot = list(df_autot['Autotunnus']); model.AUTOT = pyo.Set(initialize=autot)
    postinumerot = df_niput['Postinumero'].unique(); model.POSTINUMEROT = pyo.Set(initialize=postinumerot)
    vyohyke_sarakkeet = [c for c in df_tariff_current.columns if 'VYÖHYKE' in c]
    vyohykkeet = sorted([int(c.split(' ')[1]) for c in vyohyke_sarakkeet]); model.VYOHYKKEET = pyo.Set(initialize=vyohykkeet)
    model.y = pyo.Var(model.POSTINUMEROT, model.VYOHYKKEET, within=pyo.Binary)
    @model.Constraint(model.POSTINUMEROT)
    def vain_yksi_vyohyke_per_pnro(m, p): return sum(m.y[p, v] for v in m.VYOHYKKEET) == 1
    model.lukitut_vyohykkeet = pyo.ConstraintList()
    for pnro, vyohyke in params['lukitut_vyohykkeet'].items():
        if pnro in model.POSTINUMEROT: model.lukitut_vyohykkeet.add(model.y[pnro, int(vyohyke)] == 1)
    tariffi_dict = {(r_idx, v_idx): row[v_col] for r_idx, row in df_tariff_current.iterrows() for v_idx, v_col in zip(vyohykkeet, vyohyke_sarakkeet)}
    @model.Expression(model.AUTOT)
    def uusi_kustannus_per_auto(m, auto):
        auton_niput = df_niput[df_niput['Autotunnus'] == auto]; total_cost = 0
        for _, nippu in auton_niput.iterrows():
            pnro, rivi_idx = nippu['Postinumero'], int(nippu['tariffi_rivi_idx']); laskentatapa = df_tariff_input.at[rivi_idx, 'Laskentatapa']
            nippu_hinta = sum(m.y[pnro, v] * tariffi_dict[(rivi_idx, v)] for v in m.VYOHYKKEET)
            total_cost += nippu['nippu_paino'] * nippu_hinta if laskentatapa == '€/kg' else nippu_hinta
        return total_cost
    vanhat_kulut_dict_auto = df_autot.set_index('Autotunnus')['Kulut entisellä mallilla'].to_dict()
    heitto_kerroin = params['heitto'] / 100.0
    model.pos_ero = pyo.Var(model.AUTOT, within=pyo.NonNegativeReals); model.neg_ero = pyo.Var(model.AUTOT, within=pyo.NonNegativeReals)
    model.tavoite = pyo.Objective(rule=lambda m: sum(m.pos_ero[a] + m.neg_ero[a] for a in m.AUTOT), sense=pyo.minimize)
    @model.Constraint(model.AUTOT)
    def erotus_rajoite(m, a): return m.uusi_kustannus_per_auto[a] - vanhat_kulut_dict_auto.get(a, 0) == m.pos_ero[a] - m.neg_ero[a]
    if params['taso'] == 'Auto':
        @model.Constraint(model.AUTOT)
        def kustannusrajoite(m, a): vanha = vanhat_kulut_dict_auto.get(a,0); return pyo.inequality(vanha * (1-heitto_kerroin), m.uusi_kustannus_per_auto[a], vanha * (1+heitto_kerroin))
    elif params['taso'] == 'Liikennöitsijä':
        liikennöitsijat = df_autot['Liikennöitsijän nimi'].unique(); model.LIIKENNOITSIJAT = pyo.Set(initialize=liikennöitsijat)
        vanhat_kulut_liikennöitsijä = df_autot.groupby('Liikennöitsijän nimi')['Kulut entisellä mallilla'].sum()
        @model.Expression(model.LIIKENNOITSIJAT)
        def uusi_kustannus_per_liikennöitsijä(m, l): return sum(m.uusi_kustannus_per_auto[a] for a in df_autot[df_autot['Liikennöitsijän nimi'] == l]['Autotunnus'])
        @model.Constraint(model.LIIKENNOITSIJAT)
        def kustannusrajoite(m, l): vanha = vanhat_kulut_liikennöitsijä.get(l, 0); return pyo.inequality(vanha * (1 - heitto_kerroin), m.uusi_kustannus_per_liikennöitsijä[l], vanha * (1 + heitto_kerroin))
    else:
        vanhat_kulut_yhteensä = df_autot['Kulut entisellä mallilla'].sum(); uudet_kulut_yhteensä = sum(model.uusi_kustannus_per_auto[a] for a in model.AUTOT)
        model.kustannusrajoite = pyo.Constraint(rule=pyo.inequality(vanhat_kulut_yhteensä * (1 - heitto_kerroin), uudet_kulut_yhteensä, vanhat_kulut_yhteensä * (1 + heitto_kerroin)))
    solver = pyo.SolverFactory('cbc'); solver.options['threads'] = -1; solver.options['ratio'] = params['vaje'] / 100.0
    results = solver.solve(model, tee=False)
    if (results.solver.status == pyo.SolverStatus.ok) and (results.solver.termination_condition == pyo.TerminationCondition.optimal):
        tulokset = [{'Postinumero': p, 'Vyöhyke': v} for p in model.POSTINUMEROT for v in model.VYOHYKKEET if pyo.value(model.y[p,v]) > 0.9]
        df_tulos = pd.DataFrame(tulokset)
        df_vertailu_auto = pd.DataFrame([{'Autotunnus': a, 'Vanha kustannus (€)': vanhat_kulut_dict_auto.get(a, 0), 'Uusi kustannus (€)': pyo.value(model.uusi_kustannus_per_auto[a])} for a in model.AUTOT])
        return "ok", df_tulos, df_vertailu_auto
    else: return "virhe", "Ratkaisua ei löytynyt. Kokeile löysempiä parametreja tai poista lukituksia.", None

def laske_vyohykkeet_automaattisesti(df_keikat, df_pnro, paakeskus_pnro='60100'):
    df_keikat['Postinumero'] = df_keikat['Postinumero'].astype(str)
    df_volyymit = df_keikat.groupby('Postinumero').agg(Kilot_sum=('Kilot', 'sum'), Rahtikirjojen_lkm=('Rahtikirjanumero', 'nunique'), Nippujen_lkm=('Nippunumero', 'nunique')).reset_index()
    df_pnro['Postinumero'] = df_pnro['Postinumero'].astype(str)
    df_pnro_valmis = pd.merge(df_pnro, df_volyymit, on='Postinumero', how='left').fillna({'Rahtikirjojen_lkm': 0, 'Nippujen_lkm': 0, 'Kilot_sum': 0})
    df_pnro_valmis.replace('EILÖYDY', np.nan, inplace=True)
    coords_map = df_pnro_valmis.dropna(subset=['X-Koordinaatti', 'Y-Koordinaatti']).set_index('Postinumero')
    for idx, row in df_pnro_valmis[df_pnro_valmis['X-Koordinaatti'].isna()].iterrows():
        try:
            base_pnro = str(int(row['Postinumero']) // 100 * 100)
            if base_pnro in coords_map.index:
                df_pnro_valmis.loc[idx, 'X-Koordinaatti'] = coords_map.loc[base_pnro, 'X-Koordinaatti']; df_pnro_valmis.loc[idx, 'Y-Koordinaatti'] = coords_map.loc[base_pnro, 'Y-Koordinaatti']
        except (ValueError, TypeError): continue
    df_pnro_valmis.dropna(subset=['X-Koordinaatti', 'Y-Koordinaatti'], inplace=True)
    if df_pnro_valmis.empty: raise ValueError("Koordinaattidataa ei löytynyt tai sitä ei voitu käsitellä.")
    df_pnro_valmis['X-Koordinaatti'] = pd.to_numeric(df_pnro_valmis['X-Koordinaatti']); df_pnro_valmis['Y-Koordinaatti'] = pd.to_numeric(df_pnro_valmis['Y-Koordinaatti'])
    if paakeskus_pnro not in df_pnro_valmis['Postinumero'].values: raise ValueError(f"Pääkeskuksen postinumeroa {paakeskus_pnro} ei löytynyt datasta.")
    paakeskus_coords = df_pnro_valmis[df_pnro_valmis['Postinumero'] == paakeskus_pnro][['X-Koordinaatti', 'Y-Koordinaatti']].iloc[0]
    hub_threshold = df_pnro_valmis['Rahtikirjojen_lkm'].quantile(0.95)
    df_pnro_valmis['Onko_Hub'] = (df_pnro_valmis['Rahtikirjojen_lkm'] >= hub_threshold) & (df_pnro_valmis['Rahtikirjojen_lkm'] >= 5)
    df_pnro_valmis['Etaisyys_paakeskuksesta_km'] = np.sqrt((df_pnro_valmis['X-Koordinaatti'] - paakeskus_coords['X-Koordinaatti'])**2 + (df_pnro_valmis['Y-Koordinaatti'] - paakeskus_coords['Y-Koordinaatti'])**2) / 1000
    df_pnro_valmis['Syrjaisyyspisteet'] = df_pnro_valmis['Etaisyys_paakeskuksesta_km'] / (df_pnro_valmis['Nippujen_lkm'] + 1)
    def maarita_vyohyke(row):
        if row['Postinumero'] == paakeskus_pnro: return 1
        if row['Etaisyys_paakeskuksesta_km'] < 15: return 2
        if 15 <= row['Etaisyys_paakeskuksesta_km'] < 40 and row['Onko_Hub']: return 3
        if row['Etaisyys_paakeskuksesta_km'] >= 40 and row['Onko_Hub']: return 4
        maaseutu = df_pnro_valmis[(df_pnro_valmis['Postinumero'] != paakeskus_pnro) & (df_pnro_valmis['Etaisyys_paakeskuksesta_km'] >= 15) & (~df_pnro_valmis['Onko_Hub'])]
        if maaseutu.empty: return 5
        median_syrjaisyys = maaseutu['Syrjaisyyspisteet'].median()
        if row['Syrjaisyyspisteet'] <= median_syrjaisyys: return 5
        else: return 6
    df_pnro_valmis['Uusi_Vyohyke'] = df_pnro_valmis.apply(maarita_vyohyke, axis=1)
    zone_map = df_pnro_valmis.set_index('Postinumero')['Uusi_Vyohyke']
    def korjaa_postilokerot(row):
        p_str = row['Postinumero']; current_zone = row['Uusi_Vyohyke']
        try:
            p_int = int(p_str)
            if p_int % 10 != 0:
                base_p_str = str(p_int - (p_int % 10))
                if base_p_str in zone_map: return zone_map[base_p_str]
        except (ValueError, TypeError): pass
        return current_zone
    df_pnro_valmis['Uusi_Vyohyke'] = df_pnro_valmis.apply(korjaa_postilokerot, axis=1)
    output_cols = ['Postinumero', 'Uusi_Vyohyke', 'X-Koordinaatti', 'Y-Koordinaatti', 'Rahtikirjojen_lkm']
    return df_pnro_valmis[output_cols].rename(columns={'Uusi_Vyohyke': 'Vyöhyke'})

# =============================================================================
# STREAMLIT-KÄYTTÖLIITTYMÄ
# =============================================================================
st.set_page_config(layout="wide", page_title="Rahtioptimointi")

if 'app_loaded' not in st.session_state:
    st.session_state.app_loaded = True; st.session_state.sheets = {}
    st.session_state.df_tariff_current = pd.DataFrame(); st.session_state.df_zones_current = pd.DataFrame()
    st.session_state.df_autot_current = pd.DataFrame(); st.session_state.vertailu_auto = pd.DataFrame()
    st.session_state.lukitut_tariffit = {}; st.session_state.lukitut_vyohykkeet = {}
    st.session_state.last_error = ""

st.title("🚛 Rahtikustannusten optimointityökalu")

with st.sidebar:
    st.header("1. Data")
    st.download_button("📥 Lataa mallipohja", luo_mallipohja_exceliin(), 'syotetiedot_malli.xlsx')
    uploaded_file = st.file_uploader("Lataa Excel-pohja", type="xlsx")
    if uploaded_file and st.button("Lataa data / Aloita alusta"):
        st.session_state.sheets = pd.read_excel(uploaded_file, sheet_name=None)
        st.session_state.df_tariff_current = st.session_state.sheets['Tariffitaulukko'].copy()
        st.session_state.df_zones_current = st.session_state.sheets['Postinumerot'].copy()
        st.session_state.df_autot_current = st.session_state.sheets['AutojenYhteenveto'].copy()
        st.session_state.vertailu_auto = pd.DataFrame(); st.session_state.lukitut_tariffit = {}; st.session_state.lukitut_vyohykkeet = {}
        st.toast("Data ladattu!", icon="✅"); st.rerun()

    if 'df_tariff_current' not in st.session_state or st.session_state.df_tariff_current.empty:
        st.info("Lataa data Excel-tiedostosta aloittaaksesi."); st.stop()

    st.header("2. Yleiset parametrit")
    tasmaystaso = st.radio("Mihin hintaa täsmätään?", ('Kokonaisuus', 'Liikennöitsijä', 'Auto'), index=0, key="taso_radio")
    sallittu_heitto = st.slider("Sallittu heitto (%)", 0.5, 30.0, 5.0, 0.5, key="heitto_slider")
    
    st.header("3. Toiminnot")
    with st.expander("Vyöhykkeiden määritys", expanded=False):
        vyohyke_tapa = st.radio("Valitse toiminto:", ("Käytä alkuperäisiä (Excel)", "Generoi älykkäästi (Heuristiikka)", "Optimoi matemaattisesti (Hienosäätö)"), key="vyohyke_tapa_radio", index=0)
        if vyohyke_tapa == "Käytä alkuperäisiä (Excel)":
            if st.button("Palauta alkuperäiset vyöhykkeet"):
                st.session_state.df_zones_current = st.session_state.sheets['Postinumerot'].copy()
                st.toast("Alkuperäiset vyöhykkeet palautettu.", icon="↩️"); st.rerun()
        elif vyohyke_tapa == "Generoi älykkäästi (Heuristiikka)":
            paakeskus_pnro = st.text_input("Pääkeskuksen postinumero", "60100", key="paakeskus_input")
            with st.expander("Miten generointi toimii?"): st.markdown("""Tämä toiminto luo uuden vyöhykekartan...""")
            if st.button("Suorita älykäs generointi"):
                with st.spinner("Analysoidaan dataa..."):
                    df_keikat = pd.concat([st.session_state.sheets['Jakokeikat'], st.session_state.sheets['Noutokeikat']], ignore_index=True)
                    tulos = laske_vyohykkeet_automaattisesti(df_keikat, st.session_state.sheets['Postinumerot'], paakeskus_pnro)
                    st.session_state.df_zones_current = tulos
                st.toast("Uusi vyöhykemalli generoitu!", icon="🤖"); st.rerun()
        else:
            sallittu_optimointivaje = st.slider("Sallittu optimointivaje (%)", 0.05, 10.0, 1.0, 0.1, key="vaje_slider")
            with st.expander("Miten optimointi toimii?"): st.markdown("""Tämä toiminto hienosäätää nykyistä vyöhykekarttaa...""")
            if st.button("Suorita matemaattinen optimointi"):
                with st.spinner("Optimoidaan vyöhykkeitä..."):
                    params = {'taso': tasmaystaso, 'heitto': sallittu_heitto, 'vaje': sallittu_optimointivaje, 'lukitut_vyohykkeet': st.session_state.lukitut_vyohykkeet}
                    status, tulos, vertailu = suorita_vyohyke_optimointi(st.session_state.sheets, st.session_state.df_tariff_current, list(st.session_state.df_autot_current['Autotunnus']), params)
                    if status == "ok":
                        original_zones = st.session_state.sheets['Postinumerot'].copy()
                        original_zones.drop(columns='Vyöhyke', inplace=True, errors='ignore'); original_zones['Postinumero'] = original_zones['Postinumero'].astype(str)
                        tulos['Postinumero'] = tulos['Postinumero'].astype(str)
                        new_zones_complete = pd.merge(original_zones, tulos, on='Postinumero', how='left')
                        st.session_state.df_zones_current = new_zones_complete
                        st.session_state.vertailu_auto = vertailu; st.toast("Vyöhykkeet optimoitu!", icon="🎯")
                    else: st.session_state.last_error = tulos
                st.rerun()

    with st.expander("Tariffien laskenta", expanded=True):
        minimi_korotus = st.slider("MINIMIKOROTUS (%)", 0.0, 5.0, 0.1, 0.01, key="min_korotus_slider")
        max_korotus = st.slider("MAKSIMIKOROTUS (%)", 1.0, 20.0, 5.0, 0.1, key="max_korotus_slider")
        if st.button("Laske uudet tariffit", type="primary"):
            with st.spinner("Lasketaan tariffeja..."):
                params = {'taso': tasmaystaso, 'heitto': sallittu_heitto, 'min_korotus': minimi_korotus, 'max_korotus': max_korotus, 'lukitut_tariffit': st.session_state.lukitut_tariffit}
                status, tulos, vertailu = suorita_tariffi_optimointi(st.session_state.sheets, st.session_state.df_zones_current, list(st.session_state.df_autot_current['Autotunnus']), params)
                if status == "ok": st.session_state.df_tariff_current = tulos; st.session_state.vertailu_auto = vertailu; st.toast("Uudet tariffit laskettu!", icon="💰")
                else: st.session_state.last_error = tulos
            st.rerun()

# --- PÄÄNÄYTTÖ ---
if st.session_state.last_error: st.error(f"Tapahtui virhe: {st.session_state.last_error}"); st.session_state.last_error = ""

st.subheader("Nykyinen tariffitaulukko")
edited_tariff = st.data_editor(st.session_state.df_tariff_current, key="tariff_editor", use_container_width=True)
if not edited_tariff.equals(st.session_state.df_tariff_current):
    st.session_state.df_tariff_current = edited_tariff
    st.session_state.lukitut_tariffit = { (r, c): float(row[c]) for r, row in edited_tariff.iterrows() for c in edited_tariff.columns if 'VYÖHYKE' in c and pd.notna(row[c]) }
    st.info("Tariffimuutokset tallennettu. Ne lukitaan seuraavassa laskennassa.")

st.subheader("Nykyinen vyöhykemalli")
col1, col2 = st.columns([0.4, 0.6])
with col1:
    df_zones_display = st.session_state.df_zones_current.copy()
    if 'Postitoimipaikka' not in df_zones_display.columns:
         df_zones_display = pd.merge(df_zones_display, st.session_state.sheets['Postinumerot'][['Postinumero', 'Postitoimipaikka']].astype(str), on='Postinumero', how='left')
    
    if 'Rahtikirjojen_lkm' not in df_zones_display.columns:
        df_keikat_temp = pd.concat([st.session_state.sheets['Jakokeikat'], st.session_state.sheets['Noutokeikat']], ignore_index=True)
        df_keikat_temp['Postinumero'] = df_keikat_temp['Postinumero'].astype(str)
        rk_lkm = df_keikat_temp.groupby('Postinumero')['Rahtikirjanumero'].nunique().reset_index(name='Rahtikirjojen_lkm')
        rk_lkm['Postinumero'] = rk_lkm['Postinumero'].astype(str)
        df_zones_display['Postinumero'] = df_zones_display['Postinumero'].astype(str)
        df_zones_display = pd.merge(df_zones_display, rk_lkm, on='Postinumero', how='left').fillna({'Rahtikirjojen_lkm': 0})
    
    display_cols = ['Postinumero', 'Postitoimipaikka', 'Vyöhyke', 'Rahtikirjojen_lkm']
    edit_cols = ['Postinumero', 'Vyöhyke']
    for col in display_cols:
        if col not in df_zones_display.columns: df_zones_display[col] = np.nan
    df_zones_display['Rahtikirjojen_lkm'] = df_zones_display['Rahtikirjojen_lkm'].astype(int)

    edited_zones = st.data_editor(df_zones_display[display_cols], key="zones_editor", use_container_width=True, height=400, disabled=['Postitoimipaikka', 'Rahtikirjojen_lkm'])
    if not edited_zones[edit_cols].equals(st.session_state.df_zones_current[edit_cols]):
        st.session_state.df_zones_current = pd.merge(st.session_state.df_zones_current.drop(columns='Vyöhyke', errors='ignore'), edited_zones[edit_cols], on='Postinumero', how='left')
        st.session_state.lukitut_vyohykkeet = {row['Postinumero']: row['Vyöhyke'] for _, row in edited_zones.dropna(subset=['Vyöhyke']).iterrows()}
        st.info("Vyöhykemuutokset tallennettu. Ne lukitaan seuraavassa optimoinnissa.")
with col2:
    df_map = st.session_state.df_zones_current.copy()
    df_map.replace('EILÖYDY', np.nan, inplace=True); df_map.dropna(subset=['X-Koordinaatti', 'Y-Koordinaatti', 'Vyöhyke'], inplace=True)
    if not df_map.empty:
        try:
            transformer = Transformer.from_crs("EPSG:3067", "EPSG:4326", always_xy=True)
            df_map['lon'], df_map['lat'] = transformer.transform(df_map['X-Koordinaatti'].values, df_map['Y-Koordinaatti'].values)
            df_map['Vyöhyke'] = df_map['Vyöhyke'].astype(int)
            colors = [[33, 150, 243, 160], [100, 181, 246, 160], [255, 235, 59, 160], [255, 193, 7, 160], [255, 87, 34, 160], [213, 0, 0, 160]]
            df_map['color'] = df_map['Vyöhyke'].apply(lambda z: colors[min(z - 1, len(colors) - 1)])
            st.pydeck_chart(pdk.Deck(
                map_provider="carto", map_style="light",
                initial_view_state=pdk.ViewState(latitude=df_map['lat'].mean(), longitude=df_map['lon'].mean(), zoom=7, pitch=0, bearing=0),
                layers=[pdk.Layer('ScatterplotLayer', data=df_map, get_position='[lon, lat]', get_fill_color='color', get_radius=1500, pickable=True, auto_highlight=True)],
                tooltip={"text": "Postinumero: {Postinumero}\nVyöhyke: {Vyöhyke}"}
            ))
        except Exception as e: st.warning(f"Karttavisualisoinnin luonti epäonnistui: {e}")
    else: st.info("Ei näytettävää dataa kartalla.")

if not st.session_state.vertailu_auto.empty:
    st.header("Laskennan tulokset")
    df_vertailu = st.session_state.vertailu_auto.copy()
    df_orig_autot = st.session_state.sheets['AutojenYhteenveto'][['Autotunnus', 'Liikennöitsijän nimi']]
    df_vertailu = pd.merge(df_vertailu, df_orig_autot, on='Autotunnus', how='left')
    df_vertailu['Erotus (€)'] = df_vertailu['Uusi kustannus (€)'] - df_vertailu['Vanha kustannus (€)']
    df_vertailu['Erotus (%)'] = (df_vertailu['Erotus (€)'] / df_vertailu['Vanha kustannus (€)'] * 100).replace([np.inf, -np.inf], 0).fillna(0)
    st.subheader("Autojen valinta ja vertailu")
    df_vertailu['Mukana'] = df_vertailu['Autotunnus'].isin(list(st.session_state.df_autot_current['Autotunnus']))
    edited_autot = st.data_editor(df_vertailu[['Mukana', 'Autotunnus', 'Liikennöitsijän nimi', 'Vanha kustannus (€)', 'Uusi kustannus (€)', 'Erotus (€)', 'Erotus (%)']], key="autot_editor", hide_index=True, use_container_width=True)
    autot_nyt_mukana = set(edited_autot[edited_autot['Mukana']]['Autotunnus'])
    autot_ennen = set(st.session_state.df_autot_current['Autotunnus'])
    if autot_nyt_mukana != autot_ennen:
        st.session_state.df_autot_current = st.session_state.sheets['AutojenYhteenveto'][st.session_state.sheets['AutojenYhteenveto']['Autotunnus'].isin(autot_nyt_mukana)].copy()
        st.warning("Autojen valinta on muuttunut.")
        if st.button("Päivitä laskelmat muuttuneilla autoilla"):
            st.toast("Autovalinta päivitetty. Aja haluamasi laskenta uudelleen.")
            time.sleep(2); st.rerun()
    df_naytettava = edited_autot[edited_autot['Mukana']]
    if not df_naytettava.empty:
        st.write("**Yhteenvedot (perustuen valittuihin autoihin):**")
        summa_auto = pd.DataFrame(df_naytettava[['Vanha kustannus (€)', 'Uusi kustannus (€)', 'Erotus (€)']].sum()).T; summa_auto['Autotunnus'] = 'YHTEENSÄ'
        st.dataframe(summa_auto.set_index('Autotunnus').style.format("{:,.2f} €"), use_container_width=True)
        st.subheader("Liikennöitsijäkohtainen yhteenveto")
        df_liikenne = df_naytettava.groupby('Liikennöitsijän nimi')[['Vanha kustannus (€)', 'Uusi kustannus (€)', 'Erotus (€)']].sum().reset_index()
        df_liikenne['Erotus (%)'] = (df_liikenne['Erotus (€)'] / df_liikenne['Vanha kustannus (€)'] * 100).replace([np.inf, -np.inf], 0).fillna(0)
        st.dataframe(df_liikenne, hide_index=True, use_container_width=True)
        df_keikat = pd.concat([st.session_state.sheets['Jakokeikat'], st.session_state.sheets['Noutokeikat']], ignore_index=True)
        df_niput_tulokset = pd.merge(_valmistele_data(st.session_state.sheets, list(st.session_state.df_autot_current['Autotunnus']))[1],st.session_state.df_zones_current[['Postinumero', 'Vyöhyke']], on='Postinumero', how='inner')
        df_niput_tulokset['tariffi_rivi_idx'] = df_niput_tulokset['nippu_paino'].apply(lambda p: get_painoluokka_rivi_idx(p, st.session_state.sheets['Tariffitaulukko']))
        df_niput_tulokset.dropna(subset=['tariffi_rivi_idx'], inplace=True)
        df_niput_tulokset['Vyöhyke'] = pd.to_numeric(df_niput_tulokset['Vyöhyke'], errors='coerce').fillna(0).astype(int)
        valid_vyohykkeet = {int(c.split(' ')[1]) for c in st.session_state.df_tariff_current.columns if 'VYÖHYKE' in c}
        df_niput_tulokset = df_niput_tulokset[df_niput_tulokset['Vyöhyke'].isin(valid_vyohykkeet)]
        if not df_niput_tulokset.empty:
            df_niput_tulokset['Uusi_nippu_hinta'] = df_niput_tulokset.apply(
                    lambda r: r['nippu_paino'] * st.session_state.df_tariff_current.at[int(r['tariffi_rivi_idx']), f"VYÖHYKE {r['Vyöhyke']}"] 
                    if st.session_state.sheets['Tariffitaulukko'].at[int(r['tariffi_rivi_idx']), 'Laskentatapa'] == '€/kg' 
                    else st.session_state.df_tariff_current.at[int(r['tariffi_rivi_idx']), f"VYÖHYKE {r['Vyöhyke']}"], axis=1)
        for _, row in edited_autot.iterrows():
            if not row['Mukana']: continue
            with st.expander(f"Analysoi kustannukset: **{row['Autotunnus']}** ({row['Liikennöitsijän nimi']})"):
                data_valinnalle = df_niput_tulokset[df_niput_tulokset['Autotunnus'] == row['Autotunnus']]
                if data_valinnalle.empty: st.warning("Valinnalle ei löytynyt kustannusdataa.")
                else:
                    painoluokka_jarjestys = [get_painoluokka_str(i, st.session_state.sheets['Tariffitaulukko']) for i in st.session_state.sheets['Tariffitaulukko'].index]
                    data_valinnalle['Painoluokka'] = data_valinnalle['tariffi_rivi_idx'].apply(lambda idx: get_painoluokka_str(idx, st.session_state.sheets['Tariffitaulukko']))
                    data_valinnalle['Painoluokka'] = pd.Categorical(data_valinnalle['Painoluokka'], categories=painoluokka_jarjestys, ordered=True)
                    total_cost = data_valinnalle['Uusi_nippu_hinta'].sum()
                    st.metric("Lasketut kokonaiskustannukset", f"{total_cost:,.2f} €", delta=f"{(total_cost - row['Vanha kustannus (€)']):,.2f} €")
                    pivot_table = pd.pivot_table(data_valinnalle, values='Uusi_nippu_hinta', index='Painoluokka', columns='Vyöhyke', aggfunc='sum', fill_value=0)
                    if total_cost > 0:
                        percentage_table = (pivot_table / total_cost * 100)
                        st.write("**Kustannusten jakautuminen (%)**")
                        st.dataframe(percentage_table.style.background_gradient(cmap='Greens', axis=None).format("{:.2f}%"), use_container_width=True)
