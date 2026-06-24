"""
Robô Dashboard BRDRIVE — J&T Express
======================================
1. Lê BRDrive_BI_Novo.xlsx (local)
2. Lê lista oficial de motoristas do Controle BRDRIVE (SharePoint)
3. Calcula KPIs: Utilização % e Pontualidade % por motorista/mês
4. Gera dashboard HTML interativo com visual J&T Express
5. Abre automaticamente no navegador

Como usar:
  python robo_dashboard_brdrive.py
"""

import os, sys, io, json, unicodedata, difflib, webbrowser, requests
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import pandas as pd
from datetime import date
from openpyxl import load_workbook

# ─────────────────────────────────────────────
# CONFIGURAÇÕES
# ─────────────────────────────────────────────

ARQUIVO_BRDRIVE = r"C:\Users\robson.noberto\Desktop\Power BI Gus\BRDrive_BI_Novo.xlsx"
ARQUIVO_CONTROLE_LOCAL = r"C:\Users\robson.noberto\Desktop\Controle Kips Motoristas\Controle_BRDRIVE_atualizado.xlsx"
SHAREPOINT_URL = "https://spjtexpress-my.sharepoint.com/:x:/g/personal/gerson_silva_spjtexpress_onmicrosoft_com/IQADZHQau4f1T7KrFoI4SqRdAV-wnc4VJ_-s2J5BbqNT5qw?e=p0J8tR"

PASTA_SAIDA = r"C:\Users\robson.noberto\Desktop\Controle Kips Motoristas"
ARQUIVO_HTML = os.path.join(PASTA_SAIDA, "Dashboard_KPI_JT.html")

MESES_PT = {1:"Janeiro",2:"Fevereiro",3:"Março",4:"Abril",
            5:"Maio",6:"Junho",7:"Julho",8:"Agosto",
            9:"Setembro",10:"Outubro",11:"Novembro",12:"Dezembro"}

# ─────────────────────────────────────────────
# UTILITÁRIOS
# ─────────────────────────────────────────────

def normalizar(nome):
    s = unicodedata.normalize('NFD', str(nome))
    s = ''.join(c for c in s if unicodedata.category(c) != 'Mn')
    return ' '.join(s.upper().split())

def buscar_match(nome_ctrl, condutores, limite=0.80):
    nc = normalizar(nome_ctrl)
    norm_map = {normalizar(c): c for c in condutores}
    if nc in norm_map: return norm_map[nc]
    matches = difflib.get_close_matches(nc, norm_map.keys(), n=1, cutoff=limite)
    if matches: return norm_map[matches[0]]
    # subconjunto de palavras: "LEANDRO HENRIQUE FREIRE" casa "LEANDRO HENRIQUE RAMALHO GONCALVES FREIRE"
    palavras = set(nc.split())
    melhor_match, melhor_score = None, 0
    for candidato in norm_map:
        palavras_cand = set(candidato.split())
        if palavras.issubset(palavras_cand):
            score = len(palavras) / len(palavras_cand)
            if score > melhor_score:
                melhor_score = score
                melhor_match = candidato
    if melhor_match:
        return norm_map[melhor_match]
    return None

# ─────────────────────────────────────────────
# PASSO 1 — Lista oficial de motoristas
# ─────────────────────────────────────────────

def obter_lista_oficial():
    print(">> Carregando lista de motoristas (OneDrive)...")

    caminho = r"C:\Users\robson.noberto\OneDrive - J&T EXPRESS - FILIAL SP\Controle KPIs Motoristas\Controle BRDRIVE - KPIs 2026.xlsx"

    try:
        xls = pd.ExcelFile(caminho)
        print("Abas encontradas:", xls.sheet_names)

        # tenta achar a aba chinesa
        sheet = '司机打卡控制表' if '司机打卡控制表' in xls.sheet_names else xls.sheet_names[0]

        df = pd.read_excel(caminho, sheet_name=sheet)

        nomes = df.iloc[:,0].dropna().tolist()

        print(f"[OK] {len(nomes)} motoristas carregados da base oficial")

        return [str(n).strip() for n in nomes]

    except Exception as e:
        print(f"[ERRO] Falha ao carregar lista: {e}")
        sys.exit(1)

# ─────────────────────────────────────────────
# PASSO 2 — Calcular KPIs
# ─────────────────────────────────────────────

def calcular_kpis(nomes_oficiais):
    print(">> Calculando KPIs do BRDrive...")
    df = pd.read_excel(ARQUIVO_BRDRIVE, sheet_name='fPrincipal')
    jet = df[df['Transportador'].str.strip().str.upper()=='JET SP'].copy()
    jet['DATA'] = pd.to_datetime(jet['DATA'])

    resultado = []
    for (mes_num, mes_nome), grp_mes in jet.groupby([jet['DATA'].dt.month, jet['MÊS']]):
        condutores_mes = grp_mes['CONDUTOR'].unique().tolist()
        for nome_ctrl in nomes_oficiais:
            match = buscar_match(nome_ctrl, condutores_mes)
            grp = grp_mes[grp_mes['CONDUTOR']==match] if match else pd.DataFrame()
            tot = len(grp)
            if tot > 0:
                on_s = (grp['Tempo Saida OFF'].fillna('OFF').str.strip().str.upper()!='OFF').sum()
                on_c = (grp['Tempo chegada OFF'].fillna('OFF').str.strip().str.upper()!='OFF').sum()
                util = round(((on_s/tot)+(on_c/tot))/2*100,1)
                com_s = grp['PONTUALIDADE SAÍDA'].notna().sum()
                com_c = grp['PONTUALIDADE CHEGADA'].notna().sum()
                np_s = (grp['PONTUALIDADE SAÍDA'].str.strip().str.lower()=='no prazo').sum()
                np_c = (grp['PONTUALIDADE CHEGADA'].str.strip().str.lower()=='no prazo').sum()
                pont = round(((np_s/com_s if com_s>0 else 0)+(np_c/com_c if com_c>0 else 0))/2*100,1) if (com_s>0 or com_c>0) else None
                resultado.append({'mes_num':int(mes_num),'mes':str(mes_nome).upper(),
                                   'condutor':str(nome_ctrl).strip(),'utilizacao':util,
                                   'pontualidade':pont,'viagens':int(tot)})

    hoje = date.today()
    mes_atual = MESES_PT[hoje.month].upper()
    print(f"[OK] KPIs calculados | {len(resultado)} registros | Mes atual: {mes_atual}")
    return resultado, mes_atual

# ─────────────────────────────────────────────
# PASSO 3 — Gerar HTML
# ─────────────────────────────────────────────

COELHO_SVG = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 80 100" width="52" height="65">
  <ellipse cx="24" cy="22" rx="8" ry="18" fill="#fff" stroke="#E8001C" stroke-width="2"/>
  <ellipse cx="24" cy="22" rx="4" ry="13" fill="#ffb3bb"/>
  <ellipse cx="56" cy="22" rx="8" ry="18" fill="#fff" stroke="#E8001C" stroke-width="2"/>
  <ellipse cx="56" cy="22" rx="4" ry="13" fill="#ffb3bb"/>
  <ellipse cx="40" cy="72" rx="22" ry="20" fill="#fff" stroke="#E8001C" stroke-width="2"/>
  <circle cx="40" cy="46" r="18" fill="#fff" stroke="#E8001C" stroke-width="2"/>
  <circle cx="33" cy="43" r="3" fill="#E8001C"/>
  <circle cx="34" cy="42" r="1" fill="#fff"/>
  <circle cx="47" cy="43" r="3" fill="#E8001C"/>
  <circle cx="48" cy="42" r="1" fill="#fff"/>
  <ellipse cx="40" cy="49" rx="3" ry="2" fill="#ffb3bb"/>
  <path d="M37 52 Q40 55 43 52" fill="none" stroke="#E8001C" stroke-width="1.5" stroke-linecap="round"/>
  <ellipse cx="30" cy="50" rx="5" ry="3" fill="#ffdddd" opacity="0.6"/>
  <ellipse cx="50" cy="50" rx="5" ry="3" fill="#ffdddd" opacity="0.6"/>
  <ellipse cx="20" cy="68" rx="6" ry="10" fill="#fff" stroke="#E8001C" stroke-width="1.5" transform="rotate(-20 20 68)"/>
  <ellipse cx="60" cy="68" rx="6" ry="10" fill="#fff" stroke="#E8001C" stroke-width="1.5" transform="rotate(20 60 68)"/>
  <ellipse cx="33" cy="90" rx="7" ry="4" fill="#fff" stroke="#E8001C" stroke-width="1.5"/>
  <ellipse cx="47" cy="90" rx="7" ry="4" fill="#fff" stroke="#E8001C" stroke-width="1.5"/>
  <path d="M37 58 L40 65 L43 58 Z" fill="#E8001C"/>
  <rect x="38" y="56" width="4" height="4" rx="1" fill="#E8001C"/>
</svg>"""

COELHO_MINI = COELHO_SVG.replace('width="52"','width="28"').replace('height="65"','height="35"')

def gerar_html(dados, mes_atual, nomes_oficiais, gerado_em):
    meses_disp = sorted(set(d['mes'] for d in dados),
                        key=lambda m: {'JANEIRO':1,'FEVEREIRO':2,'MARÇO':3,'ABRIL':4,
                                       'MAIO':5,'JUNHO':6,'JULHO':7,'AGOSTO':8,
                                       'SETEMBRO':9,'OUTUBRO':10,'NOVEMBRO':11,'DEZEMBRO':12}.get(m,99))

    pills_html = '<button class="pill" onclick="setMes(\'TODOS\',this)">Todos</button>\n'
    for m in meses_disp:
        ativo = 'active' if m == mes_atual else ''
        abrev = m[:3].capitalize()
        pills_html += f'<button class="pill {ativo}" onclick="setMes(\'{m}\',this)">{abrev}</button>\n'

    data_json = json.dumps(dados, ensure_ascii=False)
    mes_inicial = mes_atual if mes_atual in meses_disp else (meses_disp[-1] if meses_disp else 'TODOS')

    html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>KPI Motoristas — J&T Express</title>
<script src="https://unpkg.com/xlsx/dist/xlsx.full.min.js"></script>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{background:#F0F0F0;color:#1a1a1a;font-family:Arial,sans-serif}}
.header{{background:#E8001C;padding:0 24px;display:flex;align-items:center;justify-content:space-between;height:66px;position:sticky;top:0;z-index:100}}
.header-left{{display:flex;align-items:center;gap:14px}}
.logo-box{{background:#fff;width:42px;height:42px;border-radius:7px;display:flex;flex-direction:column;align-items:center;justify-content:center;font-weight:900;font-size:12px;color:#E8001C;line-height:1.1}}
.logo-title{{font-size:15px;font-weight:700;color:#fff;letter-spacing:.5px}}
.logo-sub{{font-size:10px;color:rgba(255,255,255,.75);margin-top:1px}}
.header-mascot{{display:flex;align-items:flex-end;height:66px}}
.pills{{display:flex;gap:6px;align-items:center;flex-wrap:wrap}}
.pill{{background:rgba(255,255,255,.18);border:1.5px solid rgba(255,255,255,.45);color:#fff;padding:5px 15px;border-radius:20px;font-size:12px;font-weight:700;cursor:pointer;transition:all .15s}}
.pill:hover{{background:rgba(255,255,255,.28)}}
.pill.active{{background:#fff;color:#E8001C}}
.busca{{background:rgba(255,255,255,.18);border:1.5px solid rgba(255,255,255,.45);color:#fff;padding:5px 14px;border-radius:20px;font-size:12px;outline:none;width:185px}}
.busca::placeholder{{color:rgba(255,255,255,.6)}}
.busca:focus{{background:rgba(255,255,255,.26);border-color:#fff}}
.content{{padding:20px 24px;max-width:1280px;margin:0 auto}}
.cards{{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:12px;margin-bottom:18px}}
.kcard{{background:#fff;border-radius:10px;padding:16px 18px;border-bottom:4px solid #E8001C}}
.kcard-label{{font-size:9px;color:#aaa;letter-spacing:1.2px;text-transform:uppercase;margin-bottom:5px;font-weight:700}}
.kcard-val{{font-size:30px;font-weight:900;line-height:1}}
.kcard-sub{{font-size:11px;color:#bbb;margin-top:4px}}
.row2{{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:14px}}
.panel{{background:#fff;border-radius:10px;padding:18px 20px}}
.panel-title{{font-size:10px;font-weight:900;color:#E8001C;letter-spacing:1.4px;text-transform:uppercase;margin-bottom:14px;padding-bottom:9px;border-bottom:2px solid #F0F0F0}}
.chart-evol{{display:flex;gap:10px;align-items:flex-end;height:130px}}
.bar-grp{{flex:1;display:flex;flex-direction:column;align-items:center}}
.bars-inner{{display:flex;gap:5px;align-items:flex-end;height:100px;width:100%;padding-bottom:2px}}
.bar-col{{flex:1;display:flex;flex-direction:column;align-items:center;gap:3px}}
.bar-rect{{width:100%;border-radius:4px 4px 0 0}}
.bar-num{{font-size:10px;font-weight:700}}
.bar-lbl{{font-size:9px;color:#ccc}}
.mes-nome{{font-size:11px;font-weight:700;color:#bbb;margin-top:5px}}
.mes-ativo .mes-nome{{color:#E8001C}}
.mes-ativo .bars-inner{{border-bottom:2px solid #E8001C}}
.legenda{{display:flex;gap:14px;margin-top:10px;padding-top:8px;border-top:1px solid #f5f5f5}}
.leg{{display:flex;align-items:center;gap:5px;font-size:10px;color:#bbb}}
.leg-dot{{width:9px;height:9px;border-radius:2px}}
.dist-wrap{{display:flex;flex-direction:column;gap:14px}}
.dist-item{{display:flex;flex-direction:column;gap:5px}}
.dist-top{{display:flex;justify-content:space-between;align-items:baseline}}
.dist-label{{font-size:12px;font-weight:700}}
.dist-nums{{display:flex;align-items:baseline;gap:4px}}
.dist-count{{font-size:20px;font-weight:900}}
.dist-pct{{font-size:11px;color:#bbb}}
.dist-bg{{background:#F0F0F0;border-radius:6px;height:10px;overflow:hidden}}
.dist-fill{{height:10px;border-radius:6px;transition:width .5s}}
.ranking-wrap{{display:flex;flex-direction:column;gap:5px}}.ranking-wrap.scrollable{{max-height:340px;overflow-y:auto;padding-right:4px}}.ranking-wrap.scrollable::-webkit-scrollbar{{width:5px}}.ranking-wrap.scrollable::-webkit-scrollbar-track{{background:#f0f0f0;border-radius:4px}}.ranking-wrap.scrollable::-webkit-scrollbar-thumb{{background:#E8001C;border-radius:4px}}
.rank-section{{font-size:9px;font-weight:900;letter-spacing:1px;text-transform:uppercase;padding:8px 11px 3px;color:#b35e00;border-top:1px solid #f0f0f0;margin-top:2px}}.rank-section-crit{{color:#c0150a}}
.trend{{font-size:11px;font-weight:900;flex-shrink:0;width:14px;text-align:center}}
.rank-item{{display:flex;align-items:center;gap:10px;padding:7px 11px;border-radius:8px;border:1.5px solid #f0f0f0;background:#fff;transition:all .15s}}
.rank-item:hover{{border-color:#E8001C;background:#fff8f8}}.rank-item.clickable{{cursor:pointer}}
.modal-overlay{{display:none;position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:200;align-items:center;justify-content:center}}
.modal-overlay.open{{display:flex}}
.modal-box{{background:#fff;border-radius:14px;padding:24px 26px;width:440px;max-width:95vw;max-height:85vh;overflow-y:auto}}
.modal-header{{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:16px}}
.modal-nome{{font-size:15px;font-weight:900;color:#1a1a1a;line-height:1.2;flex:1}}
.modal-close{{background:#f0f0f0;border:none;border-radius:50%;width:30px;height:30px;font-size:18px;cursor:pointer;color:#666;flex-shrink:0;margin-left:12px;line-height:1}}
.modal-close:hover{{background:#E8001C;color:#fff}}
.modal-kpi-row{{display:flex;gap:8px;margin-bottom:14px}}
.modal-kpi{{flex:1;background:#f8f8f8;border-radius:8px;padding:10px;text-align:center}}
.modal-kpi-val{{font-size:22px;font-weight:900}}
.modal-kpi-lbl{{font-size:9px;color:#bbb;margin-top:2px;text-transform:uppercase;letter-spacing:.8px}}
.modal-chart{{display:flex;gap:6px;align-items:flex-end;height:90px;margin-bottom:6px}}
.modal-bar-grp{{flex:1;display:flex;flex-direction:column;align-items:center;gap:2px}}
.modal-bar{{width:100%;border-radius:3px 3px 0 0;min-height:3px}}
.modal-bar-num{{font-size:9px;font-weight:700}}
.modal-mes{{font-size:9px;color:#bbb;margin-top:2px}}
.btn-export{{background:rgba(255,255,255,.18);border:1.5px solid rgba(255,255,255,.45);color:#fff;padding:5px 13px;border-radius:20px;font-size:12px;font-weight:700;cursor:pointer;white-space:nowrap}}
.btn-export:hover{{background:rgba(255,255,255,.3)}}
.rank-num{{font-size:13px;font-weight:900;color:#E8001C;width:20px;text-align:center;flex-shrink:0}}
.rank-name{{font-size:11px;color:#333;flex:1;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;font-weight:700}}
.rank-bars{{display:flex;flex-direction:column;gap:3px;width:100px;flex-shrink:0}}
.rank-bar-row{{display:flex;align-items:center;gap:4px}}
.rank-bar-bg{{flex:1;background:#F0F0F0;border-radius:3px;height:5px;overflow:hidden}}
.rank-bar-fill{{height:5px;border-radius:3px}}
.rank-pct{{font-size:10px;font-weight:700;width:32px;text-align:right;flex-shrink:0}}
.badge{{display:inline-block;font-size:10px;font-weight:700;padding:3px 8px;border-radius:10px;flex-shrink:0;min-width:50px;text-align:center}}
.bg{{background:#E8F8EF;color:#1a7a40}}.ba{{background:#FFF3E0;color:#b35e00}}.br{{background:#FDECEA;color:#c0150a}}
.stacked-bar{{display:flex;gap:2px;margin-top:10px;height:10px}}
.stacked-seg{{height:10px;transition:flex .5s;min-width:4px}}
.atualizado{{font-size:10px;color:rgba(255,255,255,.65);white-space:nowrap}}
.footer{{text-align:center;font-size:10px;color:#bbb;padding:18px;display:flex;align-items:center;justify-content:center;gap:10px}}
.kpi-btn{{background:rgba(255,255,255,.1);border:1.5px solid rgba(255,255,255,.3);font-weight:700}}
.kpi-active{{background:#E8001C!important;border-color:#E8001C!important;color:#fff!important}}
@media(max-width:720px){{.row2{{grid-template-columns:1fr}}.pills,.busca,.header-mascot{{display:none}}}}
</style>
</head>
<body>
<div class="header">
  <div class="header-left">
    <div class="logo-box"><span>J&T</span><span style="font-size:7px;letter-spacing:2px">EXPRESS</span></div>
    <div>
      <div class="logo-title">KPI MOTORISTAS — JET SP</div>
      <div class="logo-sub">Utilização · Pontualidade · Dashboard Interativo</div>
    </div>
  </div>
  <div style="display:flex;align-items:center;gap:14px">
    <div class="pills">
      <button class="pill kpi-btn kpi-active" id="btn-util" onclick="setKpi('util',this)">Utilizacao</button>
      <button class="pill kpi-btn" id="btn-pont" onclick="setKpi('pont',this)">Pontualidade</button>
      <span style="width:1px;height:18px;background:rgba(255,255,255,.3);margin:0 4px"></span>
      {pills_html}
      <button class="btn-export" onclick="exportarXLSX()">&#8595; Exportar Excel</button>
      <input class="busca" type="text" placeholder="Buscar motorista..." oninput="atualizar()">
    </div>
    <div style="display:flex;flex-direction:column;align-items:flex-end;gap:2px">
      <span class="atualizado">Atualizado: {gerado_em}</span>
      <span class="atualizado">{len(nomes_oficiais)} motoristas oficiais</span>
    </div>
    <div class="header-mascot">{COELHO_SVG}</div>
  </div>
</div>

<div class="content">
  <div class="cards">
    <div class="kcard"><div class="kcard-label">Motoristas</div><div class="kcard-val" id="c-mot" style="color:#E8001C">—</div><div class="kcard-sub">condutores JET SP</div></div>
    <div class="kcard"><div class="kcard-label">Utilização média</div><div class="kcard-val" id="c-util" style="color:#E8001C">—</div><div class="kcard-sub" id="c-util-sub">—</div></div>
    <div class="kcard"><div class="kcard-label">Pontualidade média</div><div class="kcard-val" id="c-pont" style="color:#E8001C">—</div><div class="kcard-sub" id="c-pont-sub">—</div></div>
    <div class="kcard"><div class="kcard-label">Na meta ≥ 95%</div><div class="kcard-val" id="c-meta" style="color:#1a7a40">—</div><div class="kcard-sub" id="c-meta-pct">—</div></div>
    <div class="kcard"><div class="kcard-label">Critico &lt; 80%</div><div class="kcard-val" id="c-crit" style="color:#c0150a">—</div><div class="kcard-sub" id="c-crit-pct">—</div></div>
    <div class="kcard"><div class="kcard-label">Total viagens</div><div class="kcard-val" id="c-viag" style="color:#1a1a1a">—</div><div class="kcard-sub">no período</div></div>
  </div>
  <div class="row2">
    <div class="panel">
      <div class="panel-title">Evolução mensal</div>
      <div class="chart-evol" id="chart-evol"></div>
      <div class="legenda">
        <div class="leg"><div class="leg-dot" style="background:#E8001C"></div>Utilização</div>
        <div class="leg"><div class="leg-dot" style="background:#FF6B6B"></div>Pontualidade</div>
        <span style="margin-left:auto;font-size:10px;color:#ddd">util ≥ 95% · pont ≥ 95%</span>
      </div>
    </div>
    <div class="panel">
      <div class="panel-title" id="panel-dist-title">Distribuicao — utilizacao % dos motoristas</div>
      <div class="dist-wrap" id="chart-dist"></div>
    </div>
  </div>
  <div class="row2">
    <div class="panel">
      <div class="panel-title" style="color:#1a7a40">Melhores utilização</div>
      <div class="ranking-wrap scrollable" id="rank-top"></div>
    </div>
    <div class="panel">
      <div class="panel-title" style="color:#c0150a">Atenção — precisam melhorar</div>
      <div class="ranking-wrap scrollable" id="rank-bot"></div>
    </div>
  </div>
</div>

<div class="modal-overlay" id="modal-overlay" onclick="if(event.target===this)fecharModal()">
  <div class="modal-box">
    <div class="modal-header">
      <div class="modal-nome" id="modal-nome">—</div>
      <button class="modal-close" onclick="fecharModal()">&#215;</button>
    </div>
    <div class="modal-kpi-row" id="modal-kpis"></div>
    <div class="modal-chart" id="modal-chart"></div>
  </div>
</div>

<div class="footer">{COELHO_MINI}<span>J&T Express · JET SP · {len(nomes_oficiais)} motoristas oficiais · Fonte: BRDrive_BI_Novo.xlsx · Gerado em {gerado_em}</span></div>

<script>
const DADOS={data_json};
const MESES_ORD={{JANEIRO:1,FEVEREIRO:2,MARÇO:3,ABRIL:4,MAIO:5,JUNHO:6,JULHO:7,AGOSTO:8,SETEMBRO:9,OUTUBRO:10,NOVEMBRO:11,DEZEMBRO:12}};
let mesSel='{mes_inicial}';
let kpiSel='util';

function setKpi(kpi,el){{kpiSel=kpi;document.querySelectorAll('.kpi-btn').forEach(b=>b.classList.remove('kpi-active'));el.classList.add('kpi-active');atualizar();}}
function setMes(mes,el){{mesSel=mes;document.querySelectorAll('.pill:not(.kpi-btn)').forEach(b=>b.classList.remove('active'));el.classList.add('active');atualizar();}}
function corUtil(v){{return v>=95?'#1a7a40':v>=80?'#b35e00':'#c0150a'}}
function corPont(v){{return v>=95?'#1a7a40':v>=80?'#b35e00':'#c0150a'}}
function corVal(v){{return corUtil(v);}}
function badge(v,meta,parc){{
  if(v==null) return '<span class="badge" style="background:#f5f5f5;color:#ccc">—</span>';
  if(v>=meta) return `<span class="badge bg">${{v.toFixed(1)}}%</span>`;
  if(v>=parc) return `<span class="badge ba">${{v.toFixed(1)}}%</span>`;
  return `<span class="badge br">${{v.toFixed(1)}}%</span>`;
}}
function agrupar(dados){{
  const mapa={{}};
  dados.forEach(d=>{{
    if(!mapa[d.condutor]) mapa[d.condutor]={{condutor:d.condutor,utils:[],ponts:[],viagens:0}};
    if(d.utilizacao!=null) mapa[d.condutor].utils.push(d.utilizacao);
    if(d.pontualidade!=null) mapa[d.condutor].ponts.push(d.pontualidade);
    mapa[d.condutor].viagens+=d.viagens;
  }});
  return Object.values(mapa).map(m=>{{
    const u=m.utils.length?parseFloat((m.utils.reduce((a,b)=>a+b,0)/m.utils.length).toFixed(1)):null;
    const p=m.ponts.length?parseFloat((m.ponts.reduce((a,b)=>a+b,0)/m.ponts.length).toFixed(1)):null;
    return {{condutor:m.condutor,utilizacao:u,pontualidade:p,viagens:m.viagens}};
  }});
}}
function rankItem(d,i,campo,metaLim,parcLim,prevMap){{
  const v=d[campo];
  const cor=campo==='utilizacao'?corUtil(v||0):corPont(v||0);
  const pv=prevMap&&prevMap[d.condutor]?prevMap[d.condutor][campo]:null;
  const diff=v!=null&&pv!=null?v-pv:null;
  const trendHtml=diff==null
    ?'<span class="trend" style="color:#ddd">·</span>'
    :diff>0.5
      ?`<span class="trend" style="color:#1a7a40" title="vs mes anterior: +${{diff.toFixed(1)}}%">&#9650;</span>`
      :diff<-0.5
        ?`<span class="trend" style="color:#c0150a" title="vs mes anterior: ${{diff.toFixed(1)}}%">&#9660;</span>`
        :'<span class="trend" style="color:#bbb" title="estavel">&#8594;</span>';
  const nm=d.condutor.replace(/"/g,'&quot;');
  return `<div class="rank-item clickable" data-condutor="${{nm}}" onclick="abrirModal(this.dataset.condutor)">
    <div class="rank-num">${{i+1}}</div>
    ${{trendHtml}}
    <div class="rank-name">${{d.condutor}}</div>
    <div class="rank-bars">
      <div class="rank-bar-row"><div class="rank-bar-bg"><div class="rank-bar-fill" style="width:${{v||0}}%;background:${{cor}}"></div></div><div class="rank-pct" style="color:${{cor}}">${{v!=null?v.toFixed(0)+'%':'—'}}</div></div>
    </div>
    ${{badge(v,metaLim,parcLim)}}
  </div>`;
}}
function atualizar(){{
  const busca=document.querySelector('.busca').value.toLowerCase().trim();
  const filtrado=DADOS.filter(d=>{{
    if(mesSel!=='TODOS'&&d.mes!==mesSel) return false;
    if(busca&&!d.condutor.toLowerCase().includes(busca)) return false;
    return true;
  }});
  const ag=agrupar(filtrado);
  const agv=ag.filter(d=>d.utilizacao!=null);
  const tot=agv.length||1;
  const um=agv.length?agv.reduce((s,d)=>s+d.utilizacao,0)/agv.length:0;
  const pm=agv.length?agv.reduce((s,d)=>s+(d.pontualidade||0),0)/agv.length:0;
  const viag=ag.reduce((s,d)=>s+d.viagens,0);

  // define campos conforme KPI selecionado
  const isUtil=kpiSel==='util';
  const metaLim=isUtil?95:95, parcLim=isUtil?80:80;
  const campo=isUtil?'utilizacao':'pontualidade';
  const corFn=isUtil?corUtil:corPont;
  const agvKpi=ag.filter(d=>d[campo]!=null);
  const totKpi=agvKpi.length||1;
  const metaCnt=agvKpi.filter(d=>d[campo]>=metaLim).length;
  const critCnt=agvKpi.filter(d=>d[campo]<parcLim).length;
  const parcCnt=agvKpi.filter(d=>d[campo]>=parcLim&&d[campo]<metaLim).length;

  document.getElementById('c-mot').textContent=ag.length;
  const eu=document.getElementById('c-util'); eu.textContent=um.toFixed(1)+'%'; eu.style.color=corUtil(um);
  document.getElementById('c-util-sub').textContent=um>=95?'Meta atingida':um>=80?'Abaixo da meta':'Critico';
  const ep=document.getElementById('c-pont'); ep.textContent=pm.toFixed(1)+'%'; ep.style.color=corPont(pm);
  document.getElementById('c-pont-sub').textContent=pm>=95?'Meta atingida':pm>=80?'Abaixo da meta':'Critico';
  document.getElementById('c-meta').textContent=metaCnt;
  document.getElementById('c-meta-pct').textContent=Math.round(metaCnt/totKpi*100)+'% dos motoristas';
  document.getElementById('c-crit').textContent=critCnt;
  document.getElementById('c-crit-pct').textContent=Math.round(critCnt/totKpi*100)+'% dos motoristas';
  document.getElementById('c-viag').textContent=viag.toLocaleString('pt-BR');
  document.getElementById('c-meta').previousElementSibling.textContent='Na meta >= '+metaLim+'%';
  document.getElementById('c-crit').previousElementSibling.textContent='Critico < '+parcLim+'%';
  document.getElementById('panel-dist-title').textContent='Distribuicao — '+(isUtil?'utilizacao':'pontualidade')+'% dos motoristas';
  // mapa do mes anterior para setas de tendencia
  const mesesTodos=[...new Set(DADOS.map(d=>d.mes))].sort((a,b)=>MESES_ORD[a]-MESES_ORD[b]);
  let prevMap={{}};
  if(mesSel!=='TODOS'){{
    const idxAtual=mesesTodos.indexOf(mesSel);
    if(idxAtual>0){{
      const mesPrev=mesesTodos[idxAtual-1];
      agrupar(DADOS.filter(d=>d.mes===mesPrev)).forEach(d=>{{prevMap[d.condutor]=d;}});
    }}
  }}

  const meses_disp=mesesTodos;
  document.getElementById('chart-evol').innerHTML=meses_disp.map(mes=>{{
    const md=agrupar(DADOS.filter(d=>d.mes===mes)).filter(d=>d.utilizacao!=null);
    const u=md.length?md.reduce((s,d)=>s+d.utilizacao,0)/md.length:0;
    const p=md.length?md.reduce((s,d)=>s+(d.pontualidade||0),0)/md.length:0;
    const uh=Math.max(4,Math.round(u*.95)),ph=Math.max(4,Math.round(p*.95));
    const ativo=mes===mesSel?'mes-ativo':'',op=mes===mesSel?1:0.3;
    return `<div class="bar-grp ${{ativo}}">
      <div class="bars-inner">
        <div class="bar-col"><div class="bar-num" style="color:#E8001C;opacity:${{op}}">${{u.toFixed(1)}}</div><div class="bar-rect" style="height:${{uh}}px;background:#E8001C;opacity:${{op}}"></div><div class="bar-lbl">Util</div></div>
        <div class="bar-col"><div class="bar-num" style="color:#FF6B6B;opacity:${{op}}">${{p.toFixed(1)}}</div><div class="bar-rect" style="height:${{ph}}px;background:#FF6B6B;opacity:${{op}}"></div><div class="bar-lbl">Pont</div></div>
      </div>
      <div class="mes-nome">${{mes.charAt(0)+mes.slice(1,3).toLowerCase()}}</div>
    </div>`;
  }}).join('');

  document.getElementById('chart-dist').innerHTML=`
    <div class="dist-item"><div class="dist-top"><div class="dist-label" style="color:#1a7a40">Meta atingida >= ${{metaLim}}%</div><div class="dist-nums"><span class="dist-count" style="color:#1a7a40">${{metaCnt}}</span><span class="dist-pct">· ${{Math.round(metaCnt/totKpi*100)}}%</span></div></div><div class="dist-bg"><div class="dist-fill" style="width:${{Math.round(metaCnt/totKpi*100)}}%;background:#27ae60"></div></div></div>
    <div class="dist-item"><div class="dist-top"><div class="dist-label" style="color:#b35e00">Parcial ${{parcLim}}-${{metaLim-1}}%</div><div class="dist-nums"><span class="dist-count" style="color:#b35e00">${{parcCnt}}</span><span class="dist-pct">· ${{Math.round(parcCnt/totKpi*100)}}%</span></div></div><div class="dist-bg"><div class="dist-fill" style="width:${{Math.round(parcCnt/totKpi*100)}}%;background:#e67e22"></div></div></div>
    <div class="dist-item"><div class="dist-top"><div class="dist-label" style="color:#c0150a">Critico < ${{parcLim}}%</div><div class="dist-nums"><span class="dist-count" style="color:#c0150a">${{critCnt}}</span><span class="dist-pct">· ${{Math.round(critCnt/totKpi*100)}}%</span></div></div><div class="dist-bg"><div class="dist-fill" style="width:${{Math.round(critCnt/totKpi*100)}}%;background:#E8001C"></div></div></div>
    <div class="stacked-bar">
      <div class="stacked-seg" style="flex:${{metaCnt||1}};background:#27ae60;border-radius:4px 0 0 4px"></div>
      <div class="stacked-seg" style="flex:${{parcCnt||1}};background:#e67e22"></div>
      <div class="stacked-seg" style="flex:${{critCnt||1}};background:#E8001C;border-radius:0 4px 4px 0"></div>
    </div>
    <div style="display:flex;justify-content:space-between;font-size:9px;margin-top:4px">
      <span style="color:#1a7a40;font-weight:700">${{Math.round(metaCnt/totKpi*100)}}% meta</span>
      <span style="color:#b35e00;font-weight:700">${{Math.round(parcCnt/totKpi*100)}}% parcial</span>
      <span style="color:#c0150a;font-weight:700">${{Math.round(critCnt/totKpi*100)}}% critico</span>
    </div>`;

  const topAll=[...agvKpi].filter(d=>d[campo]>=metaLim).sort((a,b)=>b[campo]-a[campo]);
  const botParc=[...agvKpi].filter(d=>d[campo]>=parcLim&&d[campo]<metaLim).sort((a,b)=>a[campo]-b[campo]);
  const botCrit=[...agvKpi].filter(d=>d[campo]<parcLim).sort((a,b)=>a[campo]-b[campo]);

  document.getElementById('rank-top').closest('.panel').querySelector('.panel-title').textContent='Melhores '+(isUtil?'utilizacao':'pontualidade')+' ('+topAll.length+')';
  document.getElementById('rank-bot').closest('.panel').querySelector('.panel-title').textContent='Atencao — precisam melhorar ('+(botParc.length+botCrit.length)+')';

  document.getElementById('rank-top').innerHTML=topAll.map((d,i)=>rankItem(d,i,campo,metaLim,parcLim,prevMap)).join('');

  let botHtml='';
  if(botParc.length) botHtml+='<div class="rank-section">Parcial '+parcLim+'–'+(metaLim-1)+'%</div>'+botParc.map((d,i)=>rankItem(d,i,campo,metaLim,parcLim,prevMap)).join('');
  if(botCrit.length) botHtml+='<div class="rank-section rank-section-crit">Critico &lt; '+parcLim+'%</div>'+botCrit.map((d,i)=>rankItem(d,botParc.length+i,campo,metaLim,parcLim,prevMap)).join('');
  document.getElementById('rank-bot').innerHTML=botHtml;
}}

function exportarXLSX(){{
  const busca=document.querySelector('.busca').value.toLowerCase().trim();
  const filtrado=DADOS.filter(d=>{{
    if(mesSel!=='TODOS'&&d.mes!==mesSel) return false;
    if(busca&&!d.condutor.toLowerCase().includes(busca)) return false;
    return true;
  }});
  const isUtil=kpiSel==='util';
  const ag=agrupar(filtrado).sort((a,b)=>(b[isUtil?'utilizacao':'pontualidade']||0)-(a[isUtil?'utilizacao':'pontualidade']||0));
  const rows=[['Pos','Motorista','Utilizacao (%)','Pontualidade (%)','Viagens','Status Util','Status Pont']];
  ag.forEach((d,i)=>{{
    const su=d.utilizacao==null?'Sem dados':d.utilizacao>=95?'Meta':d.utilizacao>=80?'Parcial':'Critico';
    const sp=d.pontualidade==null?'Sem dados':d.pontualidade>=95?'Meta':d.pontualidade>=80?'Parcial':'Critico';
    rows.push([i+1, d.condutor, d.utilizacao??'', d.pontualidade??'', d.viagens, su, sp]);
  }});
  const ws=XLSX.utils.aoa_to_sheet(rows);
  ws['!cols']=[{{wch:4}},{{wch:36}},{{wch:14}},{{wch:16}},{{wch:8}},{{wch:12}},{{wch:12}}];
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'KPI Motoristas');
  XLSX.writeFile(wb,'KPI_Motoristas_'+(mesSel==='TODOS'?'Geral':mesSel)+'.xlsx');
}}

function abrirModal(condutor){{
  const dados=DADOS.filter(d=>d.condutor===condutor);
  if(!dados.length) return;
  const mesesM=[...new Set(dados.map(d=>d.mes))].sort((a,b)=>MESES_ORD[a]-MESES_ORD[b]);
  const ag=agrupar(dados)[0]||{{}};
  const uCor=ag.utilizacao!=null?corUtil(ag.utilizacao):'#bbb';
  const pCor=ag.pontualidade!=null?corPont(ag.pontualidade):'#bbb';
  document.getElementById('modal-nome').textContent=condutor;
  document.getElementById('modal-kpis').innerHTML=`
    <div class="modal-kpi"><div class="modal-kpi-val" style="color:${{uCor}}">${{ag.utilizacao!=null?ag.utilizacao.toFixed(1)+'%':'—'}}</div><div class="modal-kpi-lbl">Utilizacao</div></div>
    <div class="modal-kpi"><div class="modal-kpi-val" style="color:${{pCor}}">${{ag.pontualidade!=null?ag.pontualidade.toFixed(1)+'%':'—'}}</div><div class="modal-kpi-lbl">Pontualidade</div></div>
    <div class="modal-kpi"><div class="modal-kpi-val" style="color:#1a1a1a">${{ag.viagens||0}}</div><div class="modal-kpi-lbl">Viagens</div></div>`;
  const isUtil=kpiSel==='util'; const campo=isUtil?'utilizacao':'pontualidade';
  const maxV=Math.max(...mesesM.map(m=>{{const r=dados.find(d=>d.mes===m); return r&&r[campo]!=null?r[campo]:0;}}),1);
  document.getElementById('modal-chart').innerHTML=mesesM.map(m=>{{
    const r=dados.find(d=>d.mes===m);
    const v=r&&r[campo]!=null?r[campo]:null;
    const h=v!=null?Math.max(6,Math.round(v/maxV*80)):0;
    const cor=isUtil?corUtil(v||0):corPont(v||0);
    return `<div class="modal-bar-grp">
      <div class="modal-bar-num" style="color:${{cor}}">${{v!=null?v.toFixed(0)+'%':'—'}}</div>
      <div class="modal-bar" style="height:${{h}}px;background:${{v!=null?cor:'#f0f0f0'}};width:100%"></div>
      <div class="modal-mes">${{m.charAt(0)+m.slice(1,3).toLowerCase()}}</div>
    </div>`;
  }}).join('');
  document.getElementById('modal-overlay').classList.add('open');
}}
function fecharModal(){{document.getElementById('modal-overlay').classList.remove('open');}}
document.addEventListener('keydown',e=>{{if(e.key==='Escape')fecharModal();}});

atualizar();
</script>
<footer style="text-align:center;padding:18px 0 14px;font-size:11px;color:#aaa;letter-spacing:.5px;">
  Desenvolvido por <strong style="color:#E8001C;">Robson D Noberto</strong>
</footer>
</body>
</html>"""
    return html

# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────

def main():
    print()
    print("="*55)
    print("   ROBO DASHBOARD BRDRIVE - JET SP")
    print("="*55)
    print()

    if not os.path.exists(ARQUIVO_BRDRIVE):
        print(f"[ERRO] Arquivo BRDrive nao encontrado:\n   {ARQUIVO_BRDRIVE}")
        sys.exit(1)

    os.makedirs(PASTA_SAIDA, exist_ok=True)

    nomes_oficiais = obter_lista_oficial()
    dados, mes_atual = calcular_kpis(nomes_oficiais)

    hoje = date.today()
    gerado_em = hoje.strftime("%d/%m/%Y")

    print(">> Gerando dashboard HTML...")
    html = gerar_html(dados, mes_atual, nomes_oficiais, gerado_em)

    with open(ARQUIVO_HTML, 'w', encoding='utf-8') as f:
        f.write(html)

    print(f"[OK] Dashboard salvo em:")
    print(f"   {ARQUIVO_HTML}")

    # Copia para o OneDrive para compartilhamento
    import shutil
    usuario = os.environ.get("USERNAME", "robson.noberto")
    onedrive_pasta = None
    for candidato in [
        rf"C:\Users\{usuario}\OneDrive - J&T EXPRESS - FILIAL SP",
        rf"C:\Users\{usuario}\OneDrive - J&T EXPRESS",
        rf"C:\Users\{usuario}\OneDrive",
    ]:
        if os.path.exists(candidato):
            onedrive_pasta = candidato
            break
    if onedrive_pasta:
        destino_onedrive = os.path.join(onedrive_pasta, "Controle KPIs Motoristas", "Dashboard_KPI_JT.html")
        os.makedirs(os.path.dirname(destino_onedrive), exist_ok=True)
        shutil.copy2(ARQUIVO_HTML, destino_onedrive)
        print(f"   {destino_onedrive}  (OneDrive)")
    # Publica no Render via GitHub (index.html é o arquivo rastreado pelo git)
    print(">> Publicando dashboard online (GitHub)...")
    try:
        import subprocess, shutil as _shutil
        _shutil.copy2(ARQUIVO_HTML, os.path.join(PASTA_SAIDA, "index.html"))
        git_cmds = [
            ["git", "-C", PASTA_SAIDA, "add", "index.html"],
            ["git", "-C", PASTA_SAIDA, "commit", "-m", f"Atualização automática {gerado_em}"],
            ["git", "-C", PASTA_SAIDA, "push"],
        ]
        for cmd in git_cmds:
            result = subprocess.run(cmd, capture_output=True, text=True)
            if result.returncode != 0 and "nothing to commit" not in result.stdout:
                print(f"   [aviso git] {result.stderr.strip()}")
        print("   [OK] Dashboard enviado para o ar!")
    except Exception as e:
        print(f"   [aviso] Nao foi possivel publicar online: {e}")

    print()
    print(">> Abrindo no navegador...")
    webbrowser.open(f"file:///{ARQUIVO_HTML.replace(chr(92),'/')}")
    print()
    print("="*55)
    print("[OK] Concluido! Dashboard aberto no navegador.")
    print("="*55)

if __name__ == "__main__":
    main()