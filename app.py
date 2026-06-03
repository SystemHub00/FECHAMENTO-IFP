import os
import io
import re
import traceback
from collections import defaultdict
from datetime import datetime
from flask import Flask, render_template_string, request, jsonify, redirect, Response
from openpyxl import load_workbook

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "ifp-dashboard-secret-2026")

# ═════════════════════════════════════════════════════════════
# METAS
# ═════════════════════════════════════════════════════════════
METAS = {
    "matriculas":       120,
    "ticket_medio":     199.0,
    "financeiro_atual": 0.94,
    "frequencia":       0.75,
    "retencao":         0.94,
}
META_MINIMA_BOM = 3

# ═════════════════════════════════════════════════════════════
# FILTROS JINJA2
# ═════════════════════════════════════════════════════════════
def fmt_brl(value):
    try:
        v = float(value); s = f"{v:,.2f}".replace(",","X").replace(".",",").replace("X",".")
        return f"R$ {s}"
    except Exception: return "R$ 0,00"

def fmt_brl0(value):
    try:
        v = float(value); s = f"{v:,.0f}".replace(",",".")
        return f"R$ {s}"
    except Exception: return "R$ 0"

def fmt_pct(value):
    try: return f"{float(value)*100:.1f}%"
    except Exception: return "0,0%"

def fmt_int(value):
    try: return str(int(float(value)))
    except Exception: return "0"

app.jinja_env.filters["brl"]   = fmt_brl
app.jinja_env.filters["brl0"]  = fmt_brl0
app.jinja_env.filters["pct"]   = fmt_pct
app.jinja_env.filters["toint"] = fmt_int

# ═════════════════════════════════════════════════════════════
# CONVERSÃO TEXTO → FLOAT  /  DATAS
# ═════════════════════════════════════════════════════════════
def to_float(v, default=0.0):
    if v is None: return default
    if isinstance(v, bool): return default
    if isinstance(v, (int, float)):
        f = float(v); return default if (f != f) else f
    s = str(v).strip()
    if not s or s in ("-","—","#N/A","#DIV/0!","#VALOR!","N/A","n/a"): return default
    s = s.replace("R$","").replace("%","").replace("\xa0","").replace(" ","").strip()
    if not s: return default
    if re.search(r'\d\.\d{3},', s): s = s.replace(".","").replace(",",".")
    elif "," in s and "." not in s: s = s.replace(",",".")
    elif "," in s and "." in s: s = s.replace(",","")
    s = re.sub(r"[^\d.\-]","",s)
    if not s or s in ("-","."): return default
    try: return float(s)
    except ValueError: return default

def parse_data(v):
    if v is None: return None
    if isinstance(v, datetime): return v
    s = str(v).strip()
    for fmt in ("%d/%m/%Y %H:%M:%S","%d/%m/%Y","%Y-%m-%d %H:%M:%S","%Y-%m-%d"):
        try: return datetime.strptime(s, fmt)
        except Exception: pass
    return None

def celula_str(v):
    return str(v).strip().lower() if v is not None else ""

# ═════════════════════════════════════════════════════════════
# AGRUPAMENTO DE ABAS (tolerante a truncamento do Excel)
# ═════════════════════════════════════════════════════════════
def tipo_da_aba(nome):
    nl = nome.strip().lower()
    if "(visit" in nl:                  return "visitas"
    if "(matr" in nl:                   return "matricula_quitacao"
    if "(frequ" in nl or "(freq" in nl: return "frequencia"
    if "(hist" in nl:                   return "historico"
    return None

def nome_unidade_da_aba(nome):
    pos = nome.rfind("(")
    return nome[:pos].strip() if pos > 0 else nome.strip()

def agrupar_abas(wb):
    grupos = {}
    for nome in wb.sheetnames:
        tipo = tipo_da_aba(nome)
        if tipo and tipo != "historico":
            unidade = nome_unidade_da_aba(nome)
            grupos.setdefault(unidade, {})[tipo] = wb[nome]
    return grupos

def ler_linhas(ws, max_rows=20000):
    out = []
    for r in ws.iter_rows(values_only=True):
        out.append(list(r))
        if len(out) >= max_rows: break
    return out

def idx_coluna(header, *nomes):
    for j, h in enumerate(header):
        for nome in nomes:
            if nome in h: return j
    return None

def header_lower(rows):
    return [celula_str(c) for c in rows[0]] if rows else []

# ═════════════════════════════════════════════════════════════
# PARSER VISITAS → matrículas
# ═════════════════════════════════════════════════════════════
def parse_visitas(ws):
    res = {"matriculas": 0}
    rows = ler_linhas(ws)
    if len(rows) < 2: return res
    hdr = header_lower(rows)
    ci = idx_coluna(hdr, "status")
    if ci is None: return res
    mat = 0
    for r in rows[1:]:
        if not r or ci >= len(r): continue
        st = celula_str(r[ci])
        if st in ("matrícula","matricula"): mat += 1
    res["matriculas"] = mat
    return res

# ═════════════════════════════════════════════════════════════
# PARSER MATRÍCULA E QUITAÇÃO → financeiro, alunos, ticket, 30/60d
# ═════════════════════════════════════════════════════════════
def parse_matricula_quitacao(ws):
    res = {"ticket_medio":0.0,"fat_comercial":0.0,"fat_total":0.0,
           "fin_atual":0.0,"fin_30":0.0,"fin_60":0.0,
           "valor_atual":0.0,"valor_30":0.0,"valor_60":0.0,
           "valor_spc":0.0,"valor_cancelados":0.0,
           "ativos":0,"cancelados":0,"desistentes":0,"nunca_veio":0,
           "retencao":0.0,"matriculas_mq":0}
    rows = ler_linhas(ws)
    if len(rows) < 2: return res
    hdr = header_lower(rows)

    ci_tipo     = idx_coluna(hdr, "tipo cobrança", "tipo cobranca")
    ci_contrato = idx_coluna(hdr, "contrato")
    ci_status   = idx_coluna(hdr, "status contrato", "status")
    ci_pago     = idx_coluna(hdr, "valor pago")
    ci_valorbase= idx_coluna(hdr, "valor")
    ci_quit     = idx_coluna(hdr, "quitação", "quitacao")
    ci_venc     = idx_coluna(hdr, "vencimento")

    # Contratos únicos por status
    contrato_status = {}
    for r in rows[1:]:
        if not r: continue
        contr = celula_str(r[ci_contrato]) if ci_contrato is not None and ci_contrato < len(r) else ""
        st    = celula_str(r[ci_status])   if ci_status   is not None and ci_status   < len(r) else ""
        if contr: contrato_status[contr] = st
    sc = defaultdict(int)
    for st in contrato_status.values():
        if "ativo" in st or "aprovad" in st or "matrícula" in st or "matricula" in st or "resgate" in st:
            sc["ativo"] += 1
        elif "nunca" in st:   sc["nunca_veio"] += 1
        elif "desist" in st:  sc["desistente"] += 1
        elif "cancel" in st:  sc["cancelado"] += 1
        elif "tranc" in st:   sc["trancado"] += 1
    res["ativos"]=sc["ativo"]; res["cancelados"]=sc["cancelado"]
    res["desistentes"]=sc["desistente"]; res["nunca_veio"]=sc["nunca_veio"]
    base = sc["ativo"]+sc["cancelado"]+sc["desistente"]
    res["retencao"] = (sc["ativo"]/base) if base>0 else 0.0

    # Ticket médio
    tickets = []
    for r in rows[1:]:
        if not r or ci_tipo is None or ci_tipo >= len(r): continue
        if celula_str(r[ci_tipo]) == "matricula":
            v = to_float(r[ci_valorbase]) if ci_valorbase is not None and ci_valorbase < len(r) else 0
            if v>0: tickets.append(v)
    res["ticket_medio"]  = (sum(tickets)/len(tickets)) if tickets else 0.0
    res["matriculas_mq"] = len(tickets)

    # Faturamento total = soma valor pago
    fat = 0.0
    for r in rows[1:]:
        if not r or ci_pago is None or ci_pago >= len(r): continue
        fat += to_float(r[ci_pago])
    res["fat_total"] = fat; res["fat_comercial"] = fat

    # Cobrança Atual = pago / devido ; 30/60d = pago com atraso >= 30/60 dias
    dev=pago=p30=p60=0.0
    for r in rows[1:]:
        if not r or ci_tipo is None or ci_tipo >= len(r): continue
        if celula_str(r[ci_tipo]) != "parcela": continue
        vd = to_float(r[ci_valorbase]) if ci_valorbase is not None and ci_valorbase < len(r) else 0
        vp = to_float(r[ci_pago]) if ci_pago is not None and ci_pago < len(r) else 0
        if vd<=0: continue
        dev += vd; pago += vp
        venc = parse_data(r[ci_venc]) if ci_venc is not None and ci_venc < len(r) else None
        quit = parse_data(r[ci_quit]) if ci_quit is not None and ci_quit < len(r) else None
        if venc and quit:
            atraso = (quit - venc).days
            if atraso >= 30: p30 += vp
            if atraso >= 60: p60 += vp
    res["fin_atual"]   = (pago/dev) if dev>0 else 0.0
    res["valor_atual"] = pago
    res["fin_30"]      = (p30/dev) if dev>0 else 0.0
    res["valor_30"]    = p30
    res["fin_60"]      = (p60/dev) if dev>0 else 0.0
    res["valor_60"]    = p60
    return res

# ═════════════════════════════════════════════════════════════
# PARSER FREQUÊNCIA → presentes/alunos ponderado (com filtro outlier)
# ═════════════════════════════════════════════════════════════
def parse_frequencia(ws):
    rows = ler_linhas(ws)
    if len(rows) < 2: return 0.0
    hdr = header_lower(rows)
    ci_al = idx_coluna(hdr, "alunos")
    ci_pr = idx_coluna(hdr, "presentes", "presente")
    ci_fq = idx_coluna(hdr, "frequência", "frequencia")
    if ci_al is not None and ci_pr is not None:
        tot_al=tot_pr=0.0; n=0
        for r in rows[1:]:
            if not r: continue
            a = to_float(r[ci_al]) if ci_al<len(r) else 0
            p = to_float(r[ci_pr]) if ci_pr<len(r) else 0
            if 0 < a <= 200 and 0 <= p <= a:
                tot_al+=a; tot_pr+=p; n+=1
        if tot_al>0 and n>0:
            return tot_pr/tot_al
    if ci_fq is not None:
        vals=[]
        for r in rows[1:]:
            if not r or ci_fq>=len(r): continue
            v=to_float(r[ci_fq])
            if 0<v<=1.0: vals.append(v)
        if vals: return sum(vals)/len(vals)
    return 0.0

# ═════════════════════════════════════════════════════════════
# PARSE PRINCIPAL
# ═════════════════════════════════════════════════════════════
def parse_sheet_from_wb(wb):
    grupos = agrupar_abas(wb)
    if not grupos:
        raise KeyError(f"Nenhuma aba reconhecida. Encontradas: {wb.sheetnames[:10]}")
    unidades = []
    for nome_unidade, abas in sorted(grupos.items()):
        u = {"nome": nome_unidade,
             "matriculas":0.0,"fat_comercial":0.0,"ticket_medio":0.0,
             "media_diaria":0.0,"fat_total":0.0,
             "fin_atual":0.0,"fin_30":0.0,"fin_60":0.0,
             "valor_atual":0.0,"valor_30":0.0,"valor_60":0.0,
             "valor_spc":0.0,"valor_cancelados":0.0,
             "ativos":0.0,"cancelados":0.0,"desistentes":0.0,"nunca_veio":0.0,
             "m1_v1":0.0,"m1_v2":0.0,"frequencia":0.0,"retencao":0.0}
        if "visitas" in abas:
            u["matriculas"] = parse_visitas(abas["visitas"])["matriculas"]
        if "matricula_quitacao" in abas:
            d = parse_matricula_quitacao(abas["matricula_quitacao"])
            for k in ["ticket_medio","fat_comercial","fat_total","fin_atual","fin_30",
                      "fin_60","valor_atual","valor_30","valor_60","valor_spc",
                      "valor_cancelados","ativos","cancelados","desistentes",
                      "nunca_veio","retencao"]:
                u[k] = d[k]
            if u["matriculas"] == 0 and d["matriculas_mq"] > 0:
                u["matriculas"] = d["matriculas_mq"]
        if "frequencia" in abas:
            f = parse_frequencia(abas["frequencia"])
            if f > 0: u["frequencia"] = f
        unidades.append(calcular_score(u))
    return unidades

# ═════════════════════════════════════════════════════════════
# SCORE + RANKING
# ═════════════════════════════════════════════════════════════
def calcular_score(u):
    score = 0
    if u["matriculas"]   >= METAS["matriculas"]:        score += 1
    if u["ticket_medio"] >= METAS["ticket_medio"]:      score += 1
    if u["fin_atual"]    >= METAS["financeiro_atual"]:  score += 1
    if u["frequencia"]   >= METAS["frequencia"]:        score += 1
    if u["retencao"]     >= METAS["retencao"]:          score += 1
    u["score"] = score
    u["status"] = "bom" if score >= META_MINIMA_BOM else "ruim"

    # Pontuação contínua para ranking (0-100): média ponderada dos indicadores normalizados
    def n(val, meta, cap=2.0):
        return min(val/meta, cap) if meta else 0
    pts = (
        n(u["matriculas"],   METAS["matriculas"])   * 20 +
        n(u["ticket_medio"], METAS["ticket_medio"]) * 20 +
        n(u["fin_atual"],    METAS["financeiro_atual"]) * 20 +
        n(u["frequencia"],   METAS["frequencia"])   * 20 +
        n(u["retencao"],     METAS["retencao"])     * 20
    ) / 2.0  # divide por 2 (cap=2) → escala 0-100
    u["rank_score"] = round(pts, 1)

    def bar(val, mx): return 0 if mx==0 else max(0.0, min(100.0, val/mx*100))
    def cls(val, meta):
        if val >= meta: return "verde"
        if val >= meta*0.85: return "amarelo"
        return "vermelho"
    u["indicators"] = [
        {"label":"Matrículas","display":str(int(u["matriculas"])),"bar":bar(u["matriculas"],200),
         "cls":cls(u["matriculas"],METAS["matriculas"]),"meta":"Meta: 120"},
        {"label":"Ticket Médio","display":fmt_brl0(u["ticket_medio"]),"bar":bar(u["ticket_medio"],400),
         "cls":cls(u["ticket_medio"],METAS["ticket_medio"]),"meta":"Meta: R$ 199"},
        {"label":"Cobrança Atual","display":fmt_pct(u["fin_atual"]),"bar":u["fin_atual"]*100,
         "cls":cls(u["fin_atual"],METAS["financeiro_atual"]),"meta":"Meta: 94%"},
        {"label":"Frequência","display":fmt_pct(u["frequencia"]),"bar":u["frequencia"]*100,
         "cls":cls(u["frequencia"],METAS["frequencia"]),"meta":"Meta: 75%"},
        {"label":"Retenção","display":fmt_pct(u["retencao"]),"bar":u["retencao"]*100,
         "cls":cls(u["retencao"],METAS["retencao"]),"meta":"Meta: 94%"},
    ]
    return u

def ranquear(unidades):
    """Retorna lista ordenada por rank_score (desc) com posição."""
    ordenadas = sorted(unidades, key=lambda u: (u["rank_score"], u["score"], u["fat_total"]), reverse=True)
    for i, u in enumerate(ordenadas):
        u["posicao"] = i + 1
    return ordenadas

def load_data_from_bytes(file_bytes):
    wb = load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    return parse_sheet_from_wb(wb)

def detect_periodo(file_bytes):
    try:
        wb = load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
        for nome in wb.sheetnames[:5]:
            if tipo_da_aba(nome) != "frequencia": continue
            ws = wb[nome]
            for row in ws.iter_rows(min_row=2, max_row=4, values_only=True):
                for cell in row:
                    if cell and "/" in str(cell):
                        val = str(cell).strip()
                        meses=["jan","fev","mar","abr","mai","jun","jul","ago","set","out","nov","dez"]
                        if any(m in val.lower() for m in meses) and len(val)<20:
                            return val
    except Exception: pass
    return ""

# ═════════════════════════════════════════════════════════════
# CACHE
# ═════════════════════════════════════════════════════════════
_cache = {"data":None,"periodo":"","filename":"","wb_bytes":None}
def get_cached_data(): return _cache["data"], _cache["periodo"], _cache["filename"]
def set_cached_data(data, periodo="", filename="", wb_bytes=None):
    _cache.update({"data":data,"periodo":periodo,"filename":filename,"wb_bytes":wb_bytes})

# ═════════════════════════════════════════════════════════════
# DEBUG
# ═════════════════════════════════════════════════════════════
@app.route("/debug")
def debug_index():
    if not _cache["wb_bytes"]: return "<h2>Nenhuma planilha</h2>", 404
    wb = load_workbook(io.BytesIO(_cache["wb_bytes"]), read_only=True, data_only=True)
    grupos = agrupar_abas(wb)
    links = "".join(f'<li><a href="/debug/{k}">{k}</a> — {list(v.keys())}</li>' for k,v in sorted(grupos.items()))
    return f"<h2>Unidades ({len(grupos)})</h2><ul style='font-family:monospace'>{links}</ul>"

@app.route("/debug/<path:unidade_nome>")
def debug_unidade(unidade_nome):
    if not _cache["wb_bytes"]: return "<h2>Nenhuma planilha</h2>", 404
    wb = load_workbook(io.BytesIO(_cache["wb_bytes"]), read_only=True, data_only=True)
    grupos = agrupar_abas(wb)
    found = next((k for k in grupos if unidade_nome.lower() in k.lower()), None)
    if not found: return f"<h2>Não encontrada</h2><p>{list(grupos.keys())}</p>", 404
    html=[f"<h2 style='font-family:monospace'>Debug: {found}</h2>",
          "<style>table{border-collapse:collapse;font-size:10px;font-family:monospace;margin-bottom:20px}"
          "td,th{border:1px solid #ccc;padding:2px 4px;max-width:140px;overflow:hidden}th{background:#1a237e;color:#fff}</style>"]
    for tipo, ws in grupos[found].items():
        rows = ler_linhas(ws, max_rows=6)
        html.append(f"<h3>Aba: {tipo}</h3><table>")
        for i,row in enumerate(rows):
            if not row: continue
            html.append(f"<tr><th>L{i}</th>")
            for cell in row[:28]:
                html.append(f"<td>{'' if cell is None else str(cell)[:24]}</td>")
            html.append("</tr>")
        html.append("</table>")
    return "".join(html)

# ═════════════════════════════════════════════════════════════
# CSS COMPARTILHADO
# ═════════════════════════════════════════════════════════════
CSS = r"""
:root{--red:#c0021c;--red-dk:#8b0014;--red-lt:rgba(192,2,28,.07);--blue:#1a237e;--blue-lt:#3949ab;
--bg:#f2f4f8;--borda:#dde1ef;--txt:#1a1f36;--txt2:#5a6282;--verde:#2e7d32;--verde-lt:#43a047;--verde-bg:#e8f5e9;
--amber:#e65100;--amber-lt:#fb8c00;--amber-bg:#fff3e0;--rose:#c62828;--rose-bg:#ffebee;
--ouro:#d4af37;--prata:#9fa8b5;--bronze:#cd7f32;--r:14px;--shadow:0 2px 16px rgba(26,35,126,.10)}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Inter',sans-serif;background:var(--bg);color:var(--txt);min-height:100vh}
.hdr{background:linear-gradient(120deg,var(--red-dk),var(--red) 60%,#d40020);color:#fff;box-shadow:0 4px 24px rgba(192,2,28,.35)}
.hdr-in{max-width:1600px;margin:0 auto;display:flex;align-items:center;justify-content:space-between;padding:14px 24px;gap:12px;flex-wrap:wrap}
.logo-wrap{display:flex;align-items:center;gap:14px}
.logo-img{width:52px;height:52px;background:#fff;border-radius:10px;display:flex;align-items:center;justify-content:center}
.logo-img span{font-size:1.1rem;font-weight:900;color:var(--red);letter-spacing:-1px}
.logo-txt h1{font-size:1.15rem;font-weight:800;color:#fff;line-height:1.2}
.logo-txt p{font-size:.72rem;opacity:.85;margin-top:2px}
.hdr-badge{background:rgba(255,255,255,.15);border:1.5px solid rgba(255,255,255,.35);border-radius:20px;padding:6px 16px;font-size:.8rem;font-weight:700;color:#fff}
.nav{background:var(--blue);border-bottom:3px solid var(--red);padding:0 24px}
.nav-in{max-width:1600px;margin:0 auto;display:flex;align-items:center;gap:6px;flex-wrap:wrap;padding:0}
.nav a{padding:12px 18px;color:rgba(255,255,255,.7);text-decoration:none;font-size:.86rem;font-weight:700;border-bottom:3px solid transparent;margin-bottom:-3px}
.nav a:hover{color:#fff}
.nav a.ativo{color:#fff;border-bottom-color:#fff}
.toolbar{background:#fff;border-bottom:1px solid var(--borda);padding:10px 24px}
.toolbar-in{max-width:1600px;margin:0 auto;display:flex;align-items:center;gap:12px;flex-wrap:wrap}
.toolbar label{font-size:.82rem;font-weight:700;color:var(--txt2)}
.toolbar input[type=file]{flex:1;min-width:160px;font-size:.82rem;padding:6px 10px;border:1.5px solid var(--borda);border-radius:8px;background:#f8f9fc;color:var(--txt)}
.toolbar input[type=file]::-webkit-file-upload-button{background:var(--red);color:#fff;border:none;border-radius:6px;padding:4px 12px;font-size:.8rem;font-weight:700;cursor:pointer;margin-right:8px}
.btn{border:none;border-radius:8px;font:inherit;font-weight:700;cursor:pointer;padding:8px 18px;font-size:.84rem;text-decoration:none;display:inline-flex;align-items:center;gap:6px}
.btn-red{background:var(--red);color:#fff}.btn-red:hover{background:var(--red-dk)}
.btn-blue{background:var(--blue);color:#fff}.btn-blue:hover{background:var(--blue-lt)}
.smsg{font-size:.8rem;font-weight:600}.smsg.ok{color:#2e7d32}.smsg.err{color:#c62828}
.filtros{background:#fff;border-bottom:1px solid var(--borda);padding:10px 24px}
.filtros-in{max-width:1600px;margin:0 auto;display:flex;align-items:center;gap:10px;flex-wrap:wrap}
.flabel{font-size:.8rem;font-weight:700;color:var(--blue)}
.fbtn{padding:5px 15px;border-radius:20px;border:1.5px solid var(--borda);background:#f4f5f9;color:var(--txt2);font-size:.8rem;font-weight:600;cursor:pointer}
.fbtn:hover{border-color:var(--red);color:var(--red)}.fbtn.ativo{background:var(--blue);color:#fff;border-color:var(--blue)}
.search{padding:6px 14px;border-radius:20px;border:1.5px solid var(--borda);background:#f4f5f9;font-size:.82rem;outline:none;width:210px;font-family:inherit}
.search:focus{border-color:var(--red)}
.totais{max-width:1600px;margin:18px auto 0;padding:0 24px}
.totais-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(178px,1fr));gap:12px}
.tc{background:#fff;border-radius:var(--r);padding:14px 16px;box-shadow:var(--shadow);border-top:4px solid var(--blue);display:flex;flex-direction:column;gap:3px}
.tc-lbl{font-size:.66rem;font-weight:700;color:var(--txt2);text-transform:uppercase;letter-spacing:.06em}
.tc-val{font-size:1.45rem;font-weight:900;color:var(--blue)}.tc-sub{font-size:.68rem;color:var(--txt2)}
.tc.r{border-top-color:var(--red)}.tc.r .tc-val{color:var(--red)}
.tc.g{border-top-color:var(--verde-lt)}.tc.g .tc-val{color:var(--verde)}
.tc.x{border-top-color:#e53935}.tc.x .tc-val{color:#c62828}
.cards{max-width:1600px;margin:18px auto 50px;padding:0 24px}
.cards-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(345px,1fr));gap:14px}
.ucard{background:#fff;border-radius:var(--r);box-shadow:var(--shadow);overflow:hidden;border:1.5px solid var(--borda);transition:.2s}
.ucard:hover{box-shadow:0 8px 30px rgba(192,2,28,.18);transform:translateY(-2px)}
.chd{padding:13px 15px;display:flex;align-items:center;justify-content:space-between;gap:8px}
.chd.bom{background:linear-gradient(135deg,#1b5e20,#2e7d32);color:#fff}
.chd.ruim{background:linear-gradient(135deg,var(--red-dk),var(--red));color:#fff}
.cnome{font-size:.9rem;font-weight:800;flex:1}
.sbadge{background:rgba(255,255,255,.2);border:1px solid rgba(255,255,255,.35);border-radius:20px;padding:3px 10px;font-size:.7rem;font-weight:700;white-space:nowrap}
.cbody{padding:12px 13px;display:flex;flex-direction:column;gap:5px}
.ind{display:flex;align-items:center;gap:8px;padding:5px 8px;border-radius:7px}
.ind.verde{background:var(--verde-bg)}.ind.amarelo{background:var(--amber-bg)}.ind.vermelho{background:var(--rose-bg)}
.ind-lbl{font-size:.73rem;font-weight:600;color:var(--txt2);flex:1}
.ind-meta{font-size:.6rem;color:var(--txt2);opacity:.7}.ind-val{font-size:.83rem;font-weight:800}
.ind.verde .ind-val{color:var(--verde)}.ind.amarelo .ind-val{color:var(--amber)}.ind.vermelho .ind-val{color:var(--rose)}
.barwrap{width:50px;height:5px;background:#e0e4f0;border-radius:3px;overflow:hidden}
.bar{height:100%;border-radius:3px}.bar.verde{background:var(--verde-lt)}.bar.amarelo{background:var(--amber-lt)}.bar.vermelho{background:#e53935}
.sec-title{font-size:.65rem;font-weight:700;color:var(--txt2);text-transform:uppercase;padding:6px 0 2px;border-top:1px solid var(--borda);margin-top:4px;display:flex;align-items:center;gap:5px}
.sec-tag{font-size:.55rem;font-weight:700;padding:1px 6px;border-radius:8px;text-transform:none;letter-spacing:0}
.sec-tag.mq{background:#e8eaf6;color:var(--blue)}
.row3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:5px}.row4{display:grid;grid-template-columns:repeat(4,1fr);gap:5px}
.box{background:#f4f5f9;border-radius:7px;padding:6px 7px;text-align:center}
.box-lbl{font-size:.58rem;font-weight:700;color:var(--txt2);text-transform:uppercase}
.box-val{font-size:.88rem;font-weight:800;color:var(--blue);margin-top:2px}
.box.ok .box-val{color:var(--verde)}.box.al .box-val{color:var(--red)}
.fd3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:5px}
.fd{background:#f4f5f9;border-radius:7px;padding:5px 7px}
.fd-l{font-size:.56rem;font-weight:700;color:var(--txt2);text-transform:uppercase}
.fd-p{font-size:.9rem;font-weight:800;margin-top:1px}.fd-v{font-size:.65rem;color:var(--txt2);margin-top:1px}
.fd.ok .fd-p{color:var(--verde)}.fd.am .fd-p{color:var(--amber)}.fd.bad .fd-p{color:var(--rose)}
.fat-total{display:flex;align-items:center;justify-content:space-between;padding:7px 10px;border-radius:8px;background:linear-gradient(135deg,#e8eaf6,#c5cae9);margin-top:4px}
.fat-lbl{font-size:.73rem;font-weight:700;color:var(--blue)}.fat-val{font-size:.93rem;font-weight:900;color:var(--blue)}
.btn-card-pdf{display:flex;align-items:center;justify-content:center;gap:6px;margin-top:8px;padding:7px 0;border-radius:8px;background:var(--red);color:#fff;font-size:.76rem;font-weight:700;text-decoration:none}
.btn-card-pdf:hover{background:var(--red-dk)}
.landing{max-width:640px;margin:60px auto;padding:0 24px}
.land-card{background:#fff;border-radius:20px;box-shadow:0 8px 40px rgba(26,35,126,.12);overflow:hidden}
.land-top{background:linear-gradient(135deg,var(--red-dk),var(--red));padding:36px 36px 28px;text-align:center}
.land-top .lico{font-size:3.5rem;margin-bottom:10px}.land-top h2{font-size:1.5rem;font-weight:900;color:#fff;margin-bottom:6px}
.land-top p{color:rgba(255,255,255,.88);font-size:.92rem;line-height:1.55}
.land-body{padding:32px 36px}.land-body h3{font-size:1rem;font-weight:800;color:var(--blue);margin-bottom:16px}
.upload-zone{border:2px dashed var(--borda);border-radius:12px;padding:28px 20px;text-align:center;background:#f9fafc}
.upload-zone:hover{border-color:var(--red);background:var(--red-lt)}
.upload-zone input[type=file]{display:none}.upload-zone label{cursor:pointer;display:block}
.uz-icon{font-size:2.4rem;margin-bottom:8px}.uz-title{font-size:.95rem;font-weight:700;color:var(--blue);margin-bottom:4px}
.uz-sub{font-size:.8rem;color:var(--txt2)}.uz-fname{font-size:.82rem;color:var(--red);font-weight:600;margin-top:8px;min-height:18px}
.land-btn{display:block;width:100%;margin-top:18px;padding:13px;border-radius:10px;background:var(--red);color:#fff;border:none;font:inherit;font-size:.95rem;font-weight:800;cursor:pointer}
.land-btn:hover{background:var(--red-dk)}
.land-hints{margin-top:20px;display:flex;flex-direction:column;gap:8px}
.hint{display:flex;align-items:flex-start;gap:10px;padding:10px 12px;border-radius:8px;background:#f4f5f9;font-size:.8rem;color:var(--txt2)}
.hint-icon{font-size:1.1rem}
/* ===== RANKING ===== */
.rank-wrap{max-width:1100px;margin:24px auto 50px;padding:0 24px}
.podium{display:grid;grid-template-columns:1fr 1.15fr 1fr;gap:14px;align-items:end;margin-bottom:28px}
.podium-card{border-radius:16px;padding:18px 14px;text-align:center;color:#fff;box-shadow:0 8px 30px rgba(0,0,0,.18);position:relative}
.podium-card.p1{background:linear-gradient(160deg,#f4cf52,#d4af37);padding-top:28px}
.podium-card.p2{background:linear-gradient(160deg,#c3ccd6,#9fa8b5)}
.podium-card.p3{background:linear-gradient(160deg,#e0a96d,#cd7f32)}
.podium-medal{font-size:2.4rem;line-height:1;margin-bottom:6px}
.podium-pos{font-size:.7rem;font-weight:800;text-transform:uppercase;letter-spacing:.08em;opacity:.9}
.podium-nome{font-size:1.05rem;font-weight:900;margin:4px 0 8px;line-height:1.15}
.podium-score{font-size:1.9rem;font-weight:900}
.podium-score small{font-size:.7rem;font-weight:700;opacity:.85}
.podium-meta{font-size:.72rem;font-weight:700;margin-top:6px;background:rgba(255,255,255,.2);border-radius:10px;padding:3px 8px;display:inline-block}
.rank-table{background:#fff;border-radius:14px;box-shadow:var(--shadow);overflow:hidden}
.rank-row{display:grid;grid-template-columns:54px 1fr repeat(6,minmax(60px,90px)) 80px;align-items:center;gap:8px;padding:11px 16px;border-bottom:1px solid var(--borda);font-size:.82rem}
.rank-row.head{background:var(--blue);color:#fff;font-weight:700;font-size:.7rem;text-transform:uppercase;letter-spacing:.04em;position:sticky;top:0}
.rank-row:last-child{border-bottom:none}
.rank-row:not(.head):hover{background:#f8f9fc}
.rank-pos{font-weight:900;font-size:1rem;color:var(--blue);text-align:center}
.rank-pos.top1{color:#d4af37}.rank-pos.top2{color:#9fa8b5}.rank-pos.top3{color:#cd7f32}
.rank-nome{font-weight:700;display:flex;align-items:center;gap:7px}
.rank-medal{font-size:1.1rem}
.rank-cell{text-align:center;font-weight:700}
.rank-cell.g{color:var(--verde)}.rank-cell.r{color:var(--rose)}
.rank-score-cell{text-align:center;font-weight:900;color:var(--blue);font-size:.95rem}
.rank-badge{font-size:.6rem;font-weight:800;padding:2px 8px;border-radius:10px}
.rank-badge.bom{background:var(--verde-bg);color:var(--verde)}
.rank-badge.ruim{background:var(--rose-bg);color:var(--rose)}
@media(max-width:900px){.rank-row{grid-template-columns:40px 1fr 70px 70px;font-size:.74rem}
.rank-hide{display:none}}
@media(max-width:640px){.cards-grid{grid-template-columns:1fr}.totais-grid{grid-template-columns:repeat(2,1fr)}.row3,.fd3{grid-template-columns:1fr 1fr}.row4{grid-template-columns:repeat(2,1fr)}.podium{grid-template-columns:1fr;gap:10px}.podium-card.p1{order:-1}}
@media print{.hdr,.nav,.toolbar,.filtros,.totais,.btn-card-pdf,.btn,.btn-blue{display:none!important}.cards,.rank-wrap{margin:0;padding:0}.cards-grid{grid-template-columns:repeat(2,1fr)}.ucard{break-inside:avoid;box-shadow:none;border:1px solid #ccc}body{background:#fff}}
"""

# ═════════════════════════════════════════════════════════════
# TEMPLATE DASHBOARD
# ═════════════════════════════════════════════════════════════
TEMPLATE = r"""
<!DOCTYPE html><html lang="pt-br"><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>IFP – Dashboard</title>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap" rel="stylesheet">
<style>__CSS__</style></head><body>
<div class="hdr"><div class="hdr-in">
  <div class="logo-wrap"><div class="logo-img"><span>IFP</span></div>
  <div class="logo-txt"><h1>Dashboard de Fechamento Mensal</h1><p>Instituto de Formação Profissional</p></div></div>
  {% if periodo %}<div class="hdr-badge">📅 {{ periodo }}</div>{% endif %}
</div></div>
<div class="nav"><div class="nav-in">
  <a href="/" class="ativo">📊 Painel</a>
  <a href="/ranking">🏆 Ranking</a>
</div></div>
<div class="toolbar"><div class="toolbar-in">
  <form method="POST" action="/upload" enctype="multipart/form-data" style="display:flex;align-items:center;gap:10px;flex-wrap:wrap">
    <label>📁 Planilha:</label><input type="file" name="planilha" accept=".xlsx,.xlsm">
    <button type="submit" class="btn btn-red">⬆ Carregar</button>
  </form>
  {% if unidades %}<a href="/pdf/todas" class="btn btn-blue" target="_blank">⬇ PDF Geral</a>{% endif %}
  {% if msg %}<span class="smsg {{ 'ok' if msg_ok else 'err' }}">{{ msg }}</span>{% endif %}
</div></div>
{% if unidades %}
<div class="filtros"><div class="filtros-in">
  <span class="flabel">Filtrar:</span>
  <button class="fbtn ativo" onclick="filtrar('todos',this)">Todos ({{ unidades|length }})</button>
  <button class="fbtn" onclick="filtrar('bom',this)" style="color:#2e7d32;border-color:#a5d6a7">✅ Bom ({{ unidades|selectattr('status','eq','bom')|list|length }})</button>
  <button class="fbtn" onclick="filtrar('ruim',this)" style="color:#c62828;border-color:#ef9a9a">❌ Ruim ({{ unidades|selectattr('status','eq','ruim')|list|length }})</button>
  <input type="text" class="search" placeholder="🔍 Buscar unidade..." oninput="buscar(this.value)">
</div></div>
<div class="totais"><div class="totais-grid">
  <div class="tc"><span class="tc-lbl">Matrículas Totais</span><span class="tc-val">{{ totais.matriculas|toint }}</span><span class="tc-sub">{{ unidades|length }} unidades</span></div>
  <div class="tc r"><span class="tc-lbl">Alunos Ativos</span><span class="tc-val">{{ totais.ativos|toint }}</span><span class="tc-sub">Retenção: {{ totais.retencao_str }}</span></div>
  <div class="tc"><span class="tc-lbl">Faturamento Total</span><span class="tc-val">{{ totais.fat_total|brl0 }}</span><span class="tc-sub">Ticket médio: {{ totais.ticket_str }}</span></div>
  <div class="tc"><span class="tc-lbl">Cobr. Atual Média</span><span class="tc-val">{{ totais.fin_atual_str }}</span><span class="tc-sub">Meta: 94%</span></div>
  <div class="tc"><span class="tc-lbl">Frequência Média</span><span class="tc-val">{{ totais.freq_str }}</span><span class="tc-sub">Meta: 75%</span></div>
  <div class="tc g"><span class="tc-lbl">✅ Unidades Boas</span><span class="tc-val">{{ unidades|selectattr('status','eq','bom')|list|length }}</span><span class="tc-sub">≥ 3 metas</span></div>
  <div class="tc x"><span class="tc-lbl">❌ Unidades Ruins</span><span class="tc-val">{{ unidades|selectattr('status','eq','ruim')|list|length }}</span><span class="tc-sub">menos de 3 metas</span></div>
</div></div>
<div class="cards"><div class="cards-grid" id="cards-grid">
{% for u in unidades %}
<div class="ucard" data-status="{{ u.status }}" data-nome="{{ u.nome|lower }}">
  <div class="chd {{ u.status }}">
    <span class="cnome">{{ u.nome }}</span>
    <span class="sbadge">{%- if u.status=='bom' %}✅ BOM{%- else %}❌ RUIM{%- endif -%}&nbsp;· {{ u.score }}/5</span>
  </div>
  <div class="cbody">
    {% for ind in u.indicators %}
    <div class="ind {{ ind.cls }}"><span class="ind-lbl">{{ ind.label }}</span><span class="ind-meta">{{ ind.meta }}</span>
      <span class="ind-val">{{ ind.display }}</span>
      <div class="barwrap"><div class="bar {{ ind.cls }}" style="width:{{ ind.bar|round(1) }}%"></div></div></div>
    {% endfor %}
    <span class="sec-title">📋 Cobrança <span class="sec-tag mq">Matrícula e Quitação</span></span>
    <div class="fd3">
      {% set fa = 'ok' if u.fin_atual>=0.94 else ('am' if u.fin_atual>=0.80 else 'bad') %}
      <div class="fd {{ fa }}"><div class="fd-l">Atual</div><div class="fd-p">{{ u.fin_atual|pct }}</div><div class="fd-v">{{ u.valor_atual|brl0 }}</div></div>
      <div class="fd"><div class="fd-l">30 dias+</div><div class="fd-p">{{ u.fin_30|pct }}</div><div class="fd-v">{{ u.valor_30|brl0 }}</div></div>
      <div class="fd"><div class="fd-l">60 dias+</div><div class="fd-p">{{ u.fin_60|pct }}</div><div class="fd-v">{{ u.valor_60|brl0 }}</div></div>
    </div>
    <span class="sec-title">💰 Faturamento <span class="sec-tag mq">Matrícula e Quitação</span></span>
    <div class="row3">
      <div class="box"><div class="box-lbl">Comercial</div><div class="box-val">{{ u.fat_comercial|brl0 }}</div></div>
      <div class="box"><div class="box-lbl">Ticket Médio</div><div class="box-val">{{ u.ticket_medio|brl0 }}</div></div>
      <div class="box"><div class="box-lbl">Matrículas</div><div class="box-val">{{ u.matriculas|toint }}</div></div>
    </div>
    <span class="sec-title">👥 Alunos <span class="sec-tag mq">Matrícula e Quitação</span></span>
    <div class="row4">
      <div class="box ok"><div class="box-lbl">Ativos</div><div class="box-val">{{ u.ativos|toint }}</div></div>
      <div class="box {{ 'al' if u.cancelados>20 else '' }}"><div class="box-lbl">Cancelados</div><div class="box-val">{{ u.cancelados|toint }}</div></div>
      <div class="box {{ 'al' if u.desistentes>50 else '' }}"><div class="box-lbl">Desistentes</div><div class="box-val">{{ u.desistentes|toint }}</div></div>
      <div class="box {{ 'al' if u.nunca_veio>30 else '' }}"><div class="box-lbl">Nunca Veio</div><div class="box-val">{{ u.nunca_veio|toint }}</div></div>
    </div>
    <div class="fat-total"><span class="fat-lbl">Faturamento Total (carteira)</span><span class="fat-val">{{ u.fat_total|brl }}</span></div>
    <a href="/pdf/unidade/{{ loop.index0 }}" target="_blank" class="btn-card-pdf">⬇ Baixar PDF</a>
  </div>
</div>
{% endfor %}
</div></div>
{% else %}
<div class="landing"><div class="land-card">
  <div class="land-top"><div class="lico">📊</div><h2>Dashboard IFP</h2>
    <p>Envie a planilha de fechamento mensal para visualizar os indicadores das unidades.</p></div>
  <div class="land-body"><h3>Carregar planilha</h3>
    <form method="POST" action="/upload" enctype="multipart/form-data">
      <div class="upload-zone" id="upload-zone">
        <label for="file-input"><div class="uz-icon">📂</div><div class="uz-title">Clique para selecionar</div>
          <div class="uz-sub">ou arraste e solte aqui</div><div class="uz-fname" id="uz-fname">Nenhum arquivo selecionado</div></label>
        <input type="file" id="file-input" name="planilha" accept=".xlsx,.xlsm"
               onchange="document.getElementById('uz-fname').textContent=this.files[0]?.name||'Nenhum arquivo'">
      </div>
      <button type="submit" class="land-btn">⬆ Carregar Planilha</button>
      {% if msg %}<p style="margin-top:12px;text-align:center;font-size:.85rem;font-weight:600;color:{{ '#c62828' if not msg_ok else '#2e7d32' }}">{{ msg }}</p>{% endif %}
    </form>
    <div class="land-hints">
      <div class="hint"><span class="hint-icon">📋</span><span>Abas: <strong>IFP - Unidade (Visitas)</strong>, <strong>(Matrícula e Quitação)</strong>, <strong>(Frequência)</strong></span></div>
      <div class="hint"><span class="hint-icon">🏆</span><span>Após carregar, veja o <strong>Ranking</strong> das unidades</span></div>
    </div>
  </div>
</div></div>
{% endif %}
<script>
var fa='todos';
function filtrar(s,b){fa=s;document.querySelectorAll('.fbtn').forEach(function(x){x.classList.remove('ativo')});if(b)b.classList.add('ativo');ap(document.querySelector('.search')?document.querySelector('.search').value.toLowerCase():'')}
function buscar(q){ap(q.toLowerCase())}
function ap(q){q=q||'';document.querySelectorAll('.ucard').forEach(function(c){c.style.display=((fa==='todos'||c.dataset.status===fa)&&(!q||c.dataset.nome.includes(q)))?'':'none'})}
var z=document.getElementById('upload-zone');
if(z){z.addEventListener('dragover',function(e){e.preventDefault();z.style.borderColor='var(--red)'});
z.addEventListener('dragleave',function(){z.style.borderColor=''});
z.addEventListener('drop',function(e){e.preventDefault();z.style.borderColor='';var f=e.dataTransfer.files[0];if(f){var i=document.getElementById('file-input');var d=new DataTransfer();d.items.add(f);i.files=d.files;document.getElementById('uz-fname').textContent=f.name}})}
</script>
</body></html>
"""

# ═════════════════════════════════════════════════════════════
# TEMPLATE RANKING
# ═════════════════════════════════════════════════════════════
RANKING_TEMPLATE = r"""
<!DOCTYPE html><html lang="pt-br"><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>IFP – Ranking</title>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap" rel="stylesheet">
<style>__CSS__</style></head><body>
<div class="hdr"><div class="hdr-in">
  <div class="logo-wrap"><div class="logo-img"><span>IFP</span></div>
  <div class="logo-txt"><h1>Ranking de Unidades</h1><p>Instituto de Formação Profissional</p></div></div>
  {% if periodo %}<div class="hdr-badge">📅 {{ periodo }}</div>{% endif %}
</div></div>
<div class="nav"><div class="nav-in">
  <a href="/">📊 Painel</a>
  <a href="/ranking" class="ativo">🏆 Ranking</a>
</div></div>
{% if ranking %}
<div class="toolbar"><div class="toolbar-in">
  <span class="flabel">🏆 Classificação geral das {{ ranking|length }} unidades por desempenho</span>
  <div style="flex:1"></div>
  <a href="/pdf/ranking" class="btn btn-red" target="_blank">⬇ Baixar Ranking em PDF</a>
</div></div>
<div class="rank-wrap">
  <!-- PÓDIO -->
  <div class="podium">
    {% if ranking|length > 1 %}
    <div class="podium-card p2">
      <div class="podium-medal">🥈</div>
      <div class="podium-pos">2º Lugar</div>
      <div class="podium-nome">{{ ranking[1].nome }}</div>
      <div class="podium-score">{{ ranking[1].rank_score }}<small>/100</small></div>
      <div class="podium-meta">{{ ranking[1].score }}/5 metas</div>
    </div>
    {% else %}<div></div>{% endif %}
    <div class="podium-card p1">
      <div class="podium-medal">🥇</div>
      <div class="podium-pos">1º Lugar</div>
      <div class="podium-nome">{{ ranking[0].nome }}</div>
      <div class="podium-score">{{ ranking[0].rank_score }}<small>/100</small></div>
      <div class="podium-meta">{{ ranking[0].score }}/5 metas</div>
    </div>
    {% if ranking|length > 2 %}
    <div class="podium-card p3">
      <div class="podium-medal">🥉</div>
      <div class="podium-pos">3º Lugar</div>
      <div class="podium-nome">{{ ranking[2].nome }}</div>
      <div class="podium-score">{{ ranking[2].rank_score }}<small>/100</small></div>
      <div class="podium-meta">{{ ranking[2].score }}/5 metas</div>
    </div>
    {% else %}<div></div>{% endif %}
  </div>
  <!-- TABELA COMPLETA -->
  <div class="rank-table">
    <div class="rank-row head">
      <div>#</div><div>Unidade</div>
      <div class="rank-hide">Matríc.</div><div class="rank-hide">Ticket</div>
      <div class="rank-hide">Cobr.</div><div class="rank-hide">Freq.</div>
      <div class="rank-hide">Retenç.</div><div>Pontos</div><div class="rank-hide">Status</div>
    </div>
    {% for u in ranking %}
    <div class="rank-row">
      <div class="rank-pos {% if u.posicao==1 %}top1{% elif u.posicao==2 %}top2{% elif u.posicao==3 %}top3{% endif %}">
        {% if u.posicao==1 %}🥇{% elif u.posicao==2 %}🥈{% elif u.posicao==3 %}🥉{% else %}{{ u.posicao }}{% endif %}
      </div>
      <div class="rank-nome">{{ u.nome }}</div>
      <div class="rank-cell rank-hide {{ 'g' if u.matriculas>=120 else 'r' }}">{{ u.matriculas|toint }}</div>
      <div class="rank-cell rank-hide {{ 'g' if u.ticket_medio>=199 else 'r' }}">{{ u.ticket_medio|brl0 }}</div>
      <div class="rank-cell rank-hide {{ 'g' if u.fin_atual>=0.94 else 'r' }}">{{ u.fin_atual|pct }}</div>
      <div class="rank-cell rank-hide {{ 'g' if u.frequencia>=0.75 else 'r' }}">{{ u.frequencia|pct }}</div>
      <div class="rank-cell rank-hide {{ 'g' if u.retencao>=0.94 else 'r' }}">{{ u.retencao|pct }}</div>
      <div class="rank-score-cell">{{ u.rank_score }}</div>
      <div class="rank-hide"><span class="rank-badge {{ u.status }}">{{ 'BOM' if u.status=='bom' else 'RUIM' }}</span></div>
    </div>
    {% endfor %}
  </div>
</div>
{% else %}
<div class="landing"><div class="land-card"><div class="land-top"><div class="lico">🏆</div>
  <h2>Sem dados para ranquear</h2><p>Carregue uma planilha no Painel para ver o ranking das unidades.</p></div>
  <div class="land-body"><a href="/" class="land-btn" style="text-align:center;text-decoration:none">⬅ Ir para o Painel</a></div></div></div>
{% endif %}
</body></html>
"""

# ═════════════════════════════════════════════════════════════
# TEMPLATE PDF (cards)
# ═════════════════════════════════════════════════════════════
PDF_TEMPLATE = r"""
<!DOCTYPE html><html lang="pt-br"><head><meta charset="UTF-8"><title>IFP – {{ titulo }}</title>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800;900&display=swap" rel="stylesheet">
<style>
*{box-sizing:border-box;margin:0;padding:0}body{font-family:'Inter',sans-serif;background:#fff;color:#1a1f36;font-size:11px}
.ph{background:linear-gradient(135deg,#8b0014,#c0021c);color:#fff;padding:16px 22px;display:flex;align-items:center;justify-content:space-between;margin-bottom:16px}
.ph-logo{font-size:1.3rem;font-weight:900;letter-spacing:-1px}.ph-r h2{font-size:.95rem;font-weight:800;text-align:right}.ph-r p{font-size:.68rem;opacity:.85;text-align:right;margin-top:2px}
.no-print{text-align:center;margin-bottom:12px}.no-print button{border:none;border-radius:8px;padding:9px 24px;font-size:.88rem;font-weight:700;cursor:pointer;margin:0 4px}
.btn-p{background:#c0021c;color:#fff}.btn-c{background:#1a237e;color:#fff}
.grid{display:grid;grid-template-columns:repeat(2,1fr);gap:11px;padding:0 16px 16px}
.ucard{border:1.5px solid #dde1ef;border-radius:10px;overflow:hidden;break-inside:avoid}
.chd{padding:10px 12px;display:flex;align-items:center;justify-content:space-between}
.chd.bom{background:linear-gradient(135deg,#1b5e20,#2e7d32);color:#fff}.chd.ruim{background:linear-gradient(135deg,#8b0014,#c0021c);color:#fff}
.cnome{font-size:.85rem;font-weight:800}.sbadge{font-size:.66rem;font-weight:700;background:rgba(255,255,255,.2);border:1px solid rgba(255,255,255,.3);border-radius:12px;padding:2px 8px}
.cbody{padding:9px 11px;display:flex;flex-direction:column;gap:5px}
.ind{display:flex;align-items:center;gap:6px;padding:4px 7px;border-radius:6px}
.ind.verde{background:#e8f5e9}.ind.amarelo{background:#fff3e0}.ind.vermelho{background:#ffebee}
.ind-lbl{font-size:.68rem;font-weight:600;color:#5a6282;flex:1}.ind-meta{font-size:.56rem;color:#5a6282;opacity:.7}.ind-val{font-size:.78rem;font-weight:800}
.ind.verde .ind-val{color:#2e7d32}.ind.amarelo .ind-val{color:#e65100}.ind.vermelho .ind-val{color:#c62828}
.barwrap{width:42px;height:4px;background:#e0e4f0;border-radius:2px;overflow:hidden}.bar{height:100%;border-radius:2px}
.bar.verde{background:#43a047}.bar.amarelo{background:#fb8c00}.bar.vermelho{background:#e53935}
.sec{font-size:.58rem;font-weight:700;color:#5a6282;text-transform:uppercase;padding:4px 0 2px;border-top:1px solid #eee;margin-top:2px}
.r3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:4px}.r4{display:grid;grid-template-columns:repeat(4,1fr);gap:4px}
.bx{background:#f4f5f9;border-radius:5px;padding:4px 6px;text-align:center}.bx-l{font-size:.54rem;font-weight:700;color:#5a6282;text-transform:uppercase}.bx-v{font-size:.82rem;font-weight:800;color:#1a237e;margin-top:1px}
.bx.ok .bx-v{color:#2e7d32}.bx.al .bx-v{color:#c0021c}
.fat{display:flex;align-items:center;justify-content:space-between;padding:5px 7px;border-radius:6px;background:#e8eaf6;margin-top:3px}.fat-l{font-size:.68rem;font-weight:700;color:#1a237e}.fat-v{font-size:.85rem;font-weight:900;color:#1a237e}
.rod{text-align:center;color:#999;font-size:.64rem;padding:8px 0 14px;border-top:1px solid #eee;margin:0 16px}
@media print{body{-webkit-print-color-adjust:exact;print-color-adjust:exact}.no-print{display:none}}
</style></head><body>
<div class="ph"><div class="ph-logo">IFP</div><div class="ph-r"><h2>{{ titulo }}</h2><p>Instituto de Formação Profissional &nbsp;|&nbsp; {{ periodo }}</p></div></div>
<div class="no-print"><button class="btn-p" onclick="window.print()">🖨️ Imprimir / Salvar PDF</button><button class="btn-c" onclick="window.close()">✕ Fechar</button></div>
<div class="grid">
{% for u in unidades %}
<div class="ucard"><div class="chd {{ u.status }}"><span class="cnome">{{ u.nome }}</span><span class="sbadge">{%- if u.status=='bom' %}✅ BOM{%- else %}❌ RUIM{%- endif -%}&nbsp;· {{ u.score }}/5</span></div>
<div class="cbody">
{% for ind in u.indicators %}<div class="ind {{ ind.cls }}"><span class="ind-lbl">{{ ind.label }}</span><span class="ind-meta">{{ ind.meta }}</span><span class="ind-val">{{ ind.display }}</span><div class="barwrap"><div class="bar {{ ind.cls }}" style="width:{{ ind.bar|round(1) }}%"></div></div></div>{% endfor %}
<span class="sec">Cobrança</span><div class="r3"><div class="bx"><div class="bx-l">Atual</div><div class="bx-v">{{ u.fin_atual|pct }}</div></div><div class="bx"><div class="bx-l">30d+</div><div class="bx-v">{{ u.fin_30|pct }}</div></div><div class="bx"><div class="bx-l">60d+</div><div class="bx-v">{{ u.fin_60|pct }}</div></div></div>
<span class="sec">Faturamento</span><div class="r3"><div class="bx"><div class="bx-l">Comercial</div><div class="bx-v">{{ u.fat_comercial|brl0 }}</div></div><div class="bx"><div class="bx-l">Ticket</div><div class="bx-v">{{ u.ticket_medio|brl0 }}</div></div><div class="bx"><div class="bx-l">Matrículas</div><div class="bx-v">{{ u.matriculas|toint }}</div></div></div>
<span class="sec">Alunos</span><div class="r4"><div class="bx ok"><div class="bx-l">Ativos</div><div class="bx-v">{{ u.ativos|toint }}</div></div><div class="bx {{ 'al' if u.cancelados>20 else '' }}"><div class="bx-l">Cancel.</div><div class="bx-v">{{ u.cancelados|toint }}</div></div><div class="bx {{ 'al' if u.desistentes>50 else '' }}"><div class="bx-l">Desist.</div><div class="bx-v">{{ u.desistentes|toint }}</div></div><div class="bx {{ 'al' if u.nunca_veio>30 else '' }}"><div class="bx-l">N.Veio</div><div class="bx-v">{{ u.nunca_veio|toint }}</div></div></div>
<div class="fat"><span class="fat-l">Faturamento Total</span><span class="fat-v">{{ u.fat_total|brl }}</span></div>
</div></div>
{% endfor %}
</div>
<div class="rod">Gerado pelo sistema IFP Dashboard &nbsp;|&nbsp; {{ periodo }}</div>
{% if auto_print %}<script>window.addEventListener('load',function(){setTimeout(function(){window.print()},800)});</script>{% endif %}
</body></html>
"""

# ═════════════════════════════════════════════════════════════
# TEMPLATE PDF RANKING
# ═════════════════════════════════════════════════════════════
PDF_RANKING = r"""
<!DOCTYPE html><html lang="pt-br"><head><meta charset="UTF-8"><title>IFP – Ranking</title>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800;900&display=swap" rel="stylesheet">
<style>
*{box-sizing:border-box;margin:0;padding:0}body{font-family:'Inter',sans-serif;background:#fff;color:#1a1f36;font-size:11px}
.ph{background:linear-gradient(135deg,#8b0014,#c0021c);color:#fff;padding:16px 22px;display:flex;align-items:center;justify-content:space-between;margin-bottom:16px}
.ph-logo{font-size:1.3rem;font-weight:900}.ph-r h2{font-size:.95rem;font-weight:800;text-align:right}.ph-r p{font-size:.68rem;opacity:.85;text-align:right;margin-top:2px}
.no-print{text-align:center;margin-bottom:12px}.no-print button{border:none;border-radius:8px;padding:9px 24px;font-size:.88rem;font-weight:700;cursor:pointer;margin:0 4px}
.btn-p{background:#c0021c;color:#fff}.btn-c{background:#1a237e;color:#fff}
.podium{display:grid;grid-template-columns:1fr 1.15fr 1fr;gap:12px;align-items:end;max-width:760px;margin:0 auto 20px;padding:0 16px}
.pc{border-radius:14px;padding:14px 10px;text-align:center;color:#fff}
.pc.p1{background:linear-gradient(160deg,#f4cf52,#d4af37);padding-top:22px}
.pc.p2{background:linear-gradient(160deg,#c3ccd6,#9fa8b5)}
.pc.p3{background:linear-gradient(160deg,#e0a96d,#cd7f32)}
.pc .m{font-size:2rem}.pc .pos{font-size:.65rem;font-weight:800;text-transform:uppercase}
.pc .nm{font-size:.95rem;font-weight:900;margin:3px 0 6px}.pc .sc{font-size:1.5rem;font-weight:900}
table{border-collapse:collapse;width:calc(100% - 32px);margin:0 16px 16px;font-size:10.5px}
th,td{border:1px solid #e0e4f0;padding:6px 8px;text-align:center}
th{background:#1a237e;color:#fff;font-size:.62rem;text-transform:uppercase}
td.nm{text-align:left;font-weight:700}td.pos{font-weight:900;color:#1a237e}
td.g{color:#2e7d32;font-weight:700}td.r{color:#c62828;font-weight:700}
td.sc{font-weight:900;color:#1a237e}
tr:nth-child(even){background:#f8f9fc}
.rod{text-align:center;color:#999;font-size:.64rem;padding:8px 0 14px;border-top:1px solid #eee;margin:0 16px}
@media print{body{-webkit-print-color-adjust:exact;print-color-adjust:exact}.no-print{display:none}}
</style></head><body>
<div class="ph"><div class="ph-logo">IFP</div><div class="ph-r"><h2>Ranking de Unidades</h2><p>Instituto de Formação Profissional &nbsp;|&nbsp; {{ periodo }}</p></div></div>
<div class="no-print"><button class="btn-p" onclick="window.print()">🖨️ Imprimir / Salvar PDF</button><button class="btn-c" onclick="window.close()">✕ Fechar</button></div>
<div class="podium">
  {% if ranking|length>1 %}<div class="pc p2"><div class="m">🥈</div><div class="pos">2º Lugar</div><div class="nm">{{ ranking[1].nome }}</div><div class="sc">{{ ranking[1].rank_score }}</div></div>{% else %}<div></div>{% endif %}
  <div class="pc p1"><div class="m">🥇</div><div class="pos">1º Lugar</div><div class="nm">{{ ranking[0].nome }}</div><div class="sc">{{ ranking[0].rank_score }}</div></div>
  {% if ranking|length>2 %}<div class="pc p3"><div class="m">🥉</div><div class="pos">3º Lugar</div><div class="nm">{{ ranking[2].nome }}</div><div class="sc">{{ ranking[2].rank_score }}</div></div>{% else %}<div></div>{% endif %}
</div>
<table>
<tr><th>#</th><th style="text-align:left">Unidade</th><th>Matríc.</th><th>Ticket</th><th>Cobr.</th><th>Freq.</th><th>Retenç.</th><th>Pontos</th><th>Status</th></tr>
{% for u in ranking %}
<tr>
  <td class="pos">{% if u.posicao==1 %}🥇{% elif u.posicao==2 %}🥈{% elif u.posicao==3 %}🥉{% else %}{{ u.posicao }}{% endif %}</td>
  <td class="nm">{{ u.nome }}</td>
  <td class="{{ 'g' if u.matriculas>=120 else 'r' }}">{{ u.matriculas|toint }}</td>
  <td class="{{ 'g' if u.ticket_medio>=199 else 'r' }}">{{ u.ticket_medio|brl0 }}</td>
  <td class="{{ 'g' if u.fin_atual>=0.94 else 'r' }}">{{ u.fin_atual|pct }}</td>
  <td class="{{ 'g' if u.frequencia>=0.75 else 'r' }}">{{ u.frequencia|pct }}</td>
  <td class="{{ 'g' if u.retencao>=0.94 else 'r' }}">{{ u.retencao|pct }}</td>
  <td class="sc">{{ u.rank_score }}</td>
  <td class="{{ 'g' if u.status=='bom' else 'r' }}">{{ 'BOM' if u.status=='bom' else 'RUIM' }}</td>
</tr>
{% endfor %}
</table>
<div class="rod">Gerado pelo sistema IFP Dashboard &nbsp;|&nbsp; {{ periodo }}</div>
{% if auto_print %}<script>window.addEventListener('load',function(){setTimeout(function(){window.print()},800)});</script>{% endif %}
</body></html>
"""

# ═════════════════════════════════════════════════════════════
# ROTAS
# ═════════════════════════════════════════════════════════════
def _calc_totais(unidades):
    total_mat   = sum(u["matriculas"] for u in unidades)
    total_fat_t = sum(u["fat_total"]  for u in unidades)
    total_ativ  = sum(u["ativos"]     for u in unidades)
    ret_list  = [u["retencao"]   for u in unidades if u["retencao"]   > 0]
    freq_list = [u["frequencia"] for u in unidades if u["frequencia"] > 0]
    fin_list  = [u["fin_atual"]  for u in unidades if u["fin_atual"]  > 0]
    tk_list   = [u["ticket_medio"] for u in unidades if u["ticket_medio"] > 0]
    return {
        "matriculas": total_mat, "ativos": total_ativ, "fat_total": total_fat_t,
        "ticket_str": fmt_brl0(sum(tk_list)/len(tk_list) if tk_list else 0),
        "retencao_str": f"{(sum(ret_list)/len(ret_list) if ret_list else 0)*100:.1f}%",
        "freq_str": f"{(sum(freq_list)/len(freq_list) if freq_list else 0)*100:.1f}%",
        "fin_atual_str": f"{(sum(fin_list)/len(fin_list) if fin_list else 0)*100:.1f}%",
    }

@app.route("/", methods=["GET"])
def index():
    data, periodo, filename = get_cached_data()
    unidades = data or []
    totais = _calc_totais(unidades) if unidades else {}
    html = TEMPLATE.replace("__CSS__", CSS)
    return render_template_string(html, unidades=unidades, totais=totais, periodo=periodo,
        msg=request.args.get("msg",""), msg_ok=request.args.get("ok","1")=="1")

@app.route("/ranking", methods=["GET"])
def ranking():
    data, periodo, _ = get_cached_data()
    rk = ranquear(list(data)) if data else []
    html = RANKING_TEMPLATE.replace("__CSS__", CSS)
    return render_template_string(html, ranking=rk, periodo=periodo)

@app.route("/upload", methods=["POST"])
def upload():
    if "planilha" not in request.files:
        return redirect("/?msg=Nenhum+arquivo+enviado&ok=0")
    f = request.files["planilha"]
    if not f or not f.filename:
        return redirect("/?msg=Nenhum+arquivo+selecionado&ok=0")
    if not f.filename.lower().endswith((".xlsx",".xlsm")):
        return redirect("/?msg=Formato+inválido&ok=0")
    try:
        fb = f.read()
        unidades = load_data_from_bytes(fb)
        periodo  = detect_periodo(fb) or f.filename
        set_cached_data(unidades, periodo, f.filename, wb_bytes=fb)
        return redirect(f"/?msg=✅+{len(unidades)}+unidades+carregadas&ok=1")
    except KeyError as e:
        return redirect(f"/?msg=Erro:+{str(e)[:120]}&ok=0")
    except Exception as e:
        traceback.print_exc()
        return redirect(f"/?msg=Erro:+{str(e)[:120]}&ok=0")

@app.route("/pdf/todas")
def pdf_todas():
    data, periodo, _ = get_cached_data()
    if not data: return redirect("/?msg=Sem+dados&ok=0")
    return Response(render_template_string(PDF_TEMPLATE, unidades=data,
        titulo="Relatório Geral — Todas as Unidades", periodo=periodo or "—", auto_print=False),
        mimetype="text/html")

@app.route("/pdf/unidade/<int:idx>")
def pdf_unidade(idx):
    data, periodo, _ = get_cached_data()
    if not data or idx >= len(data): return redirect("/?msg=Não+encontrado&ok=0")
    u = data[idx]
    return Response(render_template_string(PDF_TEMPLATE, unidades=[u],
        titulo=f"Relatório — {u['nome']}", periodo=periodo or "—", auto_print=True),
        mimetype="text/html")

@app.route("/pdf/ranking")
def pdf_ranking():
    data, periodo, _ = get_cached_data()
    if not data: return redirect("/?msg=Sem+dados&ok=0")
    rk = ranquear(list(data))
    return Response(render_template_string(PDF_RANKING, ranking=rk,
        periodo=periodo or "—", auto_print=True), mimetype="text/html")

@app.route("/api/dados")
def api_dados():
    data, periodo, filename = get_cached_data()
    return jsonify({"periodo":periodo,"filename":filename,"total":len(data or []),"unidades":data or []})

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
