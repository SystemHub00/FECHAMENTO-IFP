import os
import io
import traceback
from flask import Flask, render_template_string, request, jsonify, redirect, Response
from openpyxl import load_workbook

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "ifp-dashboard-secret-2026")

# ─────────────────────────────────────────────────────────────
# METAS
# ─────────────────────────────────────────────────────────────
METAS = {
    "matriculas":       120,
    "ticket_medio":     199.0,
    "financeiro_atual": 0.94,
    "frequencia":       0.75,
    "retencao":         0.94,
}

# ─────────────────────────────────────────────────────────────
# FILTROS JINJA2
# ─────────────────────────────────────────────────────────────
def fmt_brl(value):
    try:
        v = float(value)
        s = f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        return f"R$ {s}"
    except Exception:
        return "R$ 0,00"

def fmt_brl0(value):
    try:
        v = float(value)
        s = f"{v:,.0f}".replace(",", ".")
        return f"R$ {s}"
    except Exception:
        return "R$ 0"

def fmt_pct(value):
    try:
        return f"{float(value) * 100:.1f}%"
    except Exception:
        return "0,0%"

def fmt_int(value):
    try:
        return str(int(float(value)))
    except Exception:
        return "0"

app.jinja_env.filters["brl"]   = fmt_brl
app.jinja_env.filters["brl0"]  = fmt_brl0
app.jinja_env.filters["pct"]   = fmt_pct
app.jinja_env.filters["toint"] = fmt_int

# ─────────────────────────────────────────────────────────────
# LEITURA DA PLANILHA
# Lógica: cada aba tem nome  "IFP - Unidade (Tipo)"
# Tipos: Visitas | Matrícula e Quitação | Frequência | Histórico
# ─────────────────────────────────────────────────────────────

SUFIXOS = {
    "visitas":              "(Visitas)",
    "matricula_quitacao":   "(Matrícula e Quitação)",
    "frequencia":           "(Frequência)",
    "historico":            "(Histórico)",
}

def safe_float(v, default=0.0):
    try:
        f = float(v)
        return f if f == f else default
    except (TypeError, ValueError):
        return default

def safe_int(v, default=0):
    try:
        return int(float(v))
    except (TypeError, ValueError):
        return default


def agrupar_abas(wb):
    """
    Recebe a workbook e agrupa as abas por unidade.
    Retorna dict: { "IFP - Águas Lindas": {"visitas": ws, "frequencia": ws, ...}, ... }
    """
    grupos = {}
    for nome in wb.sheetnames:
        nome_strip = nome.strip()
        for chave, sufixo in SUFIXOS.items():
            sufixo_lower = sufixo.lower()
            nome_lower   = nome_strip.lower()
            if nome_lower.endswith(sufixo_lower):
                unidade = nome_strip[: -len(sufixo)].strip()
                if unidade not in grupos:
                    grupos[unidade] = {}
                grupos[unidade][chave] = wb[nome]
                break
    return grupos


def parse_visitas(ws):
    """Extrai dados da aba Visitas."""
    rows = list(ws.iter_rows(max_row=200, values_only=True))
    dados = {"matriculas": 0.0, "fat_comercial": 0.0, "ticket_medio": 0.0,
             "media_diaria": 0.0, "fat_total": 0.0}
    if not rows:
        return dados

    # Procura linha de cabeçalho
    header_idx = None
    col_map = {}
    for i, row in enumerate(rows[:15]):
        if not row:
            continue
        cells = [str(c).strip().lower() if c is not None else "" for c in row]
        if any("matr" in c for c in cells) or any("unidade" in c for c in cells):
            header_idx = i
            for j, c in enumerate(cells):
                if "matr" in c:          col_map["matriculas"] = j
                if "fat" in c and "com" in c: col_map["fat_comercial"] = j
                if "ticket" in c:        col_map["ticket_medio"] = j
                if "media" in c or "média" in c: col_map["media_diaria"] = j
                if "fat" in c and "total" in c:  col_map["fat_total"] = j
            break

    # Se não encontrou cabeçalho, tenta ler primeira linha numérica
    if header_idx is None:
        for row in rows:
            if row and row[0] is not None:
                try:
                    dados["matriculas"] = safe_float(row[0])
                    if len(row) > 1: dados["fat_comercial"] = safe_float(row[1])
                    if len(row) > 2: dados["ticket_medio"]  = safe_float(row[2])
                    break
                except Exception:
                    pass
        return dados

    # Lê a ÚLTIMA linha com dados (totais/médias geralmente na última)
    for row in reversed(rows[header_idx + 1:]):
        if not row or all(c is None for c in row):
            continue
        if col_map.get("matriculas") is not None:
            v = safe_float(row[col_map["matriculas"]] if col_map["matriculas"] < len(row) else None)
            if v > 0:
                dados["matriculas"] = v
                for k, idx in col_map.items():
                    if idx < len(row):
                        dados[k] = safe_float(row[idx])
                break
    return dados


def parse_frequencia(ws):
    """Extrai frequência média da aba Frequência."""
    rows = list(ws.iter_rows(max_row=200, values_only=True))
    frequencias = []
    for row in rows:
        if not row:
            continue
        for cell in row:
            if cell is None:
                continue
            v = safe_float(cell)
            # Valores entre 0 e 1 são percentuais já normalizados
            if 0 < v <= 1:
                frequencias.append(v)
            # Valores entre 1 e 100 são percentuais em formato inteiro
            elif 1 < v <= 100:
                frequencias.append(v / 100)
    if frequencias:
        return sum(frequencias) / len(frequencias)
    return 0.0


def parse_matricula_quitacao(ws):
    """Extrai fin_atual, fin_30, fin_60, ativos, cancelados, etc."""
    rows = list(ws.iter_rows(max_row=200, values_only=True))
    resultado = {
        "fin_atual": 0.0, "fin_30": 0.0, "fin_60": 0.0,
        "valor_atual": 0.0, "valor_30": 0.0, "valor_60": 0.0,
        "valor_spc": 0.0, "valor_cancelados": 0.0,
        "ativos": 0.0, "cancelados": 0.0, "desistentes": 0.0,
        "nunca_veio": 0.0, "m1_v1": 0.0, "m1_v2": 0.0,
    }
    if not rows:
        return resultado

    header_idx = None
    col_map = {}
    for i, row in enumerate(rows[:15]):
        if not row:
            continue
        cells = [str(c).strip().lower() if c is not None else "" for c in row]
        if any(k in " ".join(cells) for k in ["atual", "ativo", "cobr", "retenç"]):
            header_idx = i
            for j, c in enumerate(cells):
                if "atual" in c and "fin" not in c and "%" not in c:    col_map["fin_atual"] = j
                elif "30" in c and ("dia" in c or "%" in c):            col_map["fin_30"] = j
                elif "60" in c and ("dia" in c or "%" in c):            col_map["fin_60"] = j
                elif "ativo" in c:      col_map["ativos"] = j
                elif "cancel" in c:    col_map["cancelados"] = j
                elif "desist" in c:    col_map["desistentes"] = j
                elif "nunca" in c:     col_map["nunca_veio"] = j
                elif "m1" in c and "v1" in c: col_map["m1_v1"] = j
                elif "m1" in c and "v2" in c: col_map["m1_v2"] = j
                elif "spc" in c:       col_map["valor_spc"] = j
            break

    if header_idx is None:
        return resultado

    for row in reversed(rows[header_idx + 1:]):
        if not row or all(c is None for c in row):
            continue
        has_data = False
        for k, idx in col_map.items():
            if idx < len(row) and row[idx] is not None:
                v = safe_float(row[idx])
                if v != 0:
                    has_data = True
                resultado[k] = v
        if has_data:
            break

    # Normaliza percentuais que vieram como inteiro (ex: 94 → 0.94)
    for k in ("fin_atual", "fin_30", "fin_60", "m1_v1", "m1_v2"):
        if resultado[k] > 1:
            resultado[k] = resultado[k] / 100
    return resultado


def parse_historico(ws):
    """Extrai retenção da aba Histórico."""
    rows = list(ws.iter_rows(max_row=200, values_only=True))
    retencoes = []
    for row in rows:
        if not row:
            continue
        for cell in row:
            v = safe_float(cell)
            if 0 < v <= 1:
                retencoes.append(v)
            elif 1 < v <= 100:
                retencoes.append(v / 100)
    if retencoes:
        return sum(retencoes) / len(retencoes)
    return 0.0


def parse_sheet_from_wb(wb):
    grupos = agrupar_abas(wb)

    if not grupos:
        nomes = wb.sheetnames
        raise KeyError(
            f"Nenhuma aba reconhecida foi encontrada. "
            f"As abas precisam ter nomes no formato: "
            f"'IFP - Unidade (Visitas)', 'IFP - Unidade (Frequência)', etc. "
            f"Abas encontradas: {nomes[:8]}"
        )

    unidades = []
    for nome_unidade, abas in sorted(grupos.items()):
        u = {
            "nome": nome_unidade,
            "matriculas": 0.0, "fat_comercial": 0.0, "ticket_medio": 0.0,
            "media_diaria": 0.0, "fat_total": 0.0,
            "fin_atual": 0.0, "fin_30": 0.0, "fin_60": 0.0,
            "valor_atual": 0.0, "valor_30": 0.0, "valor_60": 0.0,
            "valor_spc": 0.0, "valor_cancelados": 0.0,
            "ativos": 0.0, "cancelados": 0.0, "desistentes": 0.0,
            "nunca_veio": 0.0, "m1_v1": 0.0, "m1_v2": 0.0,
            "frequencia": 0.0, "retencao": 0.0,
        }

        if "visitas" in abas:
            dados_vis = parse_visitas(abas["visitas"])
            u.update(dados_vis)

        if "matricula_quitacao" in abas:
            dados_mq = parse_matricula_quitacao(abas["matricula_quitacao"])
            u.update(dados_mq)

        if "frequencia" in abas:
            u["frequencia"] = parse_frequencia(abas["frequencia"])

        if "historico" in abas:
            u["retencao"] = parse_historico(abas["historico"])

        unidades.append(calcular_score(u))

    return unidades


def calcular_score(u):
    score = 0
    if u["matriculas"]   >= METAS["matriculas"]:       score += 1
    if u["ticket_medio"] >= METAS["ticket_medio"]:     score += 1
    if u["fin_atual"]    >= METAS["financeiro_atual"]:  score += 1
    if u["frequencia"]   >= METAS["frequencia"]:       score += 1
    if u["retencao"]     >= METAS["retencao"]:         score += 1
    u["score"] = score

    if score >= 2:
        u["status"] = "bom"
    elif score == 1:
        u["status"] = "medio"
    else:
        u["status"] = "ruim"

    def pct_bar(val, max_val):
        if max_val == 0: return 0
        return max(0.0, min(100.0, val / max_val * 100))

    def ind_cls(val, meta):
        if val >= meta:              return "verde"
        if val >= meta * 0.85:       return "amarelo"
        return "vermelho"

    u["indicators"] = [
        {"label": "Matrículas",     "display": str(int(u["matriculas"])),
         "bar": pct_bar(u["matriculas"], 200),
         "cls": ind_cls(u["matriculas"], METAS["matriculas"]),   "meta": "Meta: 120"},
        {"label": "Ticket Médio",   "display": fmt_brl0(u["ticket_medio"]),
         "bar": pct_bar(u["ticket_medio"], 400),
         "cls": ind_cls(u["ticket_medio"], METAS["ticket_medio"]), "meta": "Meta: R$ 199"},
        {"label": "Cobrança Atual", "display": fmt_pct(u["fin_atual"]),
         "bar": u["fin_atual"] * 100,
         "cls": ind_cls(u["fin_atual"], METAS["financeiro_atual"]), "meta": "Meta: 94%"},
        {"label": "Frequência",     "display": fmt_pct(u["frequencia"]),
         "bar": u["frequencia"] * 100,
         "cls": ind_cls(u["frequencia"], METAS["frequencia"]),   "meta": "Meta: 75%"},
        {"label": "Retenção",       "display": fmt_pct(u["retencao"]),
         "bar": u["retencao"] * 100,
         "cls": ind_cls(u["retencao"], METAS["retencao"]),       "meta": "Meta: 94%"},
    ]
    return u


def load_data_from_bytes(file_bytes):
    wb = load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    return parse_sheet_from_wb(wb)


def detect_periodo(file_bytes):
    try:
        wb = load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
        for nome in wb.sheetnames:
            ws = wb[nome]
            for row in ws.iter_rows(max_row=5, values_only=True):
                for cell in row:
                    if cell:
                        val = str(cell).strip()
                        meses = ["jan","fev","mar","abr","mai","jun","jul","ago","set","out","nov","dez"]
                        anos  = ["2024","2025","2026","2027"]
                        val_lower = val.lower()
                        if (any(m in val_lower for m in meses) or any(a in val for a in anos)) and len(val) < 30:
                            return val
    except Exception:
        pass
    return ""


# ─────────────────────────────────────────────────────────────
# CACHE EM MEMÓRIA
# ─────────────────────────────────────────────────────────────
_cache = {"data": None, "periodo": "", "filename": ""}

def get_cached_data():
    return _cache["data"], _cache["periodo"], _cache["filename"]

def set_cached_data(data, periodo="", filename=""):
    _cache["data"]     = data
    _cache["periodo"]  = periodo
    _cache["filename"] = filename


# ─────────────────────────────────────────────────────────────
# TEMPLATE HTML — cores IFP (vermelho #c0021c, azul #1a237e, branco)
# ─────────────────────────────────────────────────────────────
TEMPLATE = """
<!DOCTYPE html>
<html lang="pt-br">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>IFP – Dashboard Fechamento Mensal</title>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap" rel="stylesheet">
<style>
:root{
  --ifp-red:#c0021c; --ifp-red-dk:#8b0014; --ifp-red-lt:#e8002099;
  --ifp-blue:#1a237e; --ifp-blue-md:#283593; --ifp-blue-lt:#3949ab;
  --white:#ffffff;
  --cinza:#f4f5f9; --borda:#dde1ef;
  --txt:#1a1f36; --txt2:#5a6282;
  --verde-dk:#1b5e20; --verde:#2e7d32; --verde-lt:#43a047;
  --laranja:#e65100; --laranja-lt:#fb8c00;
  --r:12px;
  --shadow:0 2px 12px rgba(26,35,126,0.10);
  --shadow-hover:0 8px 28px rgba(192,2,28,0.18);
}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Inter',sans-serif;background:var(--cinza);color:var(--txt);min-height:100vh}

/* ── HEADER ── */
.hdr{
  background:linear-gradient(135deg, var(--ifp-red-dk) 0%, var(--ifp-red) 55%, #d4001f 100%);
  color:#fff;padding:0 24px;
  box-shadow:0 4px 20px rgba(192,2,28,0.35);
}
.hdr-in{max-width:1600px;margin:0 auto;display:flex;align-items:center;
        justify-content:space-between;padding:14px 0;gap:12px;flex-wrap:wrap}
.logo-area{display:flex;align-items:center;gap:14px}
.logo-box{
  width:50px;height:50px;border-radius:10px;
  background:#fff;display:flex;align-items:center;justify-content:center;
  flex-shrink:0;overflow:hidden;
}
.logo-box span{
  font-size:1rem;font-weight:900;color:var(--ifp-red);letter-spacing:-1px;line-height:1
}
.logo-txt h1{font-size:1.15rem;font-weight:800;line-height:1.2;color:#fff}
.logo-txt p{font-size:0.73rem;opacity:0.85;margin-top:2px;color:rgba(255,255,255,0.9)}
.periodo-badge{
  background:rgba(255,255,255,0.15);border:1.5px solid rgba(255,255,255,0.35);
  border-radius:20px;padding:6px 16px;font-size:0.8rem;font-weight:700;
  white-space:nowrap;color:#fff;
}

/* ── BARRA AZUL UPLOAD ── */
.upload-bar{background:var(--ifp-blue);border-bottom:3px solid var(--ifp-red);padding:11px 24px}
.upload-in{max-width:1600px;margin:0 auto;display:flex;align-items:center;gap:12px;flex-wrap:wrap}
.upload-bar label{font-weight:700;font-size:0.82rem;color:rgba(255,255,255,0.85);white-space:nowrap}
.upload-bar form{display:flex;align-items:center;gap:10px;flex:1;flex-wrap:wrap}
.upload-bar input[type=file]{
  flex:1;min-width:160px;font-size:0.82rem;padding:6px 10px;
  border:1.5px solid rgba(255,255,255,0.3);border-radius:8px;
  background:rgba(255,255,255,0.1);cursor:pointer;color:#fff;
}
.upload-bar input[type=file]::-webkit-file-upload-button{
  background:var(--ifp-red);color:#fff;border:none;border-radius:6px;
  padding:4px 12px;font-size:0.8rem;font-weight:700;cursor:pointer;margin-right:8px;
}
.btn-up{
  background:var(--ifp-red);color:#fff;border:none;border-radius:8px;
  padding:8px 18px;font-size:0.85rem;font-weight:700;cursor:pointer;
  white-space:nowrap;transition:background .15s;
}
.btn-up:hover{background:var(--ifp-red-dk)}
.btn-pdf{
  background:transparent;color:#fff;border:1.5px solid rgba(255,255,255,0.5);
  border-radius:8px;padding:8px 16px;font-size:0.82rem;font-weight:700;
  cursor:pointer;white-space:nowrap;transition:all .15s;text-decoration:none;
  display:inline-flex;align-items:center;gap:6px;
}
.btn-pdf:hover{background:rgba(255,255,255,0.12);border-color:#fff}
.smsg{font-size:0.8rem;font-weight:600;color:rgba(255,255,255,0.9)}
.smsg.ok{color:#81c784}
.smsg.err{color:#ef9a9a}

/* ── FILTROS ── */
.filtros{background:#fff;border-bottom:2px solid var(--borda);padding:10px 24px}
.filtros-in{max-width:1600px;margin:0 auto;display:flex;align-items:center;gap:10px;flex-wrap:wrap}
.flabel{font-size:0.8rem;font-weight:700;color:var(--ifp-blue)}
.fbtn{
  padding:5px 15px;border-radius:20px;border:1.5px solid var(--borda);
  background:#f4f5f9;color:var(--txt2);font-size:0.8rem;font-weight:600;
  cursor:pointer;transition:all .15s;
}
.fbtn:hover{border-color:var(--ifp-red);color:var(--ifp-red)}
.fbtn.ativo{background:var(--ifp-blue);color:#fff;border-color:var(--ifp-blue)}
.search{
  padding:6px 14px;border-radius:20px;border:1.5px solid var(--borda);
  background:#f4f5f9;font-size:0.82rem;outline:none;width:210px;font-family:inherit;
}
.search:focus{border-color:var(--ifp-red)}

/* ── TOTALIZADORES ── */
.totais{max-width:1600px;margin:18px auto 0;padding:0 24px}
.totais-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(185px,1fr));gap:12px}
.tc{
  background:#fff;border-radius:var(--r);padding:16px 18px;
  box-shadow:var(--shadow);border-top:4px solid var(--ifp-blue);
  display:flex;flex-direction:column;gap:3px;
}
.tc-lbl{font-size:0.67rem;font-weight:700;color:var(--txt2);text-transform:uppercase;letter-spacing:.06em}
.tc-val{font-size:1.4rem;font-weight:900;color:var(--ifp-blue)}
.tc-sub{font-size:0.7rem;color:var(--txt2)}
.tc.tc-red{border-top-color:var(--ifp-red)} .tc.tc-red .tc-val{color:var(--ifp-red)}
.tc.tc-verde{border-top-color:var(--verde-lt)} .tc.tc-verde .tc-val{color:var(--verde)}
.tc.tc-laranja{border-top-color:var(--laranja-lt)} .tc.tc-laranja .tc-val{color:var(--laranja)}
.tc.tc-vermelho{border-top-color:#e53935} .tc.tc-vermelho .tc-val{color:#c62828}

/* ── LEGENDA ── */
.legenda{max-width:1600px;margin:12px auto 0;padding:0 24px;display:flex;align-items:center;gap:14px;flex-wrap:wrap}
.leg{display:flex;align-items:center;gap:5px;font-size:0.75rem;font-weight:600;color:var(--txt2)}
.dot{width:10px;height:10px;border-radius:50%;flex-shrink:0}
.dot.v{background:var(--verde-lt)} .dot.a{background:var(--laranja-lt)} .dot.r{background:#e53935}

/* ── CARDS ── */
.cards{max-width:1600px;margin:18px auto 50px;padding:0 24px}
.cards-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(340px,1fr));gap:14px}

.ucard{
  background:#fff;border-radius:var(--r);box-shadow:var(--shadow);
  overflow:hidden;border:1.5px solid var(--borda);
  transition:box-shadow .2s,transform .2s;
}
.ucard:hover{box-shadow:var(--shadow-hover);transform:translateY(-2px)}

/* cabeçalho card */
.chd{padding:13px 15px;display:flex;align-items:center;justify-content:space-between;gap:8px}
.chd.bom{background:linear-gradient(135deg,#1b5e20,#2e7d32);color:#fff}
.chd.medio{background:linear-gradient(135deg,#bf360c,#e64a19);color:#fff}
.chd.ruim{background:linear-gradient(135deg,var(--ifp-red-dk),var(--ifp-red));color:#fff}
.cnome{font-size:0.92rem;font-weight:800;line-height:1.2;flex:1}
.sbadge{
  background:rgba(255,255,255,0.2);border:1px solid rgba(255,255,255,0.35);
  border-radius:20px;padding:3px 10px;font-size:0.7rem;font-weight:700;white-space:nowrap;
}

/* corpo card */
.cbody{padding:12px 13px;display:flex;flex-direction:column;gap:6px}

.ind{display:flex;align-items:center;gap:8px;padding:5px 8px;border-radius:7px}
.ind.verde{background:#e8f5e9} .ind.amarelo{background:#fff3e0} .ind.vermelho{background:#ffebee}
.ind-lbl{font-size:0.74rem;font-weight:600;color:var(--txt2);flex:1;min-width:0}
.ind-meta{font-size:0.61rem;color:var(--txt2);opacity:.7;white-space:nowrap}
.ind-val{font-size:0.84rem;font-weight:700;white-space:nowrap}
.ind.verde .ind-val{color:var(--verde)} .ind.amarelo .ind-val{color:var(--laranja)} .ind.vermelho .ind-val{color:#c62828}
.barwrap{width:52px;height:5px;background:#e0e4f0;border-radius:3px;overflow:hidden;flex-shrink:0}
.bar{height:100%;border-radius:3px}
.bar.verde{background:var(--verde-lt)} .bar.amarelo{background:var(--laranja-lt)} .bar.vermelho{background:#e53935}

.fin-row{display:grid;grid-template-columns:1fr 1fr 1fr;gap:5px;padding-top:6px}
.fbox{background:#f4f5f9;border-radius:7px;padding:6px 8px;text-align:center}
.fbox-lbl{font-size:0.59rem;font-weight:700;color:var(--txt2);text-transform:uppercase;letter-spacing:.04em}
.fbox-val{font-size:0.87rem;font-weight:800;color:var(--ifp-blue);margin-top:2px}

.alunos-row{display:grid;grid-template-columns:repeat(4,1fr);gap:5px;padding-top:2px}
.abox{background:#f4f5f9;border-radius:7px;padding:5px 5px;text-align:center}
.abox-lbl{font-size:0.57rem;font-weight:700;color:var(--txt2);text-transform:uppercase}
.abox-val{font-size:0.88rem;font-weight:800;color:var(--ifp-blue);margin-top:1px}
.abox.ok .abox-val{color:var(--verde)} .abox.alerta .abox-val{color:var(--ifp-red)}

.fat-total{
  display:flex;align-items:center;justify-content:space-between;
  padding:7px 9px;border-radius:8px;
  background:linear-gradient(135deg,#e8eaf6,#c5cae9);margin-top:2px;
}
.fat-lbl{font-size:0.74rem;font-weight:700;color:var(--ifp-blue)}
.fat-val{font-size:0.92rem;font-weight:900;color:var(--ifp-blue)}

/* botão PDF no card */
.btn-card-pdf{
  display:flex;align-items:center;justify-content:center;gap:6px;
  margin-top:8px;padding:7px 0;border-radius:8px;
  background:var(--ifp-red);color:#fff;font-size:0.77rem;font-weight:700;
  text-decoration:none;transition:background .15s;
}
.btn-card-pdf:hover{background:var(--ifp-red-dk)}

/* ── SEM DADOS ── */
.nodata{max-width:700px;margin:70px auto;padding:0 24px;text-align:center}
.nodata-box{background:#fff;border-radius:var(--r);padding:60px 40px;box-shadow:var(--shadow)}
.nodata-icon{font-size:4rem;margin-bottom:14px}
.nodata-box h2{font-size:1.35rem;font-weight:800;color:var(--ifp-blue);margin-bottom:10px}
.nodata-box p{color:var(--txt2);font-size:0.92rem;line-height:1.7;margin-bottom:6px}
.nodata-box small{color:var(--txt2);font-size:0.78rem;opacity:.8}

/* ── PRINT / PDF ── */
@media print{
  .hdr,.upload-bar,.filtros,.totais,.legenda,.btn-card-pdf,.btn-pdf{display:none!important}
  .cards{margin:0;padding:0}
  .cards-grid{grid-template-columns:repeat(2,1fr);gap:10px}
  .ucard{break-inside:avoid;box-shadow:none;border:1px solid #ccc}
  body{background:#fff}
}

@media(max-width:640px){
  .cards-grid{grid-template-columns:1fr}
  .totais-grid{grid-template-columns:repeat(2,1fr)}
  .fin-row{grid-template-columns:1fr 1fr}
  .alunos-row{grid-template-columns:repeat(2,1fr)}
}
</style>
</head>
<body>

<!-- HEADER -->
<div class="hdr">
  <div class="hdr-in">
    <div class="logo-area">
      <div class="logo-box"><span>IFP</span></div>
      <div class="logo-txt">
        <h1>Dashboard de Fechamento Mensal</h1>
        <p>Instituto de Formação Profissional</p>
      </div>
    </div>
    {% if periodo %}<div class="periodo-badge">📅 {{ periodo }}</div>{% endif %}
  </div>
</div>

<!-- BARRA UPLOAD + PDF GERAL -->
<div class="upload-bar">
  <div class="upload-in">
    <label>📁 Planilha:</label>
    <form method="POST" action="/upload" enctype="multipart/form-data">
      <input type="file" name="planilha" accept=".xlsx,.xlsm">
      <button type="submit" class="btn-up">⬆ Carregar</button>
    </form>
    {% if unidades %}
    <a href="/pdf/todas" class="btn-pdf" target="_blank">⬇ PDF — Todas as Unidades</a>
    {% endif %}
    {% if msg %}
    <span class="smsg {{ 'ok' if msg_ok else 'err' }}">{{ msg }}</span>
    {% endif %}
  </div>
</div>

{% if unidades %}

<!-- FILTROS -->
<div class="filtros">
  <div class="filtros-in">
    <span class="flabel">Filtrar:</span>
    <button class="fbtn ativo" onclick="filtrar('todos',this)">Todos ({{ unidades|length }})</button>
    <button class="fbtn" onclick="filtrar('bom',this)" style="color:#2e7d32;border-color:#a5d6a7">
      ✅ Bom ({{ unidades|selectattr('status','eq','bom')|list|length }})
    </button>
    <button class="fbtn" onclick="filtrar('medio',this)" style="color:#e64a19;border-color:#ffcc80">
      ⚠️ Médio ({{ unidades|selectattr('status','eq','medio')|list|length }})
    </button>
    <button class="fbtn" onclick="filtrar('ruim',this)" style="color:#c62828;border-color:#ef9a9a">
      ❌ Ruim ({{ unidades|selectattr('status','eq','ruim')|list|length }})
    </button>
    <input type="text" class="search" placeholder="🔍 Buscar unidade..." oninput="buscar(this.value)">
  </div>
</div>

<!-- TOTALIZADORES -->
<div class="totais">
  <div class="totais-grid">
    <div class="tc">
      <span class="tc-lbl">Matrículas Totais</span>
      <span class="tc-val">{{ totais.matriculas|toint }}</span>
      <span class="tc-sub">Perda média: {{ totais.perda_str }}</span>
    </div>
    <div class="tc tc-red">
      <span class="tc-lbl">Alunos Ativos</span>
      <span class="tc-val">{{ totais.ativos|toint }}</span>
      <span class="tc-sub">Retenção média: {{ totais.retencao_str }}</span>
    </div>
    <div class="tc">
      <span class="tc-lbl">Fat. Comercial Total</span>
      <span class="tc-val">{{ totais.fat_comercial|brl0 }}</span>
      <span class="tc-sub">Ticket médio: {{ totais.ticket_str }}</span>
    </div>
    <div class="tc">
      <span class="tc-lbl">Fat. Total (carteira)</span>
      <span class="tc-val">{{ totais.fat_total|brl0 }}</span>
      <span class="tc-sub">{{ unidades|length }} unidades</span>
    </div>
    <div class="tc tc-verde">
      <span class="tc-lbl">✅ Unidades Boas</span>
      <span class="tc-val">{{ unidades|selectattr('status','eq','bom')|list|length }}</span>
      <span class="tc-sub">≥ 2 indicadores na meta</span>
    </div>
    <div class="tc tc-laranja">
      <span class="tc-lbl">⚠️ Unidades Médias</span>
      <span class="tc-val">{{ unidades|selectattr('status','eq','medio')|list|length }}</span>
      <span class="tc-sub">1 indicador na meta</span>
    </div>
    <div class="tc tc-vermelho">
      <span class="tc-lbl">❌ Unidades Ruins</span>
      <span class="tc-val">{{ unidades|selectattr('status','eq','ruim')|list|length }}</span>
      <span class="tc-sub">0 indicadores na meta</span>
    </div>
  </div>
</div>

<!-- LEGENDA -->
<div class="legenda">
  <span class="flabel">Metas:</span>
  <span class="leg"><span class="dot v"></span>Matrículas ≥ 120</span>
  <span class="leg"><span class="dot v"></span>Ticket ≥ R$199</span>
  <span class="leg"><span class="dot v"></span>Cobrança ≥ 94%</span>
  <span class="leg"><span class="dot v"></span>Frequência ≥ 75%</span>
  <span class="leg"><span class="dot v"></span>Retenção ≥ 94%</span>
  <span class="leg"><span class="dot v"></span>✅ Bom = ≥2 metas</span>
  <span class="leg"><span class="dot a"></span>⚠️ Médio = 1 meta</span>
  <span class="leg"><span class="dot r"></span>❌ Ruim = 0 metas</span>
</div>

<!-- CARDS -->
<div class="cards">
  <div class="cards-grid" id="cards-grid">
    {% for u in unidades %}
    <div class="ucard" data-status="{{ u.status }}" data-nome="{{ u.nome|lower }}">
      <div class="chd {{ u.status }}">
        <span class="cnome">{{ u.nome }}</span>
        <span class="sbadge">
          {%- if u.status=='bom' %}✅{%- elif u.status=='medio' %}⚠️{%- else %}❌{%- endif -%}
          &nbsp;{{ u.score }}/5
        </span>
      </div>
      <div class="cbody">
        {% for ind in u.indicators %}
        <div class="ind {{ ind.cls }}">
          <span class="ind-lbl">{{ ind.label }}</span>
          <span class="ind-meta">{{ ind.meta }}</span>
          <span class="ind-val">{{ ind.display }}</span>
          <div class="barwrap"><div class="bar {{ ind.cls }}" style="width:{{ ind.bar }}%"></div></div>
        </div>
        {% endfor %}
        <div class="fin-row">
          <div class="fbox">
            <div class="fbox-lbl">Cobr. 30d</div>
            <div class="fbox-val">{{ u.fin_30|pct }}</div>
          </div>
          <div class="fbox">
            <div class="fbox-lbl">Cobr. 60d</div>
            <div class="fbox-val">{{ u.fin_60|pct }}</div>
          </div>
          <div class="fbox">
            <div class="fbox-lbl">Fat. Comercial</div>
            <div class="fbox-val">{{ u.fat_comercial|brl0 }}</div>
          </div>
        </div>
        <div class="alunos-row">
          <div class="abox ok">
            <div class="abox-lbl">Ativos</div>
            <div class="abox-val">{{ u.ativos|toint }}</div>
          </div>
          <div class="abox {{ 'alerta' if u.cancelados > 20 else '' }}">
            <div class="abox-lbl">Cancelados</div>
            <div class="abox-val">{{ u.cancelados|toint }}</div>
          </div>
          <div class="abox {{ 'alerta' if u.desistentes > 50 else '' }}">
            <div class="abox-lbl">Desistentes</div>
            <div class="abox-val">{{ u.desistentes|toint }}</div>
          </div>
          <div class="abox {{ 'alerta' if u.nunca_veio > 30 else '' }}">
            <div class="abox-lbl">Nunca Veio</div>
            <div class="abox-val">{{ u.nunca_veio|toint }}</div>
          </div>
        </div>
        <div class="fat-total">
          <span class="fat-lbl">Faturamento Total (carteira)</span>
          <span class="fat-val">{{ u.fat_total|brl }}</span>
        </div>
        <a href="/pdf/unidade/{{ loop.index0 }}" target="_blank" class="btn-card-pdf">
          ⬇ Baixar PDF desta unidade
        </a>
      </div>
    </div>
    {% endfor %}
  </div>
</div>

{% else %}
<div class="nodata">
  <div class="nodata-box">
    <div class="nodata-icon">📊</div>
    <h2>Nenhum dado carregado</h2>
    <p>Faça o upload da planilha de fechamento mensal para visualizar os indicadores.</p>
    <p>Formatos aceitos: <strong>.xlsx</strong> ou <strong>.xlsm</strong></p>
    <br>
    <small>
      As abas precisam ter nomes no formato:<br>
      <strong>IFP - Unidade (Visitas)</strong>,
      <strong>IFP - Unidade (Frequência)</strong>,
      <strong>IFP - Unidade (Matrícula e Quitação)</strong>,
      <strong>IFP - Unidade (Histórico)</strong>
    </small>
  </div>
</div>
{% endif %}

<script>
var filtroAtivo='todos';
function filtrar(status,btn){
  filtroAtivo=status;
  document.querySelectorAll('.fbtn').forEach(function(b){b.classList.remove('ativo')});
  if(btn) btn.classList.add('ativo');
  var q=document.querySelector('.search')?document.querySelector('.search').value.toLowerCase():'';
  aplicar(q);
}
function buscar(q){ aplicar(q.toLowerCase()); }
function aplicar(q){
  q=q||'';
  document.querySelectorAll('.ucard').forEach(function(c){
    var okS=filtroAtivo==='todos'||c.dataset.status===filtroAtivo;
    var okN=!q||c.dataset.nome.includes(q);
    c.style.display=(okS&&okN)?'':'none';
  });
}
</script>
</body>
</html>
"""


# ─────────────────────────────────────────────────────────────
# TEMPLATE PDF (HTML que o navegador imprime como PDF)
# ─────────────────────────────────────────────────────────────
PDF_TEMPLATE = """
<!DOCTYPE html>
<html lang="pt-br">
<head>
<meta charset="UTF-8">
<title>IFP – Relatório {{ titulo }}</title>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800;900&display=swap" rel="stylesheet">
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Inter',sans-serif;background:#fff;color:#1a1f36;font-size:11px}
.page-header{
  background:linear-gradient(135deg,#8b0014,#c0021c);
  color:#fff;padding:18px 24px;display:flex;align-items:center;
  justify-content:space-between;margin-bottom:18px;
}
.ph-logo{font-size:1.4rem;font-weight:900;letter-spacing:-1px}
.ph-info h2{font-size:1rem;font-weight:800;text-align:right}
.ph-info p{font-size:0.72rem;opacity:.85;text-align:right;margin-top:3px}
.grid{display:grid;grid-template-columns:repeat(2,1fr);gap:12px;padding:0 18px 18px}
.ucard{border:1.5px solid #dde1ef;border-radius:10px;overflow:hidden;break-inside:avoid}
.chd{padding:10px 13px;display:flex;align-items:center;justify-content:space-between}
.chd.bom{background:linear-gradient(135deg,#1b5e20,#2e7d32);color:#fff}
.chd.medio{background:linear-gradient(135deg,#bf360c,#e64a19);color:#fff}
.chd.ruim{background:linear-gradient(135deg,#8b0014,#c0021c);color:#fff}
.cnome{font-size:0.88rem;font-weight:800}
.sbadge{font-size:0.68rem;font-weight:700;background:rgba(255,255,255,.2);
        border:1px solid rgba(255,255,255,.3);border-radius:12px;padding:2px 8px}
.cbody{padding:10px 12px;display:flex;flex-direction:column;gap:5px}
.ind{display:flex;align-items:center;gap:7px;padding:4px 7px;border-radius:6px}
.ind.verde{background:#e8f5e9} .ind.amarelo{background:#fff3e0} .ind.vermelho{background:#ffebee}
.ind-lbl{font-size:0.7rem;font-weight:600;color:#5a6282;flex:1}
.ind-meta{font-size:0.58rem;color:#5a6282;opacity:.7}
.ind-val{font-size:0.8rem;font-weight:700}
.ind.verde .ind-val{color:#2e7d32} .ind.amarelo .ind-val{color:#e65100} .ind.vermelho .ind-val{color:#c62828}
.barwrap{width:45px;height:4px;background:#e0e4f0;border-radius:2px;overflow:hidden}
.bar{height:100%;border-radius:2px}
.bar.verde{background:#43a047} .bar.amarelo{background:#fb8c00} .bar.vermelho{background:#e53935}
.fin-row{display:grid;grid-template-columns:1fr 1fr 1fr;gap:4px;margin-top:4px}
.fbox{background:#f4f5f9;border-radius:5px;padding:4px 6px;text-align:center}
.fbox-lbl{font-size:0.56rem;font-weight:700;color:#5a6282;text-transform:uppercase}
.fbox-val{font-size:0.82rem;font-weight:800;color:#1a237e;margin-top:1px}
.alunos-row{display:grid;grid-template-columns:repeat(4,1fr);gap:4px;margin-top:4px}
.abox{background:#f4f5f9;border-radius:5px;padding:4px 5px;text-align:center}
.abox-lbl{font-size:0.54rem;font-weight:700;color:#5a6282;text-transform:uppercase}
.abox-val{font-size:0.82rem;font-weight:800;color:#1a237e;margin-top:1px}
.abox.ok .abox-val{color:#2e7d32} .abox.alerta .abox-val{color:#c0021c}
.fat-total{display:flex;align-items:center;justify-content:space-between;
           padding:5px 7px;border-radius:6px;background:#e8eaf6;margin-top:4px}
.fat-lbl{font-size:0.7rem;font-weight:700;color:#1a237e}
.fat-val{font-size:0.88rem;font-weight:900;color:#1a237e}
.rodape{text-align:center;color:#999;font-size:0.68rem;padding:10px 0 16px;border-top:1px solid #eee;margin:0 18px}
@media print{
  body{-webkit-print-color-adjust:exact;print-color-adjust:exact}
  .no-print{display:none}
}
</style>
</head>
<body>
<div class="page-header">
  <div class="ph-logo">IFP</div>
  <div class="ph-info">
    <h2>Relatório de Fechamento — {{ titulo }}</h2>
    <p>Instituto de Formação Profissional &nbsp;|&nbsp; {{ periodo }}</p>
  </div>
</div>

<div class="no-print" style="text-align:center;margin-bottom:12px">
  <button onclick="window.print()"
    style="background:#c0021c;color:#fff;border:none;border-radius:8px;
           padding:10px 28px;font-size:0.9rem;font-weight:700;cursor:pointer;">
    🖨️ Imprimir / Salvar PDF
  </button>
  &nbsp;
  <button onclick="window.close()"
    style="background:#1a237e;color:#fff;border:none;border-radius:8px;
           padding:10px 20px;font-size:0.9rem;font-weight:700;cursor:pointer;">
    ✕ Fechar
  </button>
</div>

<div class="grid">
{% for u in unidades %}
<div class="ucard">
  <div class="chd {{ u.status }}">
    <span class="cnome">{{ u.nome }}</span>
    <span class="sbadge">
      {%- if u.status=='bom' %}✅{%- elif u.status=='medio' %}⚠️{%- else %}❌{%- endif -%}
      &nbsp;{{ u.score }}/5
    </span>
  </div>
  <div class="cbody">
    {% for ind in u.indicators %}
    <div class="ind {{ ind.cls }}">
      <span class="ind-lbl">{{ ind.label }}</span>
      <span class="ind-meta">{{ ind.meta }}</span>
      <span class="ind-val">{{ ind.display }}</span>
      <div class="barwrap"><div class="bar {{ ind.cls }}" style="width:{{ ind.bar }}%"></div></div>
    </div>
    {% endfor %}
    <div class="fin-row">
      <div class="fbox"><div class="fbox-lbl">Cobr. 30d</div><div class="fbox-val">{{ u.fin_30|pct }}</div></div>
      <div class="fbox"><div class="fbox-lbl">Cobr. 60d</div><div class="fbox-val">{{ u.fin_60|pct }}</div></div>
      <div class="fbox"><div class="fbox-lbl">Fat. Comercial</div><div class="fbox-val">{{ u.fat_comercial|brl0 }}</div></div>
    </div>
    <div class="alunos-row">
      <div class="abox ok"><div class="abox-lbl">Ativos</div><div class="abox-val">{{ u.ativos|toint }}</div></div>
      <div class="abox {{ 'alerta' if u.cancelados > 20 else '' }}">
        <div class="abox-lbl">Cancelados</div><div class="abox-val">{{ u.cancelados|toint }}</div>
      </div>
      <div class="abox {{ 'alerta' if u.desistentes > 50 else '' }}">
        <div class="abox-lbl">Desistentes</div><div class="abox-val">{{ u.desistentes|toint }}</div>
      </div>
      <div class="abox {{ 'alerta' if u.nunca_veio > 30 else '' }}">
        <div class="abox-lbl">Nunca Veio</div><div class="abox-val">{{ u.nunca_veio|toint }}</div>
      </div>
    </div>
    <div class="fat-total">
      <span class="fat-lbl">Faturamento Total</span>
      <span class="fat-val">{{ u.fat_total|brl }}</span>
    </div>
  </div>
</div>
{% endfor %}
</div>

<div class="rodape">
  Gerado pelo sistema IFP Dashboard &nbsp;|&nbsp; {{ periodo }}
</div>

<script>
// Auto-abre o diálogo de impressão ao abrir a página de PDF de unidade única
{% if auto_print %}
window.addEventListener('load', function(){ setTimeout(function(){ window.print(); }, 800); });
{% endif %}
</script>
</body>
</html>
"""


# ─────────────────────────────────────────────────────────────
# ROTAS FLASK
# ─────────────────────────────────────────────────────────────

@app.route("/", methods=["GET"])
def index():
    data, periodo, filename = get_cached_data()
    unidades = data or []
    totais = {}
    if unidades:
        total_mat   = sum(u["matriculas"]    for u in unidades)
        total_fat_c = sum(u["fat_comercial"] for u in unidades)
        total_fat_t = sum(u["fat_total"]     for u in unidades)
        total_ativ  = sum(u["ativos"]        for u in unidades)
        ret_list    = [u["retencao"] for u in unidades if u["retencao"] > 0]
        perda_list  = [u["m1_v1"]   for u in unidades if u["m1_v1"]   > 0]
        totais = {
            "matriculas":    total_mat,
            "ativos":        total_ativ,
            "fat_comercial": total_fat_c,
            "fat_total":     total_fat_t,
            "ticket_str":    fmt_brl0(total_fat_c / total_mat if total_mat else 0),
            "retencao_str":  f"{(sum(ret_list)/len(ret_list) if ret_list else 0)*100:.1f}%",
            "perda_str":     f"{(sum(perda_list)/len(perda_list) if perda_list else 0)*100:.1f}%",
        }
    return render_template_string(
        TEMPLATE,
        unidades=unidades,
        totais=totais,
        periodo=periodo,
        msg=request.args.get("msg", ""),
        msg_ok=request.args.get("ok", "1") == "1",
    )


@app.route("/upload", methods=["POST"])
def upload():
    if "planilha" not in request.files:
        return redirect("/?msg=Nenhum+arquivo+enviado&ok=0")
    f = request.files["planilha"]
    if not f or not f.filename:
        return redirect("/?msg=Nenhum+arquivo+selecionado&ok=0")
    if not f.filename.lower().endswith((".xlsx", ".xlsm")):
        return redirect("/?msg=Formato+inválido.+Use+.xlsx+ou+.xlsm&ok=0")
    try:
        file_bytes = f.read()
        unidades   = load_data_from_bytes(file_bytes)
        periodo    = detect_periodo(file_bytes) or f.filename
        set_cached_data(unidades, periodo, f.filename)
        n = len(unidades)
        return redirect(f"/?msg=✅+Planilha+carregada!+{n}+unidades+encontradas&ok=1")
    except KeyError as e:
        return redirect(f"/?msg=Erro:+{str(e)[:120]}&ok=0")
    except Exception as e:
        traceback.print_exc()
        return redirect(f"/?msg=Erro+ao+processar:+{str(e)[:120]}&ok=0")


@app.route("/pdf/todas")
def pdf_todas():
    """Retorna página HTML pronta para impressão com TODAS as unidades."""
    data, periodo, _ = get_cached_data()
    if not data:
        return redirect("/?msg=Nenhum+dado+carregado&ok=0")
    html = render_template_string(
        PDF_TEMPLATE,
        unidades=data,
        titulo="Todas as Unidades",
        periodo=periodo or "—",
        auto_print=False,
    )
    return Response(html, mimetype="text/html")


@app.route("/pdf/unidade/<int:idx>")
def pdf_unidade(idx):
    """Retorna página HTML pronta para impressão de UMA unidade."""
    data, periodo, _ = get_cached_data()
    if not data or idx >= len(data):
        return redirect("/?msg=Unidade+não+encontrada&ok=0")
    u = data[idx]
    html = render_template_string(
        PDF_TEMPLATE,
        unidades=[u],
        titulo=u["nome"],
        periodo=periodo or "—",
        auto_print=True,
    )
    return Response(html, mimetype="text/html")


@app.route("/api/dados", methods=["GET"])
def api_dados():
    data, periodo, filename = get_cached_data()
    return jsonify({
        "periodo":  periodo,
        "filename": filename,
        "total":    len(data or []),
        "unidades": data or [],
    })


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
