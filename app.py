import os, sys, hashlib, json, re, time, threading, webbrowser
from datetime import datetime, date, timedelta
from collections import defaultdict

print("Carregando modulos...", flush=True)

try:
    import pandas as pd
    from flask import Flask, Response, request, redirect
    import gspread
    from google.oauth2.service_account import Credentials
    print("  OK", flush=True)
except ImportError as e:
    print(f"  ERRO: {e}", flush=True)
    print("  Rode: pip install flask pandas gspread google-auth", flush=True)
    input("Pressione Enter para sair...")
    sys.exit(1)

# ══════════════════════════════════════════════════════════════
ID_RESIDUOS    = "1e31JtS7fuuHAk-JuetIF2yZk0iWsY44193dABLPx0-M"
ID_FATURAMENTO = "1Xc36v2pNo8VolDgQek3hBs9hYnz7sClqWhuDS_8F8lA"
CREDENCIAIS_JSON = "credenciais.json"
HASH_ADMIN  = "240be518fabd2724ddb6f04eeb1da5967448d7e831c08c8fa822809f74c720a9"
HASH_VIEWER = "151dd36dba47a3a3216ca0674f23cfb34ab0666dd5ab6077abff72b8771d5c30"
PORTA = int(os.environ.get("PORT", 5000))

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", os.urandom(24))
PASTA = os.path.dirname(os.path.abspath(__file__))
_cache = {"ds": {}, "fat": {}, "ultima_atualizacao": None}

# ── GOOGLE SHEETS ──────────────────────────────────────────────

def get_gc():
    cred_path = os.path.join(PASTA, CREDENCIAIS_JSON)
    cred_env = os.environ.get("GOOGLE_CREDENTIALS")
    if cred_env:
        import tempfile
        tmp = tempfile.NamedTemporaryFile(mode="w", suffix=".json", delete=False)
        tmp.write(cred_env); tmp.close()
        cred_path = tmp.name
    scopes = ["https://spreadsheets.google.com/feeds",
              "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_file(cred_path, scopes=scopes)
    return gspread.authorize(creds)


def ler_aba(gc, sheet_id, nome_aba):
    try:
        sh    = gc.open_by_key(sheet_id)
        ws    = sh.worksheet(nome_aba)
        dados = ws.get_all_values()
        if not dados or len(dados) < 2:
            print(f"  Aba '{nome_aba}': vazia", flush=True)
            return pd.DataFrame()
        cabecalhos = [str(c).strip() for c in dados[0]]
        df = pd.DataFrame(dados[1:], columns=cabecalhos)
        for col in df.columns:
            df[col] = df[col].apply(lambda v: str(v).strip() if v is not None else "")
        print(f"  Aba '{nome_aba}': {len(df)} linhas | colunas: {cabecalhos}", flush=True)
        return df
    except gspread.exceptions.WorksheetNotFound:
        print(f"  AVISO: aba '{nome_aba}' nao encontrada", flush=True)
        return pd.DataFrame()
    except Exception as e:
        print(f"  ERRO ao ler '{nome_aba}': {e}", flush=True)
        return pd.DataFrame()

# ── HELPERS ────────────────────────────────────────────────────

MESES_BR = {"jan":1,"fev":2,"mar":3,"abr":4,"mai":5,"jun":6,
            "jul":7,"ago":8,"set":9,"out":10,"nov":11,"dez":12}

def fmt_data(v):
    if not v or str(v).strip() in ("","nan","None","NaT"): return None
    s = str(v).strip()
    if re.match(r'^\d{4,6}$', s):
        try:
            d = date(1899,12,30) + timedelta(days=int(s))
            return d.strftime("%Y-%m-%d")
        except: pass
    partes = re.split(r'[-/]', s.lower())
    if len(partes) == 3:
        try:
            dia = int(partes[0])
            mes = MESES_BR.get(partes[1].replace(".","").strip())
            ano = int(partes[2])
            if ano < 100: ano += 2000
            if mes: return datetime(ano, mes, dia).strftime("%Y-%m-%d")
        except: pass
    # Google Sheets exporta datas no formato americano M/D/YYYY (ex: 5/1/2026 = 1 de maio)
    # Detectar se é formato americano: mês/dia/ano
    partes_barra = s.split("/")
    if len(partes_barra) == 3:
        try:
            mes  = int(partes_barra[0])
            dia  = int(partes_barra[1])
            ano  = int(partes_barra[2])
            if ano < 100: ano += 2000
            if 1 <= mes <= 12 and 1 <= dia <= 31:
                return datetime(ano, mes, dia).strftime("%Y-%m-%d")
        except: pass
    for fmt in ("%Y-%m-%d","%d-%m-%Y","%d/%m/%y"):
        try: return datetime.strptime(s[:10], fmt).strftime("%Y-%m-%d")
        except: pass
    return None

def safe_str(v, d=""):
    s = str(v).strip() if v is not None else d
    return d if s.lower() in ("nan","none","") else s

def safe_float(v):
    try:
        s = str(v).replace("R$","").replace("\xa0","").replace("\u00a0","")
        s = re.sub(r"\s+","",s).strip()
        if not s or s.lower() in ("nan","none","-",""): return 0.0
        if re.match(r"^\d{1,3}(\.\d{3})*(,\d+)?$", s):
            s = s.replace(".","").replace(",",".")
        else:
            s = s.replace(",",".")
        return round(float(s), 4)
    except: return 0.0

def norm_classe(c):
    s = safe_str(c)
    if not s: return ""
    sufixo = re.sub(r"(?i)^classe\s*","",s).strip().upper()
    if not sufixo: return ""
    MAP = {"1":"I","2A":"IIA","2B":"IIB","II A":"IIA","II B":"IIB","2 A":"IIA","2 B":"IIB"}
    sufixo = MAP.get(sufixo, sufixo)
    return "Classe " + sufixo

# ── RESÍDUOS ───────────────────────────────────────────────────

def processar_residuos(gc):
    print("\nProcessando residuos...", flush=True)
    ds = defaultdict(lambda: {"clientes":[],"saidas":[],"dest":[],"res":[]})

    lookup = {}
    df_f = ler_aba(gc, ID_RESIDUOS, "Formulas")
    if not df_f.empty:
        for _, row in df_f.iterrows():
            g = safe_str(row.get("Gerador",""))
            if g:
                lookup[g.lower()] = {
                    "tipo":   safe_str(row.get("Tipo","Privado")) or "Privado",
                    "cidade": safe_str(row.get("Cidade","")),
                    "estado": safe_str(row.get("Estado","")),
                }
        print(f"  Formulas: {len(lookup)} geradores", flush=True)

    df = ler_aba(gc, ID_RESIDUOS, "Entrada de Residuos") if False else ler_aba(gc, ID_RESIDUOS, "Entrada de Resíduos")
    if not df.empty:
        dm, rm = defaultdict(dict), defaultdict(dict)
        for i, row in df.iterrows():
            data = fmt_data(row.get("Data",""))
            if not data: continue
            key     = data[:7]
            gerador = safe_str(row.get("Gerador",""))
            if not gerador: continue
            tipo   = safe_str(row.get("Tipo",""))
            cidade = safe_str(row.get("Cidade",""))
            estado = safe_str(row.get("Estado",""))
            if not tipo or tipo.startswith("="):
                info   = lookup.get(gerador.lower(), {})
                tipo   = info.get("tipo","Privado")
                cidade = info.get("cidade", cidade)
                estado = info.get("estado", estado)
            desc   = safe_str(row.get("Descrição do resíduo",""))
            classe = norm_classe(row.get("Classe",""))
            peso   = safe_float(row.get("Peso (t)",0))
            dest   = safe_str(row.get("Destinação",""))
            ds[key]["clientes"].append({
                "id":i,"data":data,"nome":gerador,"tipo":tipo,
                "cidade":cidade,"estado":estado,"descricao":desc,
                "classe":classe,"peso":peso,"destinacao":dest
            })
            if dest:
                dk = dest.lower()
                if dk not in dm[key]: dm[key][dk] = {"id":20000+i,"nome":dest,"peso":0}
                dm[key][dk]["peso"] += peso
            if desc:
                rk = desc.lower()+"||"+classe
                if rk not in rm[key]: rm[key][rk] = {"id":30000+i,"nome":desc,"classe":classe,"peso":0}
                rm[key][rk]["peso"] += peso
        for k in dm: ds[k]["dest"] = sorted(dm[k].values(), key=lambda x:-x["peso"])
        for k in rm: ds[k]["res"]  = sorted(rm[k].values(), key=lambda x:-x["peso"])
        regs = sum(len(v["clientes"]) for v in ds.values())
        print(f"  Entrada: {regs} registros | {len(ds)} meses", flush=True)

    df_s = ler_aba(gc, ID_RESIDUOS, "Saída de Resíduos")
    if not df_s.empty:
        n = 0
        for i, row in df_s.iterrows():
            data = fmt_data(row.get("Data",""))
            if not data: continue
            ds[data[:7]]["saidas"].append({
                "id":10000+i,"data":data,
                "nome":safe_str(row.get("Descrição do resíduo","")),
                "tipo":safe_str(row.get("Destinação","")),
                "classe":norm_classe(row.get("Classe","")),
                "peso":safe_float(row.get("Peso (t)",0)),
                "destino":safe_str(row.get("Destino","")),
            })
            n += 1
        print(f"  Saidas: {n} registros", flush=True)

    return dict(sorted(ds.items()))

# ── FATURAMENTO ────────────────────────────────────────────────

def processar_faturamento(gc):
    import unicodedata
    print("\nProcessando faturamento...", flush=True)
    fat = {"empresas":[], "servicos":{k:[] for k in
        ["destinacao","transporte","locacao","limpeza","reciclaveis","maodeobra"]}}

    TIPO_MAP = {
        "destinação":"destinacao","destinacao":"destinacao",
        "transporte":"transporte",
        "locação":"locacao","locacao":"locacao",
        "limpeza pública":"limpeza","limpeza publica":"limpeza","limpeza":"limpeza",
        "venda recicláveis":"reciclaveis","venda reciclaveis":"reciclaveis","reciclaveis":"reciclaveis",
        "mão de obra":"maodeobra","mao de obra":"maodeobra","maodeobra":"maodeobra",
    }

    df = ler_aba(gc, ID_FATURAMENTO, "Financeiro")
    if df.empty:
        print("  AVISO: aba Financeiro vazia", flush=True)
        return fat

    print(f"  Colunas: {list(df.columns)}", flush=True)
    if len(df) > 0:
        r0 = df.iloc[0]
        print(f"  Linha1: Empresa='{r0.get('Empresa','')}' Valor='{r0.get('Valor (R$)','')}' Data='{r0.get('Data de Emissão','')}' Tipo='{r0.get('Tipo de serviço','')}'", flush=True)

    # Mapa nome_empresa → id fixo (gerado uma vez, reutilizado nas NFs)
    emp_id_map = {}
    ev = set()
    for i, row in df.iterrows():
        emp = safe_str(row.get("Empresa",""))
        if not emp: continue

        # Gerar id fixo por empresa baseado no nome
        if emp not in emp_id_map:
            emp_id_map[emp] = 60000 + len(emp_id_map)

        emp_id = emp_id_map[emp]

        tipo_raw  = safe_str(row.get("Tipo de serviço","")).lower()
        tipo_norm = unicodedata.normalize("NFD",tipo_raw).encode("ascii","ignore").decode().strip()
        tipo_key  = TIPO_MAP.get(tipo_raw) or TIPO_MAP.get(tipo_norm) or "destinacao"

        uf  = safe_str(row.get("ES", row.get("UF", row.get("Estado",""))))
        nf  = safe_str(row.get("N° da NF", row.get("NF",""))) or None
        if nf and nf.lower() in ("nan","none","0","-"): nf = None

        valor = safe_float(row.get("Valor (R$)", row.get("Valor", row.get("valor (R$)", 0))))

        data_emissao = fmt_data(row.get("Data de Emissão","")) or ""

        entry = {
            "id":          50000 + i,
            "data":        data_emissao,        # filtro por mês no dashboard
            "emissao":     data_emissao,
            "vencimento":  fmt_data(row.get("Vencimento","")) or "",
            "empresa":     emp,
            "empresaNome": emp,                 # exibição no dashboard
            "empresaId":   emp_id,              # vínculo com a empresa para somar faturamento
            "cnpj":        safe_str(row.get("CNPJ","")),
            "contato":     safe_str(row.get("Contato","")),
            "email":       safe_str(row.get("E-mail", row.get("Email",""))),
            "telefone":    safe_str(row.get("Telefone","")),
            "cidade":      safe_str(row.get("Cidade","")),
            "uf":          uf,
            "tipo":        tipo_key,
            "descricao":   safe_str(row.get("Descrição", row.get("Descricao",""))),
            "nf":          nf,
            "valor":       valor,
            "status":      safe_str(row.get("Status","Aguard. Aprovação")) or "Aguard. Aprovação",
        }
        fat["servicos"][tipo_key].append(entry)
        if emp not in ev:
            ev.add(emp)
            fat["empresas"].append({
                "id":      emp_id,   # mesmo id usado nas NFs
                "nome":    emp,
                "cnpj":    entry["cnpj"],
                "contato": entry["contato"],
                "email":   entry["email"],
                "telefone":entry["telefone"],
                "cidade":  entry["cidade"],
                "uf":      uf
            })

    nfs = sum(len(v) for v in fat["servicos"].values())
    print(f"  Faturamento: {nfs} NFs | {len(fat['empresas'])} empresas", flush=True)
    for tipo, lst in fat["servicos"].items():
        if lst: print(f"    {tipo}: {len(lst)} NFs", flush=True)

    # Amostra para confirmar campos
    for tipo, lst in fat["servicos"].items():
        if lst:
            ex = lst[0]
            print(f"  Exemplo {tipo}: data='{ex['data']}' valor={ex['valor']} empresaNome='{ex['empresaNome']}'", flush=True)
            break
    return fat

# ── CARREGAR ───────────────────────────────────────────────────

def carregar_dados():
    print("\n"+"="*50, flush=True)
    print(f"Carregando dados — {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}", flush=True)
    print("="*50, flush=True)
    try:
        gc = get_gc()
        _cache["ds"]  = processar_residuos(gc)
        _cache["fat"] = processar_faturamento(gc)
        _cache["ultima_atualizacao"] = datetime.now().strftime("%d/%m/%Y %H:%M")
        meses = len(_cache["ds"])
        regs  = sum(len(v.get("clientes",[])) for v in _cache["ds"].values())
        nfs   = sum(len(v) for v in _cache["fat"]["servicos"].values())
        print(f"\n✅ Dados: {meses} meses | {regs} registros | {nfs} NFs", flush=True)
    except Exception as e:
        import traceback
        print(f"\nERRO: {e}", flush=True)
        print(traceback.format_exc(), flush=True)

# ── LOGIN HTML ─────────────────────────────────────────────────

def get_login_html(erro=False):
    err_style = "display:block" if erro else "display:none"
    return f"""<!DOCTYPE html>
<html lang="pt-BR"><head><meta charset="UTF-8"><title>Gestão de Resíduos</title>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:'Segoe UI',sans-serif;background:linear-gradient(160deg,#1b4332,#0d2618);min-height:100vh;display:flex;align-items:center;justify-content:center}}
.box{{background:rgba(255,255,255,.08);border:1px solid rgba(255,255,255,.13);backdrop-filter:blur(10px);border-radius:20px;padding:36px 32px;width:340px;text-align:center}}
h1{{color:#fff;font-size:26px;font-weight:700;margin-bottom:6px}}
.sub{{color:#95d5b2;font-size:12px;margin-bottom:28px}}
.row{{display:flex;gap:10px;margin-bottom:20px}}
.p{{flex:1;padding:12px 8px;background:rgba(255,255,255,.06);border:1.5px solid rgba(255,255,255,.15);border-radius:12px;color:rgba(255,255,255,.6);cursor:pointer;font-size:13px;display:flex;flex-direction:column;align-items:center;gap:3px;transition:.2s;font-family:inherit}}
.p:hover,.p.on{{background:rgba(255,255,255,.12);border-color:#52b788;color:#fff}}
.pi{{font-size:20px}}.pl{{font-weight:700}}
input[type=password]{{width:100%;padding:12px 16px;border:1px solid rgba(255,255,255,.2);background:rgba(255,255,255,.08);color:#fff;border-radius:10px;font-size:14px;font-family:inherit;outline:none;margin-bottom:12px;transition:.2s}}
input:focus{{border-color:#52b788;background:rgba(255,255,255,.13)}}
input::placeholder{{color:rgba(255,255,255,.4)}}
.btn{{width:100%;padding:13px;background:#40916c;color:#fff;border:none;border-radius:10px;font-weight:700;font-size:15px;cursor:pointer;font-family:inherit;transition:.2s}}
.btn:hover{{background:#52b788}}
.err{{color:#fca5a5;font-size:12px;margin-top:10px;{err_style}}}
</style></head><body>
<div class="box">
  <h1>♻️ Gestão de Resíduos</h1>
  <div class="sub">Sistema de Gerenciamento Ambiental</div>
  <form method="POST" action="/login">
    <div class="row">
      <button type="button" class="p on" id="ba" onclick="sel('admin')"><span class="pi">🔑</span><span class="pl">Admin</span></button>
      <button type="button" class="p" id="bv" onclick="sel('viewer')"><span class="pi">👁️</span><span class="pl">Visualizador</span></button>
    </div>
    <input type="hidden" name="perfil" id="perfil" value="admin">
    <input type="password" name="senha" placeholder="Senha de acesso..." autofocus>
    <button type="submit" class="btn">Entrar</button>
    <div class="err">Senha incorreta</div>
  </form>
</div>
<script>
function sel(p){{
  document.getElementById('perfil').value=p;
  document.getElementById('ba').classList.toggle('on',p==='admin');
  document.getElementById('bv').classList.toggle('on',p==='viewer');
  document.querySelector('input[type=password]').value='';
  document.querySelector('input[type=password]').focus();
}}
</script></body></html>"""

# ── DASHBOARD ──────────────────────────────────────────────────

def serve_dashboard(perfil):
    html_path = os.path.join(PASTA, "dashboard.html")
    if not os.path.exists(html_path):
        return Response("<h2>dashboard.html nao encontrado</h2>", status=500, mimetype="text/html")
    with open(html_path, encoding="utf-8") as f:
        html = f.read()

    for lib in ["chart.js","html2pdf.js","xlsx.js"]:
        local = os.path.join(PASTA, lib)
        if os.path.exists(local) and os.path.getsize(local) > 10000:
            cdn_map = {
                "chart.js":    "https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.js",
                "html2pdf.js": "https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js",
                "xlsx.js":     "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js",
            }
            html = html.replace(cdn_map[lib], f"/static/{lib}")

    dados_json = json.dumps({
        "ds":  _cache["ds"],
        "fat": _cache["fat"],
        "ultima_atualizacao": _cache["ultima_atualizacao"],
    }, ensure_ascii=False, separators=(",",":"))
    ro = "true" if perfil == "viewer" else "false"

    inject = f"""<script>
window.__DADOS_SERVIDOR__ = {dados_json};
window.__PERFIL__ = "{perfil}";
window.__RO__ = {ro};
document.addEventListener('DOMContentLoaded', function() {{
    try {{ localStorage.removeItem('gr_ds'); }} catch(e) {{}}
    try {{ localStorage.removeItem('gr_fat'); }} catch(e) {{}}
    var ls = document.getElementById('login-screen');
    var ap = document.getElementById('app');
    if (ls) ls.style.display = 'none';
    if (ap) ap.style.display = 'block';
    if (typeof boot === 'function') boot();
}});
</script>"""
    html = html.replace("</head>", inject + "\n</head>", 1)
    return Response(html, mimetype="text/html; charset=utf-8")

# ── ROTAS ──────────────────────────────────────────────────────

@app.route("/")
def index():
    return Response(get_login_html(bool(request.args.get("erro"))), mimetype="text/html; charset=utf-8")

@app.route("/login", methods=["POST"])
def login():
    senha  = request.form.get("senha","")
    perfil = request.form.get("perfil","admin")
    h = hashlib.sha256(senha.encode()).hexdigest()
    if perfil == "admin"  and h == HASH_ADMIN:  return serve_dashboard("admin")
    if perfil == "viewer" and h == HASH_VIEWER: return serve_dashboard("viewer")
    return redirect("/?erro=1")

@app.route("/api/atualizar", methods=["POST"])
def api_atualizar():
    carregar_dados()
    regs = sum(len(v.get("clientes",[])) for v in _cache["ds"].values())
    nfs  = sum(len(v) for v in _cache["fat"]["servicos"].values())
    return Response(json.dumps({"ok":True,"msg":f"Atualizado em {_cache['ultima_atualizacao']} — {regs} registros | {nfs} NFs"}, ensure_ascii=False), mimetype="application/json")

@app.route("/api/status")
def api_status():
    meses = len(_cache["ds"])
    regs  = sum(len(v.get("clientes",[])) for v in _cache["ds"].values())
    nfs   = sum(len(v) for v in _cache["fat"]["servicos"].values())
    return Response(json.dumps({"ok":True,"meses":meses,"registros":regs,"nfs":nfs,"ultima_atualizacao":_cache["ultima_atualizacao"]}, ensure_ascii=False), mimetype="application/json")

@app.route("/api/debug-fat")
def debug_fat():
    todas = []
    for tipo, lista in _cache["fat"]["servicos"].items():
        for item in lista[:2]:
            todas.append({**item, "_tipo_key": tipo})
    return Response(json.dumps({
        "total_nfs": sum(len(v) for v in _cache["fat"]["servicos"].values()),
        "total_empresas": len(_cache["fat"]["empresas"]),
        "amostra_nfs": todas[:6],
        "amostra_empresas": _cache["fat"]["empresas"][:3],
    }, ensure_ascii=False, indent=2), mimetype="application/json")

@app.route("/static/<path:filename>")
def static_files(filename):
    path = os.path.join(PASTA, filename)
    if os.path.exists(path):
        ext  = filename.rsplit(".",1)[-1]
        mime = {"js":"application/javascript","css":"text/css"}.get(ext,"text/plain")
        with open(path,"rb") as f: data = f.read()
        return Response(data, mimetype=mime)
    return Response("Not found", status=404)

@app.route("/baixar-libs")
def baixar_libs():
    import requests as req
    libs = {
        "chart.js":    "https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.js",
        "html2pdf.js": "https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js",
        "xlsx.js":     "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js",
    }
    resultados = []
    for nome, url in libs.items():
        dest = os.path.join(PASTA, nome)
        try:
            r = req.get(url, timeout=30)
            with open(dest,"wb") as f: f.write(r.content)
            resultados.append(f"OK {nome}: {len(r.content):,} bytes")
        except Exception as e:
            resultados.append(f"ERRO {nome}: {e}")
    return Response("<br>".join(resultados)+"<br><br><a href='/'>Voltar</a>", mimetype="text/html")

# ── INÍCIO ─────────────────────────────────────────────────────

def abrir_navegador():
    time.sleep(2)
    webbrowser.open(f"http://localhost:{PORTA}")

carregar_dados()

if __name__ == "__main__":
    print(f"\nServidor: http://localhost:{PORTA}", flush=True)
    print("Ctrl+C para encerrar\n", flush=True)
    # threading.Thread(target=abrir_navegador, daemon=True).start()  # desativado em nuvem
    app.run(host="0.0.0.0", port=PORTA, debug=False, use_reloader=False)
