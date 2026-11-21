# -*- coding: utf-8 -*-
"""Atualiza protocolados.html (aba Protocolados, visão semanal) a partir da exportação da view de agendas.

Fonte: v_pc_agendas_processos*.xlsx mais recente nesta pasta (área TRABALHISTA, sem as tarefas automáticas quem_incluiu = PRAZO).
Status pelas datas (regra da usuária, 21/11/2026):
  baixado no dia da agenda = D-1 · na véspera = D-2 · antes = Futuro · no dia útil seguinte = Fatal · depois = Vencido.
- Prazos: por responsável e grupo do processo (Grupo B4 soma no B3). Fora do horário = antes das 9h ou a partir das 19h.
- Agendas e Agenda Equipe: por quem realizou, só advogados e sócios (= quem é responsável por prazos na exportação), equipe da exportação;
  "atrasada" = Vencido. Conferido contra a S44 publicada antes de gravar (ver --conferir).
Semanas anteriores à nova ficam como estão; só entram as semanas novas nos históricos (resumo semanal, mensal, tipos).

Uso:
  python _gerar_semanal.py --conferir            recalcula a S44 e compara com o publicado (não grava)
  python _gerar_semanal.py 2026-11-16 2026-11-20 recalcula e grava a semana (as semanas entre a última publicada e esta entram no histórico)
"""
import sys, re, json, shutil, subprocess, tempfile, datetime as dt, collections
from pathlib import Path
import numpy as np
import openpyxl

AQUI = Path(__file__).resolve().parent
HTML = AQUI / "protocolados.html"
FERIADOS = ["2026-01-01", "2026-02-16", "2026-02-17", "2026-04-03", "2026-04-21", "2026-05-01", "2026-06-04",
            "2026-09-07", "2026-10-12", "2026-11-02", "2026-11-20", "2026-12-25"]
DIAS = ["segunda-feira", "terça-feira", "quarta-feira", "quinta-feira", "sexta-feira", "sábado", "domingo"]
BS = chr(92)


def fmt(n): return f"{n:,}".replace(",", ".")


def classifica(da, dr):
    if dr == da: return "d1"
    if dr < da: return "d2" if (da - dr).days == 1 else "fut"
    return "fat" if np.busday_count(da, dr, holidays=FERIADOS) <= 1 else "ven"


def hora(h):
    m = re.match(r"(\d{1,2}):", str(h or ""))
    return int(m[1]) if m else None


def fora(h):
    x = hora(h)
    return x is not None and (x < 9 or x >= 19)


# ---------------------------------------------------------------- leitura
def dat(v):
    if isinstance(v, dt.datetime): return v.date()
    if isinstance(v, (int, float)) and 30000 < v < 60000: return (dt.datetime(1899, 12, 30) + dt.timedelta(days=int(v))).date()
    m = re.match(r"(\d{2})/(\d{2})/(\d{4})", str(v or "").strip())
    return dt.date(int(m[3]), int(m[2]), int(m[1])) if m else None


def ler_exportacao(desde):
    arq = sorted(AQUI.glob("v_pc_agendas_processos*.xlsx"), key=lambda f: f.stat().st_mtime)[-1]
    try:
        wb = openpyxl.load_workbook(arq, read_only=True)
    except PermissionError:
        tmp = Path(tempfile.gettempdir()) / "agendas_export_semanal.xlsx"
        subprocess.run(["powershell", "-NoProfile", "-Command", f"Copy-Item -LiteralPath '{arq}' -Destination '{tmp}' -Force"], check=True)
        wb = openpyxl.load_workbook(tmp, read_only=True)
    rows, vistos = [], set()
    for ws in wb:
        cab = None
        for r in ws.iter_rows(values_only=True):
            if cab is None:
                cab = {str(c).strip(): i for i, c in enumerate(r) if c}
                if not {"tipo_agenda", "status", "data_realizacao", "data_agenda", "area"} <= set(cab): break
                continue
            g = lambda k: r[cab[k]]
            if str(g("status")).strip() != "Realizado" or str(g("area")).strip().upper() != "TRABALHISTA": continue
            if str(g("quem_incluiu") or "").strip().upper() == "PRAZO": continue
            tipo = str(g("tipo_agenda")).strip()
            if tipo not in ("Prazo", "Agenda", "Agenda Equipe"): continue
            da, dr = dat(g("data_agenda")), dat(g("data_realizacao"))
            if not da or not dr or dr < desde: continue
            k = (g("codigo_processo"), str(g("responsavel")).strip(), da, dr, str(g("hora_realizacao")), str(g("descricao")).strip(), tipo)
            if k in vistos: continue
            vistos.add(k)
            grupo = str(g("grupo_processo") or "").strip()
            rows.append(dict(equipe=str(g("equipe") or "").strip(), tipo=tipo, proc=str(g("numero_processo") or ""), resp=str(g("responsavel")).strip().upper(),
                             quem=str(g("quem_realizou") or "").strip().upper(), grupo="Grupo B3" if grupo == "Grupo B4" else grupo,
                             da=da, dr=dr, hr=str(g("hora_realizacao") or "").strip(), desc=str(g("descricao") or "").replace("_x000D_", "").strip(),
                             st=classifica(da, dr)))
    print(f"{arq.name}: {len(rows)} realizados desde {desde:%d/%m/%Y}")
    return rows


def cadastro_da_exportacao(rows):
    """Advogado = quem é responsável por prazos na exportação (o cadastro em Profissionais trabalhistas.xlsx está desatualizado);
    equipe = a mais frequente nas linhas da pessoa."""
    eq = collections.defaultdict(collections.Counter)
    for r in rows:
        if r["tipo"] == "Prazo" and r["equipe"]: eq[r["resp"]][r["equipe"]] += 1
    cad = cadastro()  # quem o cadastro marca como administrativo fica de fora, salvo confirmação da usuária
    ADV_CONFIRMADOS = {"GVM"}  # Giulia Vendramini: advogada (cadastro diz assistente), confirmado em 17/11/2026
    return {n: (c.most_common(1)[0][0], n in ADV_CONFIRMADOS or n not in cad or cad[n][1]) for n, c in eq.items()}


def cadastro():
    """sigla -> (equipe, é advogado ou sócio)"""
    wb = openpyxl.load_workbook(AQUI / "Profissionais trabalhistas.xlsx", read_only=True)
    it = wb["Planilha2"].iter_rows(values_only=True)
    cab = [str(c or "").strip().lower() for c in next(it)]
    i = {c: cab.index(c) for c in ("matricula", "cargo", "nomenclatura", "equipe")}
    out = {}
    for r in it:
        if not r[i["matricula"]]: continue
        cargo = str(r[i["cargo"]] or "").upper()
        adv = str(r[i["nomenclatura"]] or "").strip().lower().startswith("advog") or "SÓCIO" in cargo or "SOCIO" in cargo
        out[str(r[i["matricula"]]).strip().upper()] = (str(r[i["equipe"]] or "").strip(), adv)
    return out


# ---------------------------------------------------------------- cálculo de uma semana
def semana(rows, ini, fim, CAD):
    W = [r for r in rows if ini <= r["dr"] <= fim]
    P = [r for r in W if r["tipo"] == "Prazo"]
    out = {"ini": ini, "fim": fim}
    pr = collections.defaultdict(collections.Counter)
    for r in P:
        if not r["grupo"]: continue
        c = pr[(r["resp"], r["grupo"])]
        c[r["st"]] += 1; c["tot"] += 1
        if fora(r["hr"]): c["fora"] += 1
        if fora(r["hr"]) and r["st"] == "fat": c["ff"] += 1
    resp = [{"n": n, "grp": g, "ven": c["ven"], "fat": c["fat"], "d1": c["d1"], "d2": c["d2"], "fut": c["fut"], "tot": c["tot"], "fora": c["fora"], "ff": c["ff"]}
            for (n, g), c in pr.items()]
    out["RESP"] = sorted(resp, key=lambda x: -x["tot"])
    bg = collections.defaultdict(collections.Counter)
    for r in P:
        if not r["grupo"]: continue
        c = bg[r["grupo"]]; c[r["st"]] += 1; c["total"] += 1
        if fora(r["hr"]): c["fora"] += 1
        if fora(r["hr"]) and r["st"] == "fat": c["ff"] += 1
    out["BY_GRUPO"] = {g: {k: c[k] for k in ("total", "ven", "fat", "d1", "d2", "fut", "fora", "ff")} for g, c in sorted(bg.items(), key=lambda x: -x[1]["total"])}
    hd = collections.Counter(hora(r["hr"]) for r in P)
    out["HR_DIST"] = [[h, hd.get(h, 0)] for h in range(24)]
    out["FAT_FORA"] = [{"proc": r["proc"], "resp": r["resp"], "grp": r["grupo"], "dt": f"{r['dr']:%d/%m}", "hora_real": r["hr"], "status": "Fatal"}
                       for r in sorted(P, key=lambda r: (r["dr"], r["hr"])) if r["st"] == "fat" and fora(r["hr"])]
    tipos = collections.Counter(r["desc"] for r in P if r["desc"])
    out["TIPOS"] = dict(tipos.most_common())
    for tipo, pre in (("Agenda", "AG"), ("Agenda Equipe", "AGEQ")):
        A = [r for r in W if r["tipo"] == tipo and CAD.get(r["quem"], ("", False))[1]]
        cr, ce = collections.defaultdict(collections.Counter), collections.defaultdict(collections.Counter)
        for r in A:
            eq = CAD[r["quem"]][0]
            for c in (cr[(r["quem"], eq)], ce[eq]):
                c["tot"] += 1; c[{"ven": "atrasada", "fat": "fatal", "d1": "d1", "d2": "d2", "fut": "futuro"}[r["st"]]] += 1
        k = ("tot", "atrasada", "fatal", "d1", "d2", "futuro")
        out[pre + "_RESP"] = sorted(({"n": n, "grp": e, **{x: c[x] for x in k}} for (n, e), c in cr.items()), key=lambda x: -x["tot"])
        out[pre + "_EQUIPE"] = sorted(({"equipe": e, **{x: c[x] for x in k}} for e, c in ce.items()), key=lambda x: -x["tot"])
        out[pre + "_N"] = len(A)
    st = collections.Counter(r["st"] for r in P)
    out["ACUM"] = {"Prazo": len(P), "Agenda": out["AG_N"], "Agenda Equipe": out["AGEQ_N"], "Vencido": st["ven"], "Fatal": st["fat"],
                   "D-1": st["d1"], "D-2": st["d2"], "Futuro": st["fut"], "Total": len(P) + out["AG_N"] + out["AGEQ_N"]}
    dias = {}
    for d in sorted({r["dr"] for r in W}):
        Pd = [r for r in P if r["dr"] == d]
        s = collections.Counter(r["st"] for r in Pd)
        ag = sum(1 for r in W if r["dr"] == d and r["tipo"] == "Agenda" and CAD.get(r["quem"], ("", False))[1])
        aq = sum(1 for r in W if r["dr"] == d and r["tipo"] == "Agenda Equipe" and CAD.get(r["quem"], ("", False))[1])
        dias[d] = {"Agenda": ag, "Agenda Equipe": aq, "Prazo": len(Pd), "Total": len(Pd) + ag + aq, "Vencido": s["ven"], "Fatal": s["fat"],
                   "D-1": s["d1"], "D-2": s["d2"], "Futuro": s["fut"]}
    out["DIAS"] = dias
    return out


# ---------------------------------------------------------------- edição do HTML
class Pagina:
    def __init__(self, t): self.t = t

    def faixa(self, nome):
        m = re.search(r"const " + nome + r"\s*=\s*", self.t)
        i = m.end(); dep = 0; instr = False
        for j in range(i, len(self.t)):
            ch = self.t[j]
            if ch == '"' and self.t[j - 1] != BS: instr = not instr
            if instr: continue
            if ch in "{[": dep += 1
            elif ch in "}]":
                dep -= 1
                if dep == 0: return i, j + 1

    def get(self, nome):
        i, j = self.faixa(nome); return json.loads(self.t[i:j])

    def put(self, nome, valor):
        i, j = self.faixa(nome); self.t = self.t[:i] + json.dumps(valor, ensure_ascii=False) + self.t[j:]

    def troca(self, a, b, n=1):
        c = self.t.count(a)
        if c != n: raise SystemExit(f"esperava {n} ocorrência(s), achei {c}: {a[:90]!r}")
        self.t = self.t.replace(a, b)


def conferir(rows, CAD):
    pg = Pagina(HTML.read_text(encoding="utf-8"))
    s = semana(rows, dt.date(2026, 10, 26), dt.date(2026, 11, 2), CAD)
    pub = pg.get("ACUMULADO_DATA")["semana"]["2026-S44"]
    print("S44 publicado:", pub); print("S44 calculado:", s["ACUM"])
    pa = {r["n"]: r["tot"] for r in pg.get("AG_RESP")}
    print("Agenda por realizador (pub x calc):", [(r["n"], pa.get(r["n"]), r["tot"]) for r in s["AG_RESP"][:8]])
    print("HR_DIST igual:", pg.get("HR_DIST") == s["HR_DIST"])


def gravar(rows, ini, fim, CAD):
    pg = Pagina(HTML.read_text(encoding="utf-8"))
    ac = pg.get("ACUMULADO_DATA")
    ult = max(ac["semana"], key=lambda k: int(k.split("S")[1]))
    num_ult = int(ult.split("S")[1])
    # semanas novas: da segunda-feira seguinte ao fim da última publicada (S44 terminou em 02/11, dia útil) até a semana pedida
    labels = pg.t[pg.t.index("const SEMANA_LABELS_MAP"):pg.t.index("}", pg.t.index("const SEMANA_LABELS_MAP"))]
    fim_ult = re.search(r"'" + ult + r"':'S\d+<br><small[^>]*>([\d/]+)-([\d/]+)</small>'", labels).group(2)
    d0 = dt.date(fim.year, int(fim_ult.split("/")[1]), int(fim_ult.split("/")[0])) + dt.timedelta(days=1)
    semanas = []
    a = d0
    while a <= fim:
        b = min(a + dt.timedelta(days=(6 - a.weekday())), fim)
        semanas.append((a, b)); a = b + dt.timedelta(days=1)
    calc = [semana(rows, a, b, CAD) for a, b in semanas]
    for i, s in enumerate(calc): s["num"] = num_ult + 1 + i; s["key"] = f"{fim.year}-S{s['num']}"
    atual, ant = calc[-1], (calc[-2] if len(calc) > 1 else None)
    for s in calc: print(f"S{s['num']} {s['ini']:%d/%m}-{s['fim']:%d/%m}:", s["ACUM"])
    shutil.copy(HTML, AQUI / "Anteriores" / f"protocolados (backup pre-semana{atual['num']}).html")

    # ---- dados da semana atual
    pg.put("RESP", atual["RESP"]); pg.put("GRP_RESP", atual["RESP"]); pg.put("BY_GRUPO", atual["BY_GRUPO"])
    pg.put("HR_DIST", atual["HR_DIST"]); pg.put("FAT_FORA", atual["FAT_FORA"])
    for k in ("AG_EQUIPE", "AG_RESP", "AGEQ_EQUIPE", "AGEQ_RESP"): pg.put(k, atual[k])
    if ant: pg.put("RESP_S20", ant["RESP"])
    pg.put("SOCIOS_TOTAL_PLANILHA", {g: v["total"] for g, v in atual["BY_GRUPO"].items()})
    # ---- históricos: só acrescenta as semanas novas
    for s in calc: ac["semana"][s["key"]] = s["ACUM"]
    nov = semana(rows, dt.date(fim.year, fim.month, 1), fim, CAD)
    mkey = f"{fim.year}-{fim.month:02d}"
    ac["mes"][mkey] = nov["ACUM"]
    pg.put("ACUMULADO_DATA", ac)
    tl = pg.get("TIMELINE_DATA")
    for d, v in nov["DIAS"].items():
        if d >= d0: tl[f"{d:%Y-%m-%d}"] = v
    pg.put("TIMELINE_DATA", dict(sorted(tl.items())))
    i, j = pg.faixa("TIPOS_DATA") if re.search(r"const TIPOS_DATA\s*=", pg.t) else (None, None)
    if i is None:
        m = re.search(r"TIPOS_DATA\s*=\s*", pg.t); i = m.end()
        dep = 0; instr = False
        for j in range(i, len(pg.t)):
            ch = pg.t[j]
            if ch == '"' and pg.t[j - 1] != BS: instr = not instr
            if instr: continue
            if ch == "{": dep += 1
            elif ch == "}":
                dep -= 1
                if dep == 0: j += 1; break
    td = json.loads(pg.t[i:j])
    for s in calc: td[s["key"]] = s["TIPOS"]
    pg.t = pg.t[:i] + json.dumps(td, ensure_ascii=False, indent=4) + pg.t[j:]
    novos_lbl = "".join(f",\n  '{s['key']}':'S{s['num']}<br><small style=\"color:var(--tx2)\">{s['ini']:%d/%m}-{s['fim']:%d/%m}</small>'" for s in calc)
    m = re.search(r"('" + ult + r"':'[^']*')", pg.t); pg.t = pg.t[:m.end()] + novos_lbl + pg.t[m.end():]

    # ---- textos fixos
    A, P0 = atual["ACUM"], f"{atual['ini']:%d/%m/%Y} a {atual['fim']:%d/%m/%Y}"
    velho = re.search(r"📅 Período apurado: ([\d/]+) a ([\d/]+)", pg.t)
    v_ini, v_fim = velho.group(1), velho.group(2)
    tot = A["Prazo"]; fat = A["Fatal"]; ven = A["Vencido"]; ant_ = A["D-2"] + A["Futuro"]
    fora_n = sum(r["fora"] for r in atual["RESP"]); ff = sum(r["ff"] for r in atual["RESP"])
    kp = re.search(r'<div class="kpis">\s*<div class="kpi" onclick="kpiClick\(this,\'\'\)">.*?</div>\s*</div>\s*\n</div>', pg.t, re.S)
    novo_kp = f'''<div class="kpis">
  <div class="kpi" onclick="kpiClick(this,'')"><div class="klbl">Total Baixado</div><div class="kval wh">{fmt(tot)}</div><div class="ksub">prazos realizados</div></div>
  <div class="kpi" onclick="kpiClick(this,'fatal')" id="kpi-fatal"><div class="klbl">Fatais</div><div class="kval" style="color:#e74c3c">{fmt(fat)}</div><div class="ksub">{100*fat/tot:.1f}% — protocolados no dia fatal</div></div>
  <div class="kpi" onclick="kpiClick(this,'vencido')" id="kpi-vencido"><div class="klbl">Vencidos</div><div class="kval" style="color:#c0392b">{fmt(ven)}</div><div class="ksub">{100*ven/tot:.1f}% — prazos vencidos</div></div>
  <div class="kpi" onclick="kpiClick(this,'d1')" id="kpi-d1"><div class="klbl">D-1 (no dia)</div><div class="kval yel">{fmt(A["D-1"])}</div><div class="ksub">cumpridos no prazo</div></div>
  <div class="kpi" onclick="kpiClick(this,'antecipado')" id="kpi-antecipado"><div class="klbl">Antecipados</div><div class="kval grn">{fmt(ant_)}</div><div class="ksub">{100*ant_/tot:.1f}% — antes do prazo</div></div>
  <div class="kpi" onclick="toggleFilter('fora')" id="kpi-fora"><div class="klbl">Fora do Horário</div><div class="kval org">{fmt(fora_n)}</div><div class="ksub">antes 9h ou após 19h</div></div>
  <div class="kpi" onclick="toggleFilter('fatfora')" id="kpi-fatfora"><div class="klbl">Fatal + Fora Hora</div><div class="kval" style="color:#e74c3c">{fmt(ff)}</div><div class="ksub">⚠ caso fatal — clique para filtrar</div></div>
</div>'''.replace(".", ",", 0)
    novo_kp = re.sub(r"(\d)\.(\d)%", r"\1,\2%", novo_kp)
    pg.t = pg.t[:kp.start()] + novo_kp + pg.t[kp.end():]
    nota = re.search(r"<span><b>Nota:</b> No período de <b>[^<]*</b>.*?</span>", pg.t, re.S)
    if ven == 0:
        txt = f"<span><b>Nota:</b> No período de <b>{P0}</b> não houve <b>nenhum prazo vencido</b>, em um total de {fmt(tot)} prazos realizados na semana. Status apurado pelas datas da agenda e da baixa.</span>"
    else:
        txt = f"<span><b>Nota:</b> No período de <b>{P0}</b> houve <b>{fmt(ven)} prazo(s) vencido(s)</b>, em um total de {fmt(tot)} prazos realizados na semana. Status apurado pelas datas da agenda e da baixa.</span>"
    pg.t = pg.t[:nota.start()] + txt + pg.t[nota.end():]
    pg.troca(f"📅 Período apurado: {v_ini} a {v_fim}", f"📅 Período apurado: {P0}")
    pg.troca(f"📅 Agendas — {v_ini} a {v_fim}</h2>", f"📅 Agendas — {P0}</h2>")
    pg.troca(f"👥 Agenda Equipe — {v_ini} a {v_fim}</h2>", f"👥 Agenda Equipe — {P0}</h2>")
    pg.t = re.sub(r'(📅 Agendas — ' + re.escape(P0) + r'</h2>\s*<span[^>]*>)[\d.]+ realizadas', lambda m: m.group(1) + f"{fmt(atual['AG_N'])} realizadas", pg.t)
    pg.t = re.sub(r'(👥 Agenda Equipe — ' + re.escape(P0) + r'</h2>\s*<span[^>]*>)[\d.]+ realizadas', lambda m: m.group(1) + f"{fmt(atual['AGEQ_N'])} realizadas", pg.t)
    # comparativo semana anterior x atual
    if ant:
        B = ant["ACUM"]
        def lin(nome, k, cls, melhor_menor):
            a, b = B[k], A[k]; d = b - a; p = f"{d/a*100:+.0f}%" if a else "—"
            bom = (d <= 0) if melhor_menor else (d >= 0)
            ev = '<span class="evolucao melhorando">✓ MELHORANDO</span>' if bom else '<span class="evolucao piorando">✗ PIORANDO</span>'
            if d == 0: ev = '<span class="evolucao estavel">= ESTÁVEL</span>'
            return (f'      <tr>\n        <td><b>{nome}</b></td>\n        <td class="c"><span class="vbad {cls}">{fmt(a)}</span></td>\n        <td class="c"><span class="vbad {cls}">{fmt(b)}</span></td>\n'
                    f'        <td class="c"><span class="vbad {"grn" if bom else "red"}">{d:+d} ({p})</span></td>\n        <td class="c">{ev}</td>\n      </tr>')
        corpo = "\n".join([lin("Vencido", "Vencido", "red", True), lin("Fatal", "Fatal", "fat", True), lin("D-1", "D-1", "yel", False), lin("D-2", "D-2", "blu", False), lin("Futuro", "Futuro", "grn", False)])
        cp = re.search(r'📊 Comparativo de Evolução — Semana \d+ vs Semana \d+</div>\s*<table style="margin:0">.*?</table>', pg.t, re.S)
        novo = (f'📊 Comparativo de Evolução — Semana {ant["num"]} vs Semana {atual["num"]}</div>\n  <table style="margin:0">\n    <thead><tr><th>Métrica</th>'
                f'<th class="c">Semana {ant["num"]}<br>({ant["ini"]:%d/%m}-{ant["fim"]:%d/%m})</th><th class="c">Semana {atual["num"]}<br>({atual["ini"]:%d/%m}-{atual["fim"]:%d/%m})</th>'
                f'<th class="c">Variação</th><th class="c">Evolução</th></tr></thead>\n    <tbody>\n{corpo}\n    </tbody>\n  </table>')
        pg.t = pg.t[:cp.start()] + novo + pg.t[cp.end():]
        sA, sB = f"S{atual['num']} ({atual['ini']:%d/%m}–{atual['fim']:%d/%m})", f"S{ant['num']} ({ant['ini']:%d/%m}–{ant['fim']:%d/%m})"
        pg.t = re.sub(r"S\d+ \(\d\d/\d\d–\d\d/\d\d\) vs S\d+ \(\d\d/\d\d–\d\d/\d\d\)", f"{sB} vs {sA}", pg.t)
        pg.t = re.sub(r"a linha clara é a semana anterior \(S\d+ · [\d/]+ a [\d/]+\) e a linha normal é a semana atual \(S\d+ · [\d/]+ a [\d/]+\)",
                      f"a linha clara é a semana anterior (S{ant['num']} · {ant['ini']:%d/%m} a {ant['fim']:%d/%m}) e a linha normal é a semana atual (S{atual['num']} · {atual['ini']:%d/%m} a {atual['fim']:%d/%m})", pg.t)
        pg.t = re.sub(r"// Constrói mapa S\d+ como semana anterior \([^)]*\)", f"// Constrói mapa S{ant['num']} como semana anterior ({ant['ini']:%d/%m}–{ant['fim']:%d/%m})", pg.t)
        pg.t = re.sub(r"// Constrói mapa S\d+ \(semana atual: [^)]*\)", f"// Constrói mapa S{atual['num']} (semana atual: {atual['ini']:%d/%m}–{atual['fim']:%d/%m})", pg.t)
        pg.t = re.sub(r'(<td><span style="font-size:10px;color:var\(--tx2\)">)\d\d/\d\d a \d\d/\d\d(</span></td>)', lambda m: m.group(1) + f"{ant['ini']:%d/%m} a {ant['fim']:%d/%m}" + m.group(2), pg.t)
    pg.t = re.sub(r"(<td><b>)\d\d/\d\d a \d\d/\d\d(</b></td>\s*<td class=\"c\"><span class=\"vbad muted\">\$\{fmtNum\(d21)", lambda m: m.group(1) + f"{atual['ini']:%d/%m} a {atual['fim']:%d/%m}" + m.group(2), pg.t)
    pg.t = re.sub(r'<div class="hlabel">Prazos S\d+</div><div class="vbad muted">\$\{fmtNum\(\d+\)\}', f'<div class="hlabel">Prazos S{atual["num"]}</div><div class="vbad muted">${{fmtNum({tot})}}', pg.t)
    pg.t = re.sub(r"\d\d/\d\d a \d\d/\d\d/\d{4} · resumo consolidado por grupo", f"{atual['ini']:%d/%m} a {atual['fim']:%d/%m/%Y} · resumo consolidado por grupo", pg.t)
    pg.t = re.sub(r"// Totais alinhados a BY_GRUPO da semana [^\n]*", f"// Totais alinhados a BY_GRUPO da semana {P0} (S{atual['num']}).", pg.t)

    # ---- tabelas estáticas: linha do tempo (dias novos), resumo semanal (semanas novas) e mensal (mês atual)
    def cel(v, c): return f'<td class="c"><span class="vbad {c}">{fmt(v)}</span></td>'
    lin_dia = ""
    for d, v in nov["DIAS"].items():
        if d < d0: continue
        risco = round(100 * v["Vencido"] / v["Total"]) if v["Total"] else 0
        chip = "high" if risco >= 30 else "mid" if risco >= 10 else "low"
        lin_dia += (f'<tr data-timeline-date="{d:%Y-%m-%d}" onclick="openTimelineDay(\'{d:%Y-%m-%d}\')" title="Clique para ver o resumo do dia">\n'
                    f'      <td><b>{d:%d/%m/%Y}</b><span class="nome-sigla">{DIAS[d.weekday()]}</span></td>\n      '
                    + "\n      ".join([cel(v["Prazo"], "muted"), cel(v["Agenda"], "blu"), cel(v["Agenda Equipe"], "grn"), cel(v["Vencido"], "red"), cel(v["Fatal"], "fat"),
                                       cel(v["D-1"], "yel"), cel(v["D-2"], "blu"), cel(v["Futuro"], "grn")])
                    + f'\n      <td class="c" style="font-weight:800;color:var(--pc-escuro)">{fmt(v["Total"])}</td>\n      <td class="c"><span class="risk-chip {chip}">{risco}%</span></td>\n    </tr>')
    k = pg.t.index("</tbody>", pg.t.index('id="tb-timeline"'))
    pg.t = pg.t[:k] + lin_dia + "\n\n    " + pg.t[k:]

    def lin_ac(tipo, key, rot, v):
        return (f'<tr data-acumulo="{tipo}" data-acumulo-key="{key}" onclick="openAcumulado(\'{tipo}\',\'{key}\')" title="Clique para detalhar o acumulado">\n'
                f'          <td><b>{rot}</b></td>\n          ' + "\n          ".join([cel(v["Prazo"], "muted"), cel(v["Agenda"], "blu"), cel(v["Agenda Equipe"], "grn"), cel(v["Vencido"], "red"),
                                                                                    cel(v["Fatal"], "fat"), cel(v["D-1"], "yel"), cel(v["D-2"], "blu"), cel(v["Futuro"], "grn")])
                + f'\n          <td class="c" style="font-weight:800;color:var(--pc-escuro)">{fmt(v["Total"])}</td>\n        </tr>')
    sem_rows = "".join(lin_ac("semana", s["key"], f"Semana {s['num']} · {s['ini']:%d/%m} a {s['fim']:%d/%m}", s["ACUM"]) for s in calc)
    a = pg.t.index(f'data-acumulo-key="{ult}"'); e = pg.t.index("</tr>", a) + 5
    pg.t = pg.t[:e] + sem_rows + pg.t[e:]
    rot_mes = f"{fim.month:02d}/{fim.year} acumulado · 01/{fim.month:02d} a {fim:%d/%m}"
    novo_mes = lin_ac("mes", mkey, rot_mes, nov["ACUM"])
    if f'data-acumulo="mes" data-acumulo-key="{mkey}"' in pg.t:
        a = pg.t.index(f'data-acumulo="mes" data-acumulo-key="{mkey}"'); a = pg.t.rfind("<tr", 0, a); e = pg.t.index("</tr>", a) + 5
        pg.t = pg.t[:a] + novo_mes + pg.t[e:]
    else:
        k = pg.t.index("</tbody>", pg.t.index('data-acumulo="mes"'))
        pg.t = pg.t[:k] + novo_mes + "\n" + pg.t[k:]
    pg.t = re.sub(r"(<title>[^<\[]*\[SEMANA )\d+(\])", lambda m: m.group(1) + str(atual["num"]) + m.group(2), pg.t)
    t0 = re.search(r"<title>([^<]*)</title>", pg.t).group(1)
    HTML.write_text(pg.t, encoding="utf-8")
    print("gravado:", HTML.name, "| semana", atual["key"], P0, "| título", t0)


if __name__ == "__main__":
    if "--conferir" in sys.argv:
        rows = ler_exportacao(dt.date(2026, 10, 20)); conferir(rows, cadastro_da_exportacao(rows))
    else:
        ini, fim = (dt.date.fromisoformat(x) for x in sys.argv[1:3])
        rows = ler_exportacao(dt.date(fim.year, fim.month, 1) - dt.timedelta(days=10))
        gravar(rows, ini, fim, cadastro_da_exportacao(rows))
