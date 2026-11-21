# -*- coding: utf-8 -*-
"""Atualiza Protocolados por Mês, Evolução dos Prazos e Análises com os meses que ainda não estão publicados.

Como funciona
- Janeiro a outubro/2026 ficam exatamente como foram publicados (status congelado pelos semanais). Na primeira rodada
  esses números são lidos das próprias páginas e gravados em _historico_mensal.json; depois disso o histórico vem do JSON.
- Os meses seguintes vêm da exportação mais recente da view de agendas (v_pc_agendas_processos*.xlsx nesta pasta):
  área TRABALHISTA, tipo_agenda Prazo, status Realizado, sem as tarefas automáticas (quem_incluiu = PRAZO), com o status calculado pelas datas (regra confirmada em 21/11/2026):
    baixado no próprio dia da agenda ............ D-1
    baixado na véspera (dia anterior) ........... D-2
    baixado antes disso ......................... Futuro
    baixado no dia útil seguinte à agenda ....... Fatal
    baixado depois disso ........................ Vencido
  Conferido contra o status do Compilado em set e out/2026: bate em 99,7% dos prazos.
- Grupo "Grupo B4" é somado ao "Grupo B3".

Uso: python _gerar_mensal.py            (faz backup dos 3 HTML antes de gravar)
"""
import json, re, glob, os, shutil, subprocess, tempfile, datetime as dt, collections, math
from pathlib import Path
import numpy as np
import openpyxl

AQUI = Path(__file__).resolve().parent
HIST = AQUI / "_historico_mensal.json"
MESES = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
ST = ["Fatal", "D-1", "D-2", "Futuro", "Vencido"]
FERIADOS = ["2026-01-01", "2026-02-16", "2026-02-17", "2026-04-03", "2026-04-21", "2026-05-01", "2026-06-04",
            "2026-09-07", "2026-10-12", "2026-11-02", "2026-11-20", "2026-12-25"]
OVERRIDES = {"TLI": "Tamires Lins Costa", "FCG": "Fabiana Correia dos Anjos Gomes", "TAM": "Tamara Silva Marcílio",
             "BTF": "Bruno Tanganelli Farah", "LWP": "Lucas Willott Pereira", "JTR": "Jackeline Romeu Torres", "JRT": "Jackeline Romeu Torres"}


def rotulo(ano, mes): return f"{MESES[mes - 1]}/{ano}"


# ---------------------------------------------------------------- histórico publicado (jan a out/2026)
def historico_das_paginas():
    pm = (AQUI / "Protocolados_Mes.html").read_text(encoding="utf-8")
    export = json.loads(re.search(r"const EXPORT_DATA\s*=\s*(\{.*?\});\s*\n", pm, re.S).group(1))
    ev = (AQUI / "Evolucao_Prazos.html").read_text(encoding="utf-8")
    tb = re.search(r'<table class="tb" id="tblEvGrupo">(.*?)</table>', ev, re.S).group(1)
    meses = re.findall(r'<th class="cc">([^<]+)</th>', tb)[:-1]
    grupos = {}
    for tr in re.findall(r"<tr><td>(.*?)</tr>", tb.split("<tbody>")[1], re.S):
        nome = tr.split("</td>")[0]
        vals = [re.sub(r"<[^>]+>", "", v) for v in re.findall(r'<td class="cc"[^>]*>(.*?)</td>', tr)][:-1]
        grupos[nome] = {m: (0 if v.strip() in ("—", "") else int(v)) for m, v in zip(meses, vals)}
    an = (AQUI / "Analises.html").read_text(encoding="utf-8")
    rank = re.search(r"(Ranking por Grupo[^<]*)", an)
    return {"meses": list(export.keys()), "profissional": export, "fataisGrupo": grupos,
            "origem": "páginas publicadas em 03/11/2026 (status congelado pelos semanais)"}


# ---------------------------------------------------------------- meses novos (exportação da view)
def ler_exportacao():
    arqs = sorted(AQUI.glob("v_pc_agendas_processos*.xlsx"), key=lambda f: f.stat().st_mtime)
    if not arqs: raise SystemExit("Nenhuma exportação v_pc_agendas_processos*.xlsx na pasta.")
    arq = arqs[-1]
    try:
        wb = openpyxl.load_workbook(arq, read_only=True)
    except PermissionError:  # aberta no Excel
        tmp = Path(tempfile.gettempdir()) / "agendas_export_baixados.xlsx"
        subprocess.run(["powershell", "-NoProfile", "-Command", f"Copy-Item -LiteralPath '{arq}' -Destination '{tmp}' -Force"], check=True)
        wb = openpyxl.load_workbook(tmp, read_only=True)

    def dat(v):
        if isinstance(v, dt.datetime): return v.date()
        if isinstance(v, (int, float)) and 30000 < v < 60000: return (dt.datetime(1899, 12, 30) + dt.timedelta(days=int(v))).date()
        m = re.match(r"(\d{2})/(\d{2})/(\d{4})", str(v or "").strip())
        return dt.date(int(m[3]), int(m[2]), int(m[1])) if m else None

    vistos, rows = set(), []
    for ws in wb:
        cab = None
        for r in ws.iter_rows(values_only=True):
            if cab is None:
                cab = {str(c).strip(): i for i, c in enumerate(r) if c}
                if not {"tipo_agenda", "status", "responsavel", "data_realizacao", "data_agenda", "area"} <= set(cab): break
                continue
            g = lambda k: r[cab[k]]
            if str(g("tipo_agenda")).strip() != "Prazo" or str(g("status")).strip() != "Realizado": continue
            if str(g("area")).strip().upper() != "TRABALHISTA": continue
            # tarefas automáticas (incluídas pelo usuário PRAZO e baixadas por robô WMS/ASR-LM) não entram no Compilado
            if "quem_incluiu" in cab and str(g("quem_incluiu") or "").strip().upper() == "PRAZO": continue
            da, dr = dat(g("data_agenda")), dat(g("data_realizacao"))
            if not da or not dr: continue
            chave = (g("codigo_processo"), str(g("responsavel")).strip().upper(), da, dr, str(g("hora_realizacao")), str(g("descricao")).strip())
            if chave in vistos: continue
            vistos.add(chave)
            grupo = str(g("grupo_processo") or "Sem Grupo").strip() or "Sem Grupo"
            rows.append(dict(resp=str(g("responsavel")).strip().upper(), grupo="Grupo B3" if grupo == "Grupo B4" else grupo, da=da, dr=dr))
    return arq.name, rows


def classifica(da, dr):
    if dr == da: return "D-1"
    if dr < da: return "D-2" if (da - dr).days == 1 else "Futuro"
    return "Fatal" if np.busday_count(da, dr, holidays=FERIADOS) <= 1 else "Vencido"


def nomes():
    m = {}
    for arq, aba, ck, cn in (("profissionais_peixoto(editado).xlsx", None, None, None), ("Profissionais trabalhistas.xlsx", "Planilha2", "matricula", "nome")):
        f = AQUI / arq
        if not f.exists(): continue
        wb = openpyxl.load_workbook(f, read_only=True)
        ws = wb[aba] if aba else wb.worksheets[0]
        it = ws.iter_rows(values_only=True); cab = [str(c or "").strip().lower() for c in next(it)]
        ik = cab.index(ck) if ck else next((i for i, c in enumerate(cab) if c in ("matricula", "sigla", "codigo")), None)
        inn = cab.index(cn) if cn else next((i for i, c in enumerate(cab) if c == "nome"), None)
        if ik is None or inn is None: continue
        for r in it:
            if r[ik] and r[inn]: m[str(r[ik]).strip().upper()] = str(r[inn]).replace("\xa0", " ").strip()
    m.update(OVERRIDES)
    return m


# ---------------------------------------------------------------- render
def cor(v, mx):
    if not v: return "#EDF1F4", "#003B5C"
    t = math.sqrt(v / mx) if mx else 0
    a, b = (0xED, 0xF1, 0xF4), (0x00, 0x3B, 0x5C)
    c = "#%02X%02X%02X" % tuple(round(a[i] + (b[i] - a[i]) * t) for i in range(3))
    return c, ("#ffffff" if t >= 0.4 else "#003B5C")


def tabela_heat(tid, titulo_col, linhas, meses, rot_mes):
    mx = max([v for _, vals in linhas for v in vals] + [1])
    h = f'<table class="tb" id="{tid}"><thead><tr><th>{titulo_col}</th>' + "".join(f'<th class="cc">{rot_mes(m)}</th>' for m in meses) + '<th class="cc">Total</th></tr></thead><tbody>'
    for nome, vals in linhas:
        h += f"<tr><td>{nome}</td>"
        for v in vals:
            bg, fg = cor(v, mx)
            h += f'<td class="cc" style="background:{bg};color:{fg};font-weight:700;">{v if v else "—"}</td>'
        h += f'<td class="cc"><strong>{sum(vals)}</strong></td></tr>'
    return h + "</tbody></table>"


def pct(a, b): return (100 * a / b) if b else 0
def fp(x): return f"{x:.1f}".replace(".", ",") + "%"


def main():
    if HIST.exists():
        H = json.loads(HIST.read_text(encoding="utf-8"))
    else:
        H = historico_das_paginas()
        HIST.write_text(json.dumps(H, ensure_ascii=False, indent=1), encoding="utf-8")
        print("histórico gravado em", HIST.name, "com", len(H["meses"]), "meses")
    fonte, rows = ler_exportacao()
    ult = dt.date(int(H["meses"][-1].split("/")[1]), MESES.index(H["meses"][-1].split("/")[0]) + 1, 1)
    novos = sorted({(r["dr"].year, r["dr"].month) for r in rows if (r["dr"].year, r["dr"].month) > (ult.year, ult.month)})
    ate = max(r["dr"] for r in rows)
    NM = nomes()
    prof, grup = dict(H["profissional"]), {g: dict(v) for g, v in H["fataisGrupo"].items()}
    grupo_status = {}
    for (a, m) in novos:
        rot = rotulo(a, m)
        sel = [r for r in rows if (r["dr"].year, r["dr"].month) == (a, m)]
        cp = collections.defaultdict(collections.Counter)
        cg = collections.defaultdict(collections.Counter)
        for r in sel:
            s = classifica(r["da"], r["dr"])
            cp[NM.get(r["resp"], r["resp"])][s] += 1
            cg[r["grupo"]][s] += 1
        prof[rot] = sorted([{"Profissional": n, **{k: c[k] for k in ST}, "Total": sum(c.values())} for n, c in cp.items()], key=lambda x: -x["Total"])
        for g, c in cg.items(): grup.setdefault(g, {})[rot] = c["Fatal"]
        grupo_status[rot] = cg
        print(rot, len(sel), "prazos", dict(collections.Counter(classifica(r["da"], r["dr"]) for r in sel)))
    meses = list(prof.keys())
    parcial = rotulo(ate.year, ate.month) if ate.month == novos[-1][1] and ate < dt.date(ate.year, ate.month, 28) else None
    rot_mes = lambda m: m + ("*" if m == parcial else "")
    nota_parcial = f" * {parcial} parcial, com prazos baixados até {ate:%d/%m/%Y}." if parcial else ""
    cheio = [m for m in meses if m != parcial]
    m1, m0 = cheio[-1], cheio[-2]
    tot = {m: collections.Counter() for m in meses}
    for m in meses:
        for r in prof[m]:
            for k in ST: tot[m][k] += r[k]
            tot[m]["Total"] += r["Total"]

    # ---------- Evolução dos Prazos
    ev = (AQUI / "Evolucao_Prazos.html").read_text(encoding="utf-8")
    shutil.copy(AQUI / "Evolucao_Prazos.html", AQUI / "Anteriores" / f"Evolucao_Prazos (backup pre-{dt.date.today():%d%b%Y}).html")
    lg = sorted(((g, [v.get(m, 0) for m in meses]) for g, v in grup.items()), key=lambda x: -sum(x[1]))
    pfat = collections.defaultdict(dict)
    for m in meses:
        for r in prof[m]: pfat[r["Profissional"]][m] = r["Fatal"]
    lp = sorted(((n, [v.get(m, 0) for m in meses]) for n, v in pfat.items() if sum(v.values())), key=lambda x: -sum(x[1]))
    ev = re.sub(r'<table class="tb" id="tblEvGrupo">.*?</table>', lambda _: tabela_heat("tblEvGrupo", "Grupo", lg, meses, rot_mes), ev, flags=re.S)
    ev = re.sub(r'<table class="tb" id="tblEvPessoaMensal">.*?</table>', lambda _: tabela_heat("tblEvPessoaMensal", "Profissional", lp, meses, rot_mes), ev, flags=re.S)
    ev = re.sub(r"(📅 Evolução mensal de prazos Fatais concluídos — )[^<]*", lambda mm: mm.group(1) + " · ".join(rot_mes(m) for m in meses), ev)
    f0, f1 = tot[m0]["Fatal"], tot[m1]["Fatal"]
    var = (f1 - f0) / f0 * 100 if f0 else 0
    gmax = lg[0]
    divididos = {g for g, _ in lg if any(h.startswith(g) and h != g and grup[h].get(m1) and not grup[h].get(m0) for h in grup)}  # ex.: Grupo C virou C1/C2/C3
    melh = sorted(((g, v[meses.index(m0)], v[meses.index(m1)]) for g, v in lg if v[meses.index(m0)] >= 5 and g not in divididos), key=lambda x: (x[2] - x[1]) / x[1])
    txt = (f"📌 No total, os prazos Fatais concluídos passaram de {f0} em {m0} para {f1} em {m1} ({'queda' if var < 0 else 'alta'} de {abs(var):.0f}%). "
           f"{gmax[0]} concentra o maior volume acumulado no período ({sum(gmax[1])} fatais).")
    if melh: txt += f" Maior melhora de {m0} para {m1}: {melh[0][0]}, de {melh[0][1]} para {melh[0][2]} fatais ({(melh[0][2]-melh[0][1])/melh[0][1]*100:+.0f}%)."
    if parcial: txt += f" Em {parcial}, até {ate:%d/%m}, foram {tot[parcial]['Fatal']} fatais em {tot[parcial]['Total']} protocolos ({fp(pct(tot[parcial]['Fatal'], tot[parcial]['Total']))}, contra {fp(pct(f1, tot[m1]['Total']))} em {m1})."
    notas = re.findall(r'<div class="note">📌[^<]*</div>', ev)
    base = H.get("notasEvolucao")  # texto publicado dos meses fechados fica como está; só acrescenta o mês novo
    if base:
        txt = base[0].rstrip() + (f" Em {parcial}, até {ate:%d/%m}, foram {tot[parcial]['Fatal']} fatais em {tot[parcial]['Total']} protocolos ({fp(pct(tot[parcial]['Fatal'], tot[parcial]['Total']))}, contra {fp(pct(f1, tot[m1]['Total']))} em {m1})." if parcial else "")
    ev = ev.replace(notas[0], f'<div class="note">{txt}</div>', 1)
    pm = sorted(((n, v[meses.index(m0)], v[meses.index(m1)]) for n, v in lp if v[meses.index(m0)] >= 5), key=lambda x: (x[2] - x[1]) / x[1])
    txt2 = f"📌 {lp[0][0]} tem o maior volume acumulado de fatais no período ({sum(lp[0][1])})."
    if pm: txt2 += f" Maior melhora individual de {m0} para {m1}: {pm[0][0]}, de {pm[0][1]} para {pm[0][2]} fatais ({(pm[0][2]-pm[0][1])/pm[0][1]*100:+.0f}%)."
    if base: txt2 = base[1].rstrip()
    ev = ev.replace(notas[1], f'<div class="note">{txt2}{nota_parcial}</div>', 1)
    (AQUI / "Evolucao_Prazos.html").write_text(ev, encoding="utf-8")

    # ---------- Protocolados por Mês
    pmh = (AQUI / "Protocolados_Mes.html").read_text(encoding="utf-8")
    shutil.copy(AQUI / "Protocolados_Mes.html", AQUI / "Anteriores" / f"Protocolados_Mes (backup pre-{dt.date.today():%d%b%Y}).html")
    bar = "".join(f'<button class="mes-btn{" active" if i == len(meses) - 1 else ""}" data-mes="{i}" onclick="selecionarMes({i})">{rot_mes(m)}</button>' for i, m in enumerate(meses))
    pmh = re.sub(r'<div class="mes-bar">.*?</div>', f'<div class="mes-bar">{bar}</div>', pmh, count=1, flags=re.S)
    paineis = []
    for i, m in enumerate(meses):
        t = tot[m]
        kp = (f'<div class="kpi-g"><div class="kpi"><div class="kv">{t["Total"]}</div><div class="kl">Total de Protocolos</div></div>'
              f'<div class="kpi vm"><div class="kv">{t["Fatal"]}</div><div class="kl">Fatal</div></div><div class="kpi am"><div class="kv">{t["D-1"]}</div><div class="kl">D-1</div></div>'
              f'<div class="kpi ci"><div class="kv">{t["D-2"]}</div><div class="kl">D-2</div></div><div class="kpi vd"><div class="kv">{t["Futuro"]}</div><div class="kl">Futuro</div></div>'
              f'<div class="kpi lr"><div class="kv">{t["Vencido"]}</div><div class="kl">Vencido</div></div></div>')
        linhas = "".join(f'<tr><td>{r["Profissional"]}</td>' + "".join(f'<td class="cc">{r[k]}</td>' for k in ST) + f'<td class="cc"><strong>{r["Total"]}</strong></td></tr>' for r in prof[m])
        paineis.append(f'<div class="mes-painel" id="mes-painel-{i}" style="display:{"block" if i == len(meses) - 1 else "none"};">\n{kp}\n'
                       f'<input class="srch" id="buscaMes{i}" placeholder="🔎 Buscar profissional..." oninput="filtrarTabela(\'buscaMes{i}\',\'tblMes{i}\')">\n'
                       f'<div class="card">\n<div class="tbl-scroll">\n<table class="tb" id="tblMes{i}"><thead><tr><th>Profissional</th>' + "".join(f'<th class="cc">{k}</th>' for k in ST) +
                       f'<th class="cc">Total</th></tr></thead>\n<tbody>{linhas}</tbody></table>\n</div>\n</div>\n</div>')
    a = pmh.index('<div class="mes-painel" id="mes-painel-0"'); b = pmh.index('<div class="stit"><i></i>Desempenho por Profissional</div>')
    pmh = pmh[:a] + "\n".join(paineis) + "\n\n" + pmh[b:]
    pessoa = collections.defaultdict(list)
    todos = sorted({r["Profissional"] for m in meses for r in prof[m]})
    idx = {m: {r["Profissional"]: r for r in prof[m]} for m in meses}
    for n in todos:
        for m in meses:
            r = idx[m].get(n); pessoa[n].append({k: (r[k] if r else 0) for k in ST + ["Total"]})
    pmh = re.sub(r"const EXPORT_DATA\s*=\s*\{.*?\};\s*\n", lambda _: "const EXPORT_DATA = " + json.dumps(prof, ensure_ascii=False) + ";\n", pmh, count=1, flags=re.S)
    pmh = re.sub(r"const MESES_PERFIL\s*=\s*\[.*?\];", lambda _: "const MESES_PERFIL = " + json.dumps([rot_mes(m) for m in meses], ensure_ascii=False) + ";", pmh, count=1, flags=re.S)
    pmh = re.sub(r"const PESSOA_MES\s*=\s*\{.*?\};\s*\n", lambda _: "const PESSOA_MES = " + json.dumps(dict(pessoa), ensure_ascii=False) + ";\n", pmh, count=1, flags=re.S)
    pmh = re.sub(r'(<div class="ft">Peixoto &amp; Cury Advogados · Controladoria · )[^<]*', lambda mm: mm.group(1) +
                 f"Janeiro a {m1} com o status congelado das planilhas semanais; {parcial or 'meses seguintes'} com o status apurado pelas datas da agenda e da baixa{f' (até {ate:%d/%m/%Y})' if parcial else ''}.", pmh, count=1)
    (AQUI / "Protocolados_Mes.html").write_text(pmh, encoding="utf-8")

    # ---------- Análises
    an = (AQUI / "Analises.html").read_text(encoding="utf-8")
    shutil.copy(AQUI / "Analises.html", AQUI / "Anteriores" / f"Analises (backup pre-{dt.date.today():%d%b%Y}).html")
    an = re.sub(r"(🔍 Análises automáticas — atualizadas a cada rodada de dados \()[^)]*\)", lambda mm: mm.group(1) + " · ".join(rot_mes(m) for m in meses) + ")", an)
    # cartões mensais: reaproveita o cartão do último mês publicado como molde
    cards = list(re.finditer(r'<div class="card" style="text-align:center;">.*?protocolos no mês</div>\s*</div>', an, re.S))
    if cards:
        molde = cards[-1].group(0)
        novos_cards = ""
        for m in meses[len(H["meses"]):]:
            t = tot[m]; c = molde
            vals = [fp(pct(t["Fatal"], t["Total"])), fp(pct(t["D-1"] + t["D-2"] + t["Futuro"], t["Total"])), fp(pct(t["D-1"], t["Total"]))]
            antigos = re.findall(r"\d+,\d%", molde)[:3]
            for o, n in zip(antigos, vals): c = c.replace(o, n, 1)
            c = c.replace(H["meses"][-1], rot_mes(m)).replace(re.search(r"(\d+) protocolos no mês", molde).group(0), f"{t['Total']} protocolos no mês" + (f" (até {ate:%d/%m})" if m == parcial else ""))
            novos_cards += "\n" + c
        an = an[:cards[-1].end()] + novos_cards + an[cards[-1].end():]
    if parcial:
        t = tot[parcial]
        extra = (f'<div class="note">📌 {parcial} (parcial, até {ate:%d/%m/%Y}): {t["Total"]} protocolos, Taxa Fatal de {fp(pct(t["Fatal"], t["Total"]))} '
                 f'e Taxa de Antecipação de {fp(pct(t["D-1"] + t["D-2"] + t["Futuro"], t["Total"]))}, contra {fp(pct(tot[m1]["Fatal"], tot[m1]["Total"]))} e '
                 f'{fp(pct(tot[m1]["D-1"] + tot[m1]["D-2"] + tot[m1]["Futuro"], tot[m1]["Total"]))} em {m1}. O status deste mês foi apurado pelas datas da agenda e da baixa. '
                 f'Os rankings abaixo seguem em {m1}, o mês cheio mais recente.</div>')
        k = an.find('<div class="note">📌')
        an = an[:k] + extra + "\n" + an[k:]
    (AQUI / "Analises.html").write_text(an, encoding="utf-8")
    print("ok:", ", ".join(meses[len(H['meses']):]), "| fonte", fonte, "| até", ate)


if __name__ == "__main__":
    main()
