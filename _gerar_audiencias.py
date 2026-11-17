"""Gera Audiencias.html (aba Audiências do prazosbaixados) a partir de audiencias_fonte.carregar():
Compilado Audiências até a última data dele e, depois, a Pauta de Audiências (que guarda o advogado de cada audiência).

Grupo = grupo do profissional no cadastro (dados da Visão Geral do Escritório); correspondentes em bucket próprio;
"Grupo B4" soma no B3. Por profissional: 3 cards (Audiências, Sessões, Perícias), todo mundo listado, sem "Outros".
Uso: python _gerar_audiencias.py
"""
import datetime as dt, html, json, pathlib, re
import pandas as pd
from audiencias_fonte import carregar

AQUI = pathlib.Path(__file__).resolve().parent
EQUIPE = AQUI.parent / "Visão Geral dos processos do escritório" / "dados.json"
MESES = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
SIGLA_FIX = {"JTR": "JRT"}   # código digitado com letras trocadas na planilha
fmt = lambda n: f"{int(n):,}".replace(",", ".")
pct = lambda a, b: f"{a / b * 100:.1f}%".replace(".", ",") if b else "0,0%"


def externo(a):
    """Correspondente: escrito como tal ou qualquer nome que não seja uma sigla de profissional do escritório."""
    return a.lower().startswith(("correspondente", "contratação de correspondente")) or not re.fullmatch(r"[A-Za-z]{2,4}", a.strip())


def correspondente(a):
    a = re.sub(r"\s+", " ", a).strip()
    a = re.sub(r"(?i)^contrata[çc][ãa]o de correspondente\s*", "Correspondente - ", a)
    if not a.lower().startswith("correspondente"):
        a = "Correspondente - " + a.title()
    if re.match(r"(?i)correspondente\s*-\s*ana\b", a):
        return "Correspondente - Ana Carolina"
    if re.match(r"(?i)correspondente\s*-\s*camila\b", a):
        return "Correspondente - Camila Cruz"
    return a


def main():
    df, info = carregar()
    hoje = pd.Timestamp(dt.date.today())
    eq = json.loads(EQUIPE.read_text(encoding="utf-8"))["equipe"] if EQUIPE.exists() else []
    nome = {e["sigla"]: e["nome"] for e in eq}
    grupo_de = {e["sigla"]: (e["grupos"] or [""])[0] for e in eq}
    grupo_de["AAP"] = "E1"   # confirmado pela usuária

    fut = df[df["dt"] > hoje]
    v = df[df["dt"] <= hoje].copy()
    v["sig"] = v["adv"].str.upper().map(lambda s: SIGLA_FIX.get(s, s))
    def quem(r):
        if r.pendente:
            return "Sem advogado designado"
        if externo(r.adv):
            return correspondente(r.adv)
        return nome.get(r.sig, r.adv)
    def grupo(r):
        if externo(r.adv) and not r.pendente:
            return "Correspondentes"
        g = grupo_de.get(r.sig, "")
        if not re.fullmatch(r"[A-E]\d?", g or ""):
            m = re.search(r"Grupo\s+([A-E]\d?)", r.equipe or "") or re.search(r"Grupo\s+([A-E]\d?)", r.grupo_proc or "")
            g = m.group(1) if m else ""
        g = "B3" if g == "B4" else g
        return f"Grupo {g}" if g else "Sem Grupo"
    v["quem"] = v.apply(quem, axis=1)
    v["grupo"] = v.apply(grupo, axis=1)
    tot = len(v); n = v["cat"].value_counts(); m = v["mod"].value_counts()
    periodo = f"Janeiro a {MESES[hoje.month - 1]}/{hoje.year} (até {hoje:%d/%m})"

    linhas_mes = []
    for i in range(hoje.month):
        x = v[v["dt"].dt.month == i + 1]; mm = x["mod"].value_counts()
        nm = MESES[i] + (" (parcial)" if i + 1 == hoje.month else "")
        linhas_mes.append(f'<tr><td>{nm}</td><td class="cc"><strong>{fmt(len(x))}</strong></td><td class="cc">{fmt(mm.get("Presencial",0))}</td>'
                          f'<td class="cc">{fmt(mm.get("Telepresencial",0))}</td><td class="cc">{fmt(mm.get("Semipresencial",0))}</td></tr>')
    g = v["grupo"].value_counts()
    linhas_grupo = [f'<tr><td>{html.escape(k)}</td><td class="cc"><strong>{fmt(c)}</strong></td><td class="cc">{pct(c, tot)}</td></tr>' for k, c in g.items()]

    def card(cat, icone, titulo):
        x = v[v["cat"] == cat]["quem"].value_counts()
        rows = "".join(f'<tr><td>{html.escape(k)}</td><td class="cc"><strong>{fmt(c)}</strong></td></tr>' for k, c in x.items())
        return (f'  <div class="card">\n    <div class="ct">{icone} {titulo} — por Profissional <span class="tag">{fmt(x.sum())} no total</span></div>\n'
                f'    <div class="tbl-scroll">\n    <table class="tb"><thead><tr><th>Profissional</th><th class="cc">Total</th></tr></thead>\n'
                f'    <tbody>{rows}</tbody></table>\n    </div>\n  </div>\n')

    futuros_sem = fut[fut["pendente"]].sort_values("dt")
    alerta = ""
    if len(futuros_sem):
        lin = "".join(f'<tr><td>{r.dt:%d/%m/%Y}</td><td>{html.escape(str(r.horario)[:5])}</td><td>{html.escape(r.tipo)}</td><td>{html.escape(r.processo)}</td>'
                      f'<td>{html.escape(r.grupo_proc or "")}</td></tr>' for r in futuros_sem.itertuples())
        alerta = f'''
<div class="sec" id="sec-pendentes">
<div class="stit"><i></i>Próximas audiências sem advogado designado</div>
  <div class="card">
    <div class="ct">⚠️ Agendadas a partir de {hoje + pd.Timedelta(days=1):%d/%m/%Y} ainda como "Indicar Advogado" na Pauta <span class="tag">{fmt(len(futuros_sem))} no total</span></div>
    <div class="tbl-scroll">
    <table class="tb"><thead><tr><th>Data</th><th>Horário</th><th>Tipo</th><th>Processo</th><th>Grupo do processo</th></tr></thead>
    <tbody>{lin}</tbody></table>
    </div>
  </div>
</div>
'''

    atual = (AQUI / "Audiencias.html").read_text(encoding="utf-8")
    cabeca = atual[: atual.index("<body")]
    fonte_txt = (f'Compilado Audiências até {info["compilado_ate"]} e Pauta de Audiências de {info["pauta_de"]} até {info["ate"]}'
                 if info["compilado_ate"] else f'Pauta de Audiências até {info["ate"]}')
    corpo = f'''<body>
<div style="padding:10px 20px;font-size:11.5px;color:#607488;border-bottom:1px solid #e3eaf0;">⚖️ Audiências — {periodo}</div>
<div class="main">

<div class="sec" id="sec-visao">
<div class="stit"><i></i>Panorama do Período</div>
<div class="kpi-g">
<div class="kpi"><div class="kv">{fmt(tot)}</div><div class="kl">Total de Audiências<br>{periodo}</div></div><div class="kpi ci"><div class="kv">{fmt(n.get("Audiência",0))}</div><div class="kl">Audiências<br>inicial/instrução/una/conciliação</div></div><div class="kpi"><div class="kv">{fmt(n.get("Sessão de Julgamento",0))}</div><div class="kl">Sessões de Julgamento<br>TRT/TST/TJ</div></div><div class="kpi lr"><div class="kv">{fmt(n.get("Perícia",0))}</div><div class="kl">Perícias</div></div><div class="kpi vd"><div class="kv">{fmt(m.get("Telepresencial",0))}</div><div class="kl">Telepresencial<br>{pct(m.get("Telepresencial",0), tot)} do total</div></div><div class="kpi am"><div class="kv">{fmt(m.get("Presencial",0))}</div><div class="kl">Presencial<br>{pct(m.get("Presencial",0), tot)} do total</div></div><div class="kpi vm"><div class="kv">{fmt(len(futuros_sem))}</div><div class="kl">Próximas sem advogado<br>designado na Pauta</div></div>
</div>
<div class="note">
📋 <strong>Fonte: {fonte_txt}.</strong> Audiências, sessões de julgamento e perícias acompanhadas pelo escritório, sem registros repetidos do mesmo processo, data e horário. A Pauta de Audiências é atualizada diariamente pela equipe e registra o advogado de cada audiência, o que mantém esta aba em dia sem esperar o fechamento da planilha mensal.
</div>
</div>

<div class="sec" id="sec-mes">
<div class="stit"><i></i>Distribuição por Mês</div>
  <div class="card">
    <div class="ct">📊 Audiências — por Mês</div>
    <div class="tbl-scroll">
    <table class="tb"><thead><tr><th>Mês</th><th class="cc">Total</th><th class="cc">Presencial</th><th class="cc">Telepresencial</th><th class="cc">Semipresencial</th></tr></thead>
    <tbody>{''.join(linhas_mes)}</tbody></table>
    </div>
  </div>
</div>

<div class="sec" id="sec-grupo">
<div class="stit"><i></i>Distribuição por Grupo</div>
  <div class="card">
    <div class="ct">👥 Audiências — por Grupo</div>
    <div class="tbl-scroll">
    <table class="tb"><thead><tr><th>Grupo</th><th class="cc">Total</th><th class="cc">% do total</th></tr></thead>
    <tbody>{''.join(linhas_grupo)}</tbody></table>
    </div>
  </div>
<div class="note">
📌 <strong>Metodologia:</strong> o Grupo de cada audiência é o grupo do profissional que a acompanha, conforme o cadastro de profissionais. Correspondentes externos formam um grupo próprio. Quando o profissional não é encontrado no cadastro, vale o grupo indicado na planilha ou na Pauta.
</div>
</div>

<div class="sec" id="sec-profissional-tipo">
<div class="stit"><i></i>Por Profissional — por Tipo de Compromisso</div>
<div class="g3">
{card("Audiência", "🗣️", "Audiências")}{card("Sessão de Julgamento", "⚖️", "Sessões de Julgamento")}{card("Perícia", "🔬", "Perícias")}</div>
</div>
{alerta}
</div>
<div class="ft">Peixoto &amp; Cury Advogados · Dashboard gerado a partir do Compilado Audiências e da Pauta de Audiências · {periodo}</div>
</body>
</html>
'''
    (AQUI / "Audiencias.html").write_text(cabeca + corpo, encoding="utf-8")
    print(f"Audiencias.html: {tot} ({dict(n)}); fonte: {fonte_txt}; próximas sem advogado: {len(futuros_sem)}")
    print("grupos:", g.to_dict())
    print("sem nome no cadastro:", sorted(set(v[(v['quem']==v['adv'])&~v['adv'].str.lower().str.startswith('correspondente')]['adv']))[:20])


if __name__ == "__main__":
    main()
