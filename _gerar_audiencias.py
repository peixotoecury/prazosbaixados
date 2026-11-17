"""Gera Audiencias.html (aba Audiências do prazosbaixados).

Panorama, por mês e por grupo: view do LM v_pc_agendas_audiencia_completo (arquivo exportado na pasta), área trabalhista,
de janeiro do ano até hoje, sem duplicados (processo + data + horário).
Por profissional: a view não traz o advogado da audiência (fica na conta genérica AUDIENCIA), então os 3 cards por
profissional continuam vindo do último Compilado Audiências com a coluna do advogado; o script reaproveita esses cards
da página atual e informa até quando valem.

Uso: python _gerar_audiencias.py "v_pc_agendas_audiencia_completo 2 7.xlsx" [--prof-ate 29/10/2026]
"""
import datetime as dt, html, pathlib, re, shutil, sys, tempfile
import pandas as pd

AQUI = pathlib.Path(__file__).resolve().parent
MESES = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
fmt = lambda n: f"{int(n):,}".replace(",", ".")
pct = lambda a, b: f"{a / b * 100:.1f}%".replace(".", ",") if b else "0,0%"


def categoria(t):
    t = t.lower()
    if "julgamento - vara" in t or "diário oficial" in t or "diario oficial" in t:
        return None   # data de publicação da sentença, não é compromisso com presença
    if "per[íi]cia" in t or "perícia" in t or "pericia" in t or "vistoria" in t:
        return "Perícia"
    if "sess" in t or "julgamento" in t:
        return "Sessão de Julgamento"
    if "audi" in t:
        return "Audiência"
    return "Outros"


def modalidade(t):
    t = t.lower()
    return "Semipresencial" if "semipresencial" in t else "Telepresencial" if "telepresencial" in t else "Presencial" if "presencial" in t else "Não informada"


def main():
    origem = AQUI / sys.argv[1]
    prof_ate = sys.argv[sys.argv.index("--prof-ate") + 1] if "--prof-ate" in sys.argv else ""
    tmp = pathlib.Path(tempfile.gettempdir()) / "aud_lm.xlsx"
    shutil.copy(origem, tmp)   # a planilha pode estar aberta no Excel
    v = pd.read_excel(tmp)
    v["tipo_agenda"] = v["tipo_agenda"].astype(str).str.replace("\xa0", " ").str.replace("\x96", "-").str.strip()
    v["d"] = pd.to_datetime(v["data_audiencia"], dayfirst=True, errors="coerce")
    hoje = pd.Timestamp(dt.date.today())
    ano = hoje.year
    v = v[(v["area"].astype(str).str.upper().str.strip() == "TRABALHISTA") & (v["d"].dt.year == ano)]
    v = v.drop_duplicates(["processo", "data_audiencia", "horario"])
    v["cat"] = v["tipo_agenda"].map(categoria)
    publ = int(v["cat"].isna().sum())
    v = v[v["cat"].notna()]
    fut = v[v["d"] > hoje]
    v = v[v["d"] <= hoje]
    v["mod"] = v["tipo_agenda"].map(modalidade)
    v["grupo"] = v["grupo_processual"].astype(str).str.strip().replace({"Grupo B4": "Grupo B3", "nan": "Sem grupo"})
    tot = len(v)
    n = v["cat"].value_counts()
    m = v["mod"].value_counts()
    mes_fim = MESES[hoje.month - 1]
    periodo = f"Janeiro a {mes_fim}/{ano} (até {hoje:%d/%m})"

    atual = (AQUI / "Audiencias.html").read_text(encoding="utf-8")
    cabeca = atual[: atual.index("<body")]
    prof = re.search(r'<div class="sec" id="sec-profissional-tipo">.*?\n</div>\n</div>\n', atual, re.S)
    prof = prof.group(0) if prof else ""
    if prof and prof_ate:
        prof = prof.replace('<div class="stit"><i></i>Por Profissional — por Tipo de Compromisso</div>',
                            f'<div class="stit"><i></i>Por Profissional — por Tipo de Compromisso (até {prof_ate})</div>'
                            f'\n<div class="note" style="margin-bottom:14px">📌 O cadastro do LM registra as audiências numa conta genérica, sem o advogado que compareceu. '
                            f'Por isso a divisão por profissional segue o último Compilado de Audiências, com dados até {prof_ate}.</div>')

    linhas_mes = []
    for i in range(hoje.month):
        x = v[v["d"].dt.month == i + 1]
        mm = x["mod"].value_counts()
        nome = MESES[i] + (" (parcial)" if i + 1 == hoje.month else "")
        linhas_mes.append(f'<tr><td>{nome}</td><td class="cc"><strong>{fmt(len(x))}</strong></td><td class="cc">{fmt((x["cat"]=="Audiência").sum())}</td>'
                          f'<td class="cc">{fmt((x["cat"]=="Sessão de Julgamento").sum())}</td><td class="cc">{fmt((x["cat"]=="Perícia").sum())}</td>'
                          f'<td class="cc">{fmt(mm.get("Presencial",0))}</td><td class="cc">{fmt(mm.get("Telepresencial",0))}</td><td class="cc">{fmt(mm.get("Semipresencial",0))}</td></tr>')
    g = v.groupby("grupo").agg(t=("cat", "size"), a=("cat", lambda s: (s == "Audiência").sum()), s=("cat", lambda s: (s == "Sessão de Julgamento").sum()),
                               p=("cat", lambda s: (s == "Perícia").sum())).sort_values("t", ascending=False)
    linhas_grupo = [f'<tr><td>{html.escape(k)}</td><td class="cc"><strong>{fmt(r.t)}</strong></td><td class="cc">{pct(r.t, tot)}</td>'
                    f'<td class="cc">{fmt(r.a)}</td><td class="cc">{fmt(r.s)}</td><td class="cc">{fmt(r.p)}</td></tr>' for k, r in g.iterrows()]

    corpo = f'''<body>
<div style="padding:10px 20px;font-size:11.5px;color:#607488;border-bottom:1px solid #e3eaf0;">⚖️ Audiências — {periodo}</div>
<div class="main">

<div class="sec" id="sec-visao">
<div class="stit"><i></i>Panorama do Período</div>
<div class="kpi-g">
<div class="kpi"><div class="kv">{fmt(tot)}</div><div class="kl">Total de compromissos<br>{periodo}</div></div><div class="kpi ci"><div class="kv">{fmt(n.get("Audiência",0))}</div><div class="kl">Audiências<br>inicial/instrução/una/conciliação</div></div><div class="kpi"><div class="kv">{fmt(n.get("Sessão de Julgamento",0))}</div><div class="kl">Sessões de Julgamento<br>TRT/TST/TJ</div></div><div class="kpi lr"><div class="kv">{fmt(n.get("Perícia",0))}</div><div class="kl">Perícias e vistorias</div></div><div class="kpi vd"><div class="kv">{fmt(m.get("Telepresencial",0))}</div><div class="kl">Telepresencial<br>{pct(m.get("Telepresencial",0), tot)} do total</div></div><div class="kpi am"><div class="kv">{fmt(m.get("Presencial",0))}</div><div class="kl">Presencial<br>{pct(m.get("Presencial",0), tot)} do total</div></div><div class="kpi"><div class="kv">{fmt(len(fut))}</div><div class="kl">Já agendados<br>depois de {hoje:%d/%m/%Y}</div></div>
</div>
<div class="note">
📋 <strong>Fonte: cadastro de audiências do LM</strong>, área trabalhista: audiências, sessões de julgamento e perícias de janeiro até {hoje:%d/%m/%Y}, sem registros repetidos do mesmo processo, data e horário. As datas de julgamento publicadas no Diário Oficial ({fmt(publ)} no ano) não entram, por não serem compromissos com comparecimento. O total é maior que o das versões anteriores desta aba, que usavam a planilha de controle da equipe e não incluíam todas as perícias e sessões.
</div>
</div>

<div class="sec" id="sec-mes">
<div class="stit"><i></i>Distribuição por Mês</div>
  <div class="card">
    <div class="ct">📊 Compromissos — por Mês</div>
    <div class="tbl-scroll">
    <table class="tb"><thead><tr><th>Mês</th><th class="cc">Total</th><th class="cc">Audiências</th><th class="cc">Sessões</th><th class="cc">Perícias</th><th class="cc">Presencial</th><th class="cc">Telepresencial</th><th class="cc">Semipresencial</th></tr></thead>
    <tbody>{''.join(linhas_mes)}</tbody></table>
    </div>
  </div>
</div>

<div class="sec" id="sec-grupo">
<div class="stit"><i></i>Distribuição por Grupo</div>
  <div class="card">
    <div class="ct">👥 Compromissos — por Grupo</div>
    <div class="tbl-scroll">
    <table class="tb"><thead><tr><th>Grupo</th><th class="cc">Total</th><th class="cc">% do total</th><th class="cc">Audiências</th><th class="cc">Sessões</th><th class="cc">Perícias</th></tr></thead>
    <tbody>{''.join(linhas_grupo)}</tbody></table>
    </div>
  </div>
<div class="note">
📌 <strong>Metodologia:</strong> o grupo é o grupo processual do processo no LM. Registros do antigo Grupo B4 foram somados ao Grupo B3.
</div>
</div>

{prof}
</div>
<div class="ft">Peixoto &amp; Cury Advogados · Dashboard gerado a partir do cadastro de audiências do LM · {periodo}</div>
</body>
</html>
'''
    (AQUI / "Audiencias.html").write_text(cabeca + corpo, encoding="utf-8")
    print(f"Audiencias.html: {tot} compromissos ({dict(n)}); {len(fut)} futuros; {publ} julgamentos no DO fora; período {periodo}")


if __name__ == "__main__":
    main()
