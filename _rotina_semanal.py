# -*- coding: utf-8 -*-
"""Rotina semanal do prazosbaixados (tarefa agendada PC_BaixadosSemanal, segunda-feira às 7h).

1. Lê prazos e agendas direto do LM (BAIXADOS_FONTE=LM).
2. Protocolados: gera a semana que fechou no domingo (_gerar_semanal.py). Se a última semana publicada era parcial
   (ex.: S47 publicada até sexta 20/11), ela é refeita inteira a partir do backup anterior a ela.
3. Protocolados por Mês, Evolução e Análises: atualiza o mês corrente (_gerar_mensal.py). Meses fechados não mudam.
4. Publica (commit + push) e sincroniza o outro clone (temp_mobile_fix).
5. Opcional: e-mail com o resumo dos Fatais fora do horário (destinatários em _rotina_config.json; lista vazia = não envia).
Log em _rotina_semanal.log.

Teste sem publicar:  python _rotina_semanal.py --sem-publicar
"""
import os, sys, re, json, shutil, subprocess, datetime as dt, traceback
from pathlib import Path

AQUI = Path(__file__).resolve().parent
OUTRO_CLONE = Path(r"C:\Users\ach\temp_mobile_fix\prazosbaixados")
LOG = AQUI / "_rotina_semanal.log"
CONFIG = AQUI / "_rotina_config.json"
PY = sys.executable.replace("pythonw.exe", "python.exe")
os.environ["BAIXADOS_FONTE"] = "LM"


def log(msg):
    linha = f"[{dt.datetime.now():%d/%m/%Y %H:%M}] {msg}"
    print(linha)
    with open(LOG, "a", encoding="utf-8") as f: f.write(linha + "\n")


def roda(*args, cwd=AQUI):
    r = subprocess.run(list(args), cwd=cwd, capture_output=True, text=True, encoding="utf-8", errors="replace", env=os.environ)
    saida = (r.stdout + r.stderr).strip()
    if saida: log("  " + saida.replace("\n", "\n  ")[-1500:])
    if r.returncode: raise RuntimeError(f"falhou: {' '.join(map(str, args))}")
    return saida


def ultima_semana():
    t = (AQUI / "protocolados.html").read_text(encoding="utf-8")
    labels = re.findall(r"'(\d{4})-S(\d+)':'S\d+<br><small[^>]*>([\d/]+)-([\d/]+)</small>'", t)
    ano, num, ini, fim = max(labels, key=lambda x: (int(x[0]), int(x[1])))
    f = dt.date(int(ano), int(fim.split("/")[1]), int(fim.split("/")[0]))
    return int(num), f


def main():
    publicar = "--sem-publicar" not in sys.argv
    hoje = next((dt.date.fromisoformat(a.split("=")[1]) for a in sys.argv if a.startswith("--hoje=")), dt.date.today())  # --hoje= só para teste
    fim = hoje - dt.timedelta(days=hoje.weekday() + 1)          # domingo passado
    ini = fim - dt.timedelta(days=6)                              # segunda da semana fechada
    log(f"rotina semanal: semana {ini:%d/%m} a {fim:%d/%m}")
    if publicar:
        roda("git", "pull", "--ff-only", "-q")
    num, fim_pub = ultima_semana()
    if fim_pub >= fim:
        log(f"S{num} já publicada até {fim_pub:%d/%m}; nada a fazer no semanal")
    else:
        if fim_pub >= ini:  # última publicada é a mesma semana, parcial: volta ao estado anterior a ela e refaz inteira
            bk = AQUI / "Anteriores" / f"protocolados (backup pre-semana{num}).html"
            if not bk.exists(): raise RuntimeError(f"semana S{num} parcial e sem backup {bk.name}")
            shutil.copy(bk, AQUI / "protocolados.html")
            log(f"S{num} estava parcial (até {fim_pub:%d/%m}); refeita a partir de {bk.name}")
        roda(PY, "-X", "utf8", "_gerar_semanal.py", f"{ini:%Y-%m-%d}", f"{fim:%Y-%m-%d}")
    roda(PY, "-X", "utf8", "_gerar_mensal.py")

    if publicar:
        roda("git", "add", "protocolados.html", "Protocolados_Mes.html", "Evolucao_Prazos.html", "Analises.html")
        st = subprocess.run(["git", "diff", "--cached", "--quiet"], cwd=AQUI).returncode
        if st:
            roda("git", "commit", "-q", "-m", f"Rotina semanal: semana {ini:%d/%m} a {fim:%d/%m}/{fim:%Y} e mês corrente (LM)")
            roda("git", "push", "-q")
            if OUTRO_CLONE.exists():
                roda("git", "fetch", "-q", cwd=OUTRO_CLONE); roda("git", "merge", "--ff-only", "-q", "origin/main", cwd=OUTRO_CLONE)
            log("publicado")
        else:
            log("sem mudanças para publicar")
    email(ini, fim)


def email(ini, fim):
    cfg = json.loads(CONFIG.read_text(encoding="utf-8")) if CONFIG.exists() else {}
    para = cfg.get("email_para") or []
    if not para:
        log("e-mail: sem destinatários em _rotina_config.json, não enviado"); return
    assunto, corpo = montar_email(ini, fim)
    import win32com.client
    ol = win32com.client.Dispatch("Outlook.Application")
    mail = ol.CreateItem(0)
    if "eu" in para:  # "eu" = a própria conta do Outlook deste computador
        u = ol.Session.CurrentUser.AddressEntry
        eu = (u.GetExchangeUser().PrimarySmtpAddress if u.GetExchangeUser() else u.Address)
        para = [eu if x == "eu" else x for x in para]
    mail.To = "; ".join(para)
    mail.Subject = assunto
    logo = AQUI.parent / "BAND" / "Bases" / "logo_pc_branca.png"
    if logo.exists():  # logo embutido no corpo (cid), o Outlook não mostra imagem em data URI
        anexo = mail.Attachments.Add(str(logo))
        anexo.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "logo_pc")
    mail.HTMLBody = corpo
    mail.Send()
    log(f"e-mail enviado para {', '.join(para)}")


# ---------------------------------------------------------------- e-mail visual (tabelas e estilos inline, para o Outlook)
NAVY, CIANO, CINZA, LINHA = "#003B5C", "#00A9E0", "#5F6B7A", "#E3E8EE"


def _json_const(t, nome):
    m = re.search(r"const " + nome + r"\s*=\s*", t)
    i = m.end(); dep = 0; instr = False
    for j in range(i, len(t)):
        ch = t[j]
        if ch == '"' and t[j - 1] != chr(92): instr = not instr
        if instr: continue
        if ch in "{[": dep += 1
        elif ch in "}]":
            dep -= 1
            if dep == 0: return json.loads(t[i:j + 1])


def _nomes(t):
    """sigla -> nome: mapas da própria página (NOMES e MAPA_SIGLAS) + cadastro de profissionais."""
    m = {}
    try:
        import openpyxl
        it = openpyxl.load_workbook(AQUI / "Profissionais trabalhistas.xlsx", read_only=True)["Planilha2"].iter_rows(values_only=True); next(it)
        for r in it:
            if r[0] and r[1]: m[str(r[0]).strip().upper()] = str(r[1]).replace(" ", " ").strip()
    except Exception: pass
    for nome in ("NOMES", "MAPA_SIGLAS"):
        k = t.find("const " + nome)
        if k >= 0:
            bloco = t[k:t.index("}", k)]
            m.update({a: b for a, b in re.findall(r"'([A-Z]{2,4})'\s*:\s*'([^']+)'", bloco)})
    m.update({"AEM": "Aline Menendes Aponte Moraes", "TAR": "Thais Amaral Rocca", "MIM": "Maria Imperatriz", "TLI": "Tamires Lins Costa",
              "BTF": "Bruno Tanganelli Farah", "LWP": "Lucas Willott Pereira"})  # quem entrou depois do cadastro
    m["PRAZO"] = "Conta genérica PRAZO"
    return m


def _curto(n):
    p = n.split()
    return " ".join(p[:2]) if len(p) <= 2 or p[1].lower() not in ("de", "da", "do", "dos", "das") else " ".join(p[:3])


def _tabela_grupo_sigla(contagem, cor, NOM=None, extra=None, tot_grupo=None):
    """Uma tabela só: Grupo (mesclado, com o total do grupo) | Sigla | barra | Qtde."""
    import collections
    pg = collections.Counter()
    for (g, _), n in contagem.items(): pg[g] += n
    if not pg:
        return f'<p style="font-family:Segoe UI,Arial;font-size:13px;color:{CINZA};margin:4px 0 18px">Nenhum caso na semana.</p>'
    mx = max(contagem.values())
    th = f'style="background:{NAVY};color:#ffffff;font-size:11px;text-transform:uppercase;letter-spacing:.5px;padding:7px 10px;text-align:left"'
    h = (f'<table cellpadding="0" cellspacing="0" style="border-collapse:collapse;font-family:Segoe UI,Arial;font-size:13px;width:620px;margin:4px 0 22px">'
         f'<tr><th {th}>Grupo</th><th {th}>Profissional</th><th {th} colspan="2">Qtde</th></tr>')
    for g, tot in pg.most_common():
        itens = sorted(((sg, n) for (gg, sg), n in contagem.items() if gg == g), key=lambda x: (-x[1], x[0]))
        for k, (sg, n) in enumerate(itens):
            borda = f"border-top:2px solid {NAVY};" if k == 0 else f"border-top:1px solid {LINHA};"
            cel_g = (f'<td rowspan="{len(itens)}" style="{borda}padding:7px 10px;vertical-align:top;background:#F4F7FA;width:120px">'
                     f'<b style="color:{NAVY}">{g}</b><br><span style="font-size:11px;color:{CINZA}">{(tot_grupo or {}).get(g, tot)} no grupo</span></td>') if k == 0 else ""
            barra = max(4, round(150 * n / mx))
            nome = _curto((NOM or {}).get(sg, sg))
            comp = f" · {extra[(g, sg)]}" if extra and (g, sg) in extra else ""
            rot = f'{nome}<br><span style="font-size:10px;color:{CINZA};font-weight:400">{sg}{comp}</span>'
            h += (f'<tr>{cel_g}<td style="{borda}padding:6px 10px;font-weight:600;width:170px">{rot}</td>'
                  f'<td style="{borda}padding:6px 4px 6px 10px;width:160px"><table cellpadding="0" cellspacing="0"><tr>'
                  f'<td style="background:{cor};height:10px;width:{barra}px;font-size:1px;line-height:1px">&nbsp;</td></tr></table></td>'
                  f'<td style="{borda}padding:6px 10px;text-align:right;font-weight:700;color:{cor};width:40px">{n}</td></tr>')
    return h + "</table>"


def _secao_abertos(NOM):
    """Prazos em aberto no LM cujo dia fatal (dia útil seguinte à agenda) já passou, na hora do envio."""
    import collections, numpy as np
    try:
        from _fonte_lm import pendentes
        hoje = dt.date.today()
        fer = ["2026-11-02", "2026-11-20", "2026-12-25", "2027-01-01"]
        lin = [r for r in pendentes(hoje) if str(r[2] or "").strip().upper() not in ("PRAZO", "")
               and str(r[5] or "").strip().upper() != "PRAZO" and np.busday_count(r[3], hoje, holidays=fer) > 1]
    except Exception as e:
        log(f"e-mail: prazos em aberto não consultados ({e})")
        return ""
    cont = collections.Counter((str(r[1] or "Sem grupo").strip(), str(r[2]).strip().upper()) for r in lin)
    return (_secao_titulo(f"Em aberto e já vencidos · {len(lin)}", f"prazos ainda não baixados no LM cujo dia fatal já passou (posição de {dt.datetime.now():%d/%m às %Hh%M})")
            + _tabela_grupo_sigla(cont, "#8E1B10", NOM))


def _secao_titulo(titulo, sub):
    return (f'<div style="font-family:Segoe UI,Arial;margin-top:6px"><div style="font-size:15px;font-weight:700;color:{NAVY};border-left:4px solid {CIANO};padding-left:8px">{titulo}</div>'
            f'<div style="font-size:11px;color:{CINZA};padding-left:12px;margin-bottom:6px">{sub}</div></div>')


def montar_email(ini, fim):
    import collections
    t = (AQUI / "protocolados.html").read_text(encoding="utf-8")
    resp = _json_const(t, "RESP")
    ff = _json_const(t, "FAT_FORA") or []
    kp = {k: v for k, v in re.findall(r'<div class="klbl">([^<]+)</div><div class="kval[^"]*"[^>]*>([^<]+)</div>', t) if "${" not in k + v}
    ven = collections.Counter({(r["grp"], r["n"]): r["ven"] for r in resp if r.get("ven")})
    ffc = collections.Counter((x["grp"], x["resp"]) for x in ff)
    n_ven, n_ff = sum(ven.values()), len(ff)
    NOM = _nomes(t)
    # quem mais protocolou no fatal: as 20 pessoas com mais fatais na semana, agrupadas pelo grupo
    fat_all = {(r["grp"], r["n"]): r["fat"] for r in resp if r.get("fat") and r["n"] != "PRAZO"}
    top = dict(sorted(fat_all.items(), key=lambda kv: -kv[1]))  # todos os que protocolaram no fatal
    fat_top = collections.Counter(top)
    tot_p = {(r["grp"], r["n"]): r["tot"] for r in resp}
    pct = {k: f"{100 * v / tot_p[k]:.0f}% dos prazos da semana".replace(".", ",") for k, v in top.items() if tot_p.get(k)}

    def tile(rot, val, cor, sub=""):
        return (f'<td style="padding:4px"><table cellpadding="0" cellspacing="0" style="width:160px;border-collapse:collapse"><tr>'
                f'<td style="border-top:4px solid {cor};background:#F4F7FA;padding:10px 12px;font-family:Segoe UI,Arial">'
                f'<div style="font-size:10px;color:{CINZA};text-transform:uppercase;letter-spacing:.6px">{rot}</div>'
                f'<div style="font-size:24px;font-weight:700;color:{cor};line-height:1.2">{val}</div>'
                f'<div style="font-size:10px;color:{CINZA}">{sub}</div></td></tr></table></td>')
    cards = ('<table cellpadding="0" cellspacing="0" style="margin:10px 0 18px"><tr>'
             + tile("Vencidos", kp.get("Vencidos", n_ven), "#C0392B", "baixados depois do fatal")
             + tile("Fatal fora do horário", n_ff, "#E74C3C", "antes 9h ou após 19h")
             + tile("Fatais", kp.get("Fatais", "-"), "#E67E22", "baixados no dia fatal")
             + "</tr><tr>"
             + tile("Total baixado", kp.get("Total Baixado", "-"), NAVY, "prazos realizados")
             + tile("D-1 (no dia)", kp.get("D-1 (no dia)", "-"), "#B8860B", "cumpridos no prazo")
             + tile("Antecipados", kp.get("Antecipados", "-"), "#1E7A3E", "antes do prazo")
             + "</tr></table>")

    def secao(titulo, sub):
        return (f'<div style="font-family:Segoe UI,Arial;margin-top:6px"><div style="font-size:15px;font-weight:700;color:{NAVY};border-left:4px solid {CIANO};padding-left:8px">{titulo}</div>'
                f'<div style="font-size:11px;color:{CINZA};padding-left:12px;margin-bottom:6px">{sub}</div></div>')

    td = f'style="padding:4px 8px;border-top:1px solid {LINHA}"'
    casos = "".join(f'<tr><td {td}>{x["dt"]}</td><td {td}>{x["hora_real"][:2]}h</td><td {td}><b>{_curto(NOM.get(x["resp"], x["resp"]))}</b></td><td {td}>{x["grp"]}</td><td {td}>{x["proc"]}</td></tr>' for x in ff)
    th2 = f'style="text-align:left;padding:5px 8px;font-size:10px;color:{CINZA};text-transform:uppercase;border-bottom:2px solid {NAVY}"'
    corpo = (f'<div style="font-family:Segoe UI,Arial;color:#1A2130;max-width:700px">'
             f'<table cellpadding="0" cellspacing="0" style="border-collapse:collapse;width:700px"><tr>'
             f'<td style="background:{NAVY};padding:16px 20px;font-family:Segoe UI,Arial;vertical-align:middle">'
             f'<div style="font-size:11px;color:#ffffff;letter-spacing:1.5px;text-transform:uppercase;opacity:.85">Controladoria · Área Trabalhista</div>'
             f'<div style="font-size:20px;font-weight:700;color:#ffffff;margin-top:2px">Prazos baixados de {ini:%d/%m} a {fim:%d/%m/%Y}</div></td>'
             f'<td style="background:{NAVY};padding:12px 20px;text-align:right;vertical-align:middle;width:150px">'
             f'<img src="cid:logo_pc" alt="Peixoto &amp; Cury Advogados" height="44" style="height:44px;display:block;margin-left:auto"></td>'
             f'</tr><tr><td colspan="2" style="background:{CIANO};height:3px;font-size:1px;line-height:1px">&nbsp;</td></tr></table>'
             + cards
             + secao(f"Vencidos · {n_ven}", "prazos baixados depois do dia fatal, por grupo e profissional")
             + _tabela_grupo_sigla(ven, "#C0392B", NOM)
             + _secao_abertos(NOM)
             + secao(f"Fatais fora do horário · {n_ff}", "prazos baixados no dia fatal antes das 9h ou a partir das 19h, por grupo e profissional")
             + _tabela_grupo_sigla(ffc, "#E74C3C", NOM)
             + secao(f"Quem protocolou no fatal · {sum(fat_top.values())}", "todos os prazos baixados no dia fatal na semana, por grupo e profissional (do maior para o menor)")
             + _tabela_grupo_sigla(fat_top, "#E67E22", NOM, pct, {g: v["fat"] for g, v in _json_const(t, "BY_GRUPO").items()})
             + secao("Casos de fatal fora do horário", "a hora aparece só cheia, porque o LM grava os minutos com erro")
             + f'<table cellpadding="0" cellspacing="0" style="border-collapse:collapse;font-size:11.5px;width:620px;margin:4px 0 18px"><tr><th {th2}>Dia</th><th {th2}>Hora</th><th {th2}>Profissional</th><th {th2}>Grupo</th><th {th2}>Processo</th></tr>{casos}</table>'
             + f'<div style="font-size:12px;margin:10px 0">Painel completo: <a href="https://peixotoecury.github.io/prazosbaixados/" style="color:{CIANO}">peixotoecury.github.io/prazosbaixados</a></div>'
             + f'<div style="font-size:10px;color:{CINZA};border-top:1px solid {LINHA};padding-top:6px">Enviado automaticamente pela rotina semanal da Controladoria. Peixoto &amp; Cury Advogados.</div></div>')
    assunto = f"Prazos baixados {ini:%d/%m} a {fim:%d/%m}: {n_ven} vencidos e {n_ff} fatais fora do horário"
    (AQUI / "_previa_email.html").write_text(corpo, encoding="utf-8")
    return assunto, corpo


if __name__ == "__main__":
    try:
        main()
    except Exception:
        log("ERRO\n" + traceback.format_exc())
        sys.exit(1)
