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
    t = (AQUI / "protocolados.html").read_text(encoding="utf-8")
    m = re.search(r"const FAT_FORA\s*=\s*(\[.*?\]);", t, re.S)
    ff = json.loads(m.group(1)) if m else []
    kp = dict(re.findall(r'<div class="klbl">([^<]+)</div><div class="kval[^"]*"[^>]*>([^<]+)</div>', t))
    linhas = "".join(f"<tr><td>{x['dt']}</td><td>{x['hora_real'][:5]}</td><td>{x['resp']}</td><td>{x['grp']}</td><td>{x['proc']}</td></tr>" for x in ff)
    corpo = (f"<p>Prazos baixados de {ini:%d/%m} a {fim:%d/%m/%Y}:</p><ul>"
             + "".join(f"<li>{k}: <b>{v}</b></li>" for k, v in kp.items())
             + f"</ul><p><b>Fatais fora do horário (antes das 9h ou a partir das 19h): {len(ff)}</b></p>"
             + f"<table border='1' cellpadding='4' style='border-collapse:collapse;font-size:12px'><tr><th>Dia</th><th>Hora</th><th>Responsável</th><th>Grupo</th><th>Processo</th></tr>{linhas}</table>"
             + "<p>Painel completo: https://peixotoecury.github.io/prazosbaixados/</p>")
    import win32com.client
    ol = win32com.client.Dispatch("Outlook.Application")
    mail = ol.CreateItem(0)
    if "eu" in para:  # "eu" = a própria conta do Outlook deste computador
        u = ol.Session.CurrentUser.AddressEntry
        eu = (u.GetExchangeUser().PrimarySmtpAddress if u.GetExchangeUser() else u.Address)
        para = [eu if x == "eu" else x for x in para]
    mail.To = "; ".join(para)
    mail.Subject = f"Prazos baixados {ini:%d/%m} a {fim:%d/%m}: {len(ff)} fatais fora do horário"
    mail.HTMLBody = corpo
    mail.Send()
    log(f"e-mail enviado para {', '.join(para)}")


if __name__ == "__main__":
    try:
        main()
    except Exception:
        log("ERRO\n" + traceback.format_exc())
        sys.exit(1)
