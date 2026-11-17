"""Audiências, sessões de julgamento e perícias do ano, com o advogado que acompanha.

Fonte 1, Compilado Audiências (planilha da equipe, coluna "Advogado Responsável Audiência/SO"): até a última data que ele tiver.
Fonte 2, Pauta de Audiências (peixotoecury.github.io/pauta-audiencias, base Supabase "global": histórico + pauta atual
+ designações feitas no site): do dia seguinte ao fim do Compilado até hoje (e o que já está agendado depois).
O LM não serve para isso: no LM toda audiência fica na conta genérica AUDIENCIA, sem o advogado.

Universo igual ao do Compilado: audiências que o escritório acompanha. Na Pauta, "N/A" (não iremos acompanhar) fica de
fora; "Indicar Advogado" entra no total e é sinalizado como pendente de designação.

Uso: from audiencias_fonte import carregar; df, info = carregar()
"""
import datetime as dt, json, pathlib, re, shutil, tempfile, urllib.request
import pandas as pd

AQUI = pathlib.Path(__file__).resolve().parent
PAUTA_URL = "https://zorxoydoaxluolrnovbq.supabase.co/rest/v1/pauta_audiencias?select=dados&id=eq.global"
PAUTA_INDEX = pathlib.Path(r"C:\Users\ach\temp_mobile_fix\pauta-audiencias\index.html")   # de onde vem a chave pública do site
NAO = {"", "N/A", "NAN", "NONE", "NÃO IREMOS ACOMPANHAR"}


def _limpa_tipo(t):
    return re.sub(r"\s+", " ", str(t or "").replace("\xa0", " ").replace("\x96", "-")).strip()


def categoria(t):
    t = t.lower()
    if "julgamento - vara" in t or "diário oficial" in t or "diario oficial" in t:
        return None   # data de publicação de sentença, não é compromisso com comparecimento
    if "perícia" in t or "pericia" in t or "vistoria" in t:
        return "Perícia"
    if "sess" in t or "julgamento" in t:
        return "Sessão de Julgamento"
    return "Audiência"


def modalidade(t):
    t = t.lower()
    return "Semipresencial" if "semipresencial" in t else "Telepresencial" if "telepresencial" in t else "Presencial" if "presencial" in t else "Não informada"


def _compilado():
    arqs = [f for f in AQUI.glob("Compilado Audiências - 20*.xlsx") if not f.name.startswith("~$")]
    melhor, df_melhor = None, None
    for f in arqs:   # usa o Compilado que for mais longe (às vezes chega uma cópia antiga com outro nome)
        tmp = pathlib.Path(tempfile.gettempdir()) / f"aud_{abs(hash(f.name))}.xlsx"
        try:
            shutil.copy(f, tmp)
        except Exception:
            pass
        for h in (1, 0):
            try:
                d = pd.read_excel(tmp, header=h)
            except Exception:
                continue
            if "Data" in d.columns and "Advogado Responsável Audiência/SO" in d.columns:
                d["dt"] = pd.to_datetime(d["Data"], errors="coerce")
                if df_melhor is None or d["dt"].max() > df_melhor["dt"].max():
                    melhor, df_melhor = f.name, d
                break
    if df_melhor is None:
        return pd.DataFrame(), None
    d = df_melhor[df_melhor["dt"].notna()].copy()
    d = d.drop_duplicates(["Processo", "Data", "Horário"])
    out = pd.DataFrame({"dt": d["dt"], "horario": d["Horário"].astype(str), "processo": d["Processo"].astype(str).str.strip(),
                        "tipo": d["Tipo Agendamento"].map(_limpa_tipo), "adv": d["Advogado Responsável Audiência/SO"].astype(str).str.strip(),
                        "equipe": d.get("Equipe responsável", pd.Series("", index=d.index)).astype(str).str.strip(), "grupo_proc": "", "fonte": "Compilado"})
    return out, melhor


def _pauta():
    m = re.search(r"const SUPABASE_KEY = '([^']+)'", PAUTA_INDEX.read_text(encoding="utf-8"))
    chave = m.group(1)
    req = urllib.request.Request(PAUTA_URL, headers={"apikey": chave, "Authorization": f"Bearer {chave}"})
    dados = json.load(urllib.request.urlopen(req, timeout=60))[0]["dados"]
    des = dados.get("designacoes") or {}
    reg = {}
    for x in (dados.get("historico") or []) + (dados.get("pauta_main") or []):
        reg[(str(x.get("processo")), str(x.get("data")), str(x.get("horario")))] = x   # a pauta atual vem depois e prevalece
    linhas = []
    for (proc, data, hora), x in reg.items():
        adv = str(des.get(str(x.get("codigo"))) or x.get("adv_audiencia") or "").strip()
        linhas.append({"dt": pd.to_datetime(data, dayfirst=True, errors="coerce"), "horario": hora, "processo": proc.strip(),
                       "tipo": _limpa_tipo(x.get("tipo")), "adv": adv, "equipe": "", "grupo_proc": str(x.get("grupo") or "").strip(), "fonte": "Pauta"})
    return pd.DataFrame(linhas)


def carregar(ano=None):
    hoje = pd.Timestamp(dt.date.today())
    ano = ano or hoje.year
    comp, arq = _compilado()
    corte = comp["dt"].max() if not comp.empty else pd.Timestamp(ano, 1, 1) - pd.Timedelta(days=1)
    try:
        pau = _pauta()
        pau = pau[pau["dt"] > corte]
        erro_pauta = ""
    except Exception as e:
        pau, erro_pauta = pd.DataFrame(), f"{type(e).__name__}: {e}"
    df = pd.concat([comp, pau], ignore_index=True)
    df = df[df["dt"].dt.year.eq(ano)].copy()
    df["cat"] = df["tipo"].map(categoria)
    df = df[df["cat"].notna()]
    up = df["adv"].str.upper()
    df = df[~up.isin(NAO)]                                        # "N/A": o escritório não acompanha
    df["pendente"] = df["adv"].str.lower().eq("indicar advogado")  # ainda sem advogado designado
    df["mod"] = df["tipo"].map(modalidade)
    df = df.drop_duplicates(["processo", "dt", "horario"]).sort_values("dt")
    info = {"compilado": arq, "compilado_ate": corte.strftime("%d/%m/%Y") if not comp.empty else None,
            "pauta_de": (corte + pd.Timedelta(days=1)).strftime("%d/%m/%Y"), "ate": hoje.strftime("%d/%m/%Y"), "erro_pauta": erro_pauta}
    return df, info


if __name__ == "__main__":
    df, info = carregar()
    h = pd.Timestamp(dt.date.today())
    print(info)
    feitos = df[df["dt"] <= h]
    print("até hoje:", len(feitos), "| por fonte:", feitos["fonte"].value_counts().to_dict(), "| pendentes (indicar advogado):", int(feitos["pendente"].sum()),
          "| futuros:", int((df["dt"] > h).sum()), "| futuros sem advogado:", int(df[(df["dt"] > h)]["pendente"].sum()))
    print(feitos.groupby(feitos["dt"].dt.month).size().to_dict())
