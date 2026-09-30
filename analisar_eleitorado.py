# -*- coding: utf-8 -*-
"""
Perfis do eleitorado (TSE) em 3 escopos:
  1. ESTADO_SP      - todo o estado de São Paulo
  2. CAPITAL_SP     - cidade de São Paulo
  3. RMSP           - Região Metropolitana de São Paulo (39 municípios, inclui a capital)

Uso:  python analisar_eleitorado.py 2022
Gera: resultados/<ANO>/<ESCOPO>/<BASE>/*.csv  +  resultados/<ANO>/RESUMO_<ANO>_<ESCOPO>.xlsx
"""
import sys, os, re, zipfile, itertools, unicodedata, urllib.request
import pandas as pd

ANO = sys.argv[1] if len(sys.argv) > 1 else "2022"
UF = "SP"
BASE = "https://cdn.tse.jus.br/estatistica/sead/odsele"
OUT = f"resultados/{ANO}"
LIMITE_EXCEL = 1_000_000
CHUNK = 400_000
MIN_ELEITORES_RANKING = 1000   # perfis menores que isso ficam fora dos rankings

BASES = {
    "eleitorado_secao": f"{BASE}/perfil_eleitor_secao/perfil_eleitor_secao_{ANO}_{UF}.zip",
    "comparecimento":   f"{BASE}/perfil_comparecimento_abstencao/perfil_comparecimento_abstencao_{ANO}.zip",
    "transito_tte":     f"{BASE}/perfil_comparecimento_abstencao_eleitor_tte/perfil_comparecimento_abstencao_eleitor_tte_{ANO}.zip",
    "deficiencia":      f"{BASE}/perfil_comparecimento_abstencao_eleitor_deficiente/perfil_comparecimento_abstencao_eleitor_deficiente_{ANO}.zip",
}


def norm(txt):
    txt = unicodedata.normalize("NFKD", str(txt)).encode("ascii", "ignore").decode()
    return re.sub(r"[\s\-']+", " ", txt.upper()).strip()


# 39 municípios da RMSP (+ grafias alternativas usadas em cadastros)
RMSP = {norm(m) for m in [
    "Arujá", "Barueri", "Biritiba-Mirim", "Caieiras", "Cajamar", "Carapicuíba", "Cotia",
    "Diadema", "Embu das Artes", "Embu", "Embu-Guaçu", "Ferraz de Vasconcelos",
    "Francisco Morato", "Franco da Rocha", "Guararema", "Guarulhos", "Itapecerica da Serra",
    "Itapevi", "Itaquaquecetuba", "Jandira", "Juquitiba", "Mairiporã", "Mauá",
    "Mogi das Cruzes", "Moji das Cruzes", "Osasco", "Pirapora do Bom Jesus", "Poá",
    "Ribeirão Pires", "Rio Grande da Serra", "Salesópolis", "Santa Isabel",
    "Santana de Parnaíba", "Santo André", "São Bernardo do Campo", "São Caetano do Sul",
    "São Lourenço da Serra", "São Paulo", "Suzano", "Taboão da Serra", "Vargem Grande Paulista",
]}
CAPITAL = norm("São Paulo")

# Recortes conhecidos; qualquer outra coluna DS_ do arquivo também entra
DIMENSOES = [
    "DS_GENERO", "DS_IDENTIDADE_GENERO", "DS_COR_RACA", "DS_RACA_COR",
    "DS_FAIXA_ETARIA", "DS_GRAU_ESCOLARIDADE", "DS_ESTADO_CIVIL",
    "DS_QUILOMBOLA", "DS_INTERPRETE_LIBRAS", "DS_TIPO_DEFICIENCIA",
    "DS_TIPO_LOCAL", "DS_TIPO_SECAO", "DS_SITUACAO_LOCAL", "TIPO_VOTO",
]
PRIORITARIAS = ["DS_GENERO", "DS_FAIXA_ETARIA", "DS_GRAU_ESCOLARIDADE", "DS_ESTADO_CIVIL",
                "DS_COR_RACA", "DS_RACA_COR", "TIPO_VOTO", "DS_TIPO_DEFICIENCIA"]
IGNORAR_DS = ("MUNICIPIO", "ELEICAO", "_UF", "PAIS", "ENDERECO", "BAIRRO", "CARGO", "LOCAL_VOTACAO")

# Unidade geográfica usada em cada escopo
GEO_ESCOPO = {
    "ESTADO_SP":  [["NM_MUNICIPIO"], ["NM_MUNICIPIO", "NR_ZONA"]],
    "RMSP":       [["NM_MUNICIPIO"], ["NM_MUNICIPIO", "NR_ZONA"]],
    "CAPITAL_SP": [["NR_ZONA"]],
}


def log(*a):
    print(*a, flush=True)


def baixar(nome, url):
    os.makedirs("dados_brutos", exist_ok=True)
    destino = f"dados_brutos/{nome}_{ANO}.zip"
    if os.path.exists(destino):
        return destino
    log(f"Baixando {url}")
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=900) as r, open(destino, "wb") as f:
            while True:
                b = r.read(1 << 20)
                if not b:
                    break
                f.write(b)
        return destino
    except Exception as e:
        log(f"  !! Não consegui baixar {nome}: {e}")
        if os.path.exists(destino):
            os.remove(destino)
        return None


def csv_do_zip(caminho):
    z = zipfile.ZipFile(caminho)
    nomes = [n for n in z.namelist() if n.lower().endswith((".csv", ".txt"))]
    do_uf = [n for n in nomes if re.search(rf"_{UF}\.(csv|txt)$", n, re.I)]
    alvo = do_uf[0] if do_uf else max(nomes, key=lambda n: z.getinfo(n).file_size)
    return z, alvo


def coluna_municipio(cols):
    if "NM_MUNICIPIO" in cols:
        return "NM_MUNICIPIO"
    cand = [c for c in cols if c.startswith("NM_MUNICIPIO")]
    return cand[0] if cand else None


def coluna_uf(cols):
    if "SG_UF" in cols:
        return "SG_UF"
    cand = [c for c in cols if c.startswith("SG_UF")]
    return cand[0] if cand else None


def tipo_voto(df):
    idade = pd.to_numeric(
        df.get("DS_FAIXA_ETARIA", pd.Series("", index=df.index)).astype(str).str.extract(r"(\d+)")[0],
        errors="coerce")
    esc = df.get("DS_GRAU_ESCOLARIDADE", pd.Series("", index=df.index)).astype(str).str.upper()
    analf = esc.str.contains("ANALFABETO", na=False)
    tv = pd.Series("OBRIGATÓRIO", index=df.index)
    tv[(idade < 18) | (idade >= 70) | analf] = "FACULTATIVO"
    tv[idade.isna() & ~analf] = "INDEFINIDO"
    return tv


def montar_grupos(dims, extra, geo_opcoes):
    g = {"00_total": extra}
    for d in dims:
        g[f"1_{d}"] = extra + [d]
    for a, b in itertools.combinations(dims, 2):
        g[f"2_{a}_x_{b}"] = extra + [a, b]
    prio = [d for d in PRIORITARIAS if d in dims]
    for trio in itertools.combinations(prio, 3):
        g["3_" + "_x_".join(trio)] = extra + list(trio)
    for quad in itertools.combinations(prio, 4):
        g["4_" + "_x_".join(quad)] = extra + list(quad)
    # perfil completo: todas as dimensões juntas (base para tabela dinâmica)
    g["6_PERFIL_COMPLETO"] = extra + dims
    for geo in geo_opcoes:
        nome_geo = "_".join(geo)
        g[f"5_{nome_geo}"] = extra + geo
        for d in dims:
            g[f"5_{nome_geo}_x_{d}"] = extra + geo + [d]
    return g


def processar(nome, caminho):
    z, arq = csv_do_zip(caminho)
    log(f"Processando {nome}: {arq}")
    acumulado = {esc: {} for esc in GEO_ESCOPO}
    colunas_arquivo = None
    for bloco in pd.read_csv(z.open(arq), sep=";", encoding="latin-1", dtype=str,
                             chunksize=CHUNK, low_memory=False):
        if colunas_arquivo is None:
            colunas_arquivo = list(bloco.columns)
            log(f"  colunas: {', '.join(colunas_arquivo)}")
        cuf = coluna_uf(bloco.columns)
        if cuf:
            bloco = bloco[bloco[cuf] == UF]
        cmun = coluna_municipio(bloco.columns)
        if bloco.empty or cmun is None:
            continue
        if cmun != "NM_MUNICIPIO":
            bloco = bloco.rename(columns={cmun: "NM_MUNICIPIO"})
        medidas = [c for c in bloco.columns if c.startswith("QT_")]
        for c in medidas:
            bloco[c] = pd.to_numeric(bloco[c], errors="coerce").fillna(0)
        bloco["TIPO_VOTO"] = tipo_voto(bloco)
        extras_ds = [c for c in bloco.columns if c.startswith("DS_") and c not in DIMENSOES
                     and not any(k in c for k in IGNORAR_DS)]
        dims = [d for d in DIMENSOES if d in bloco.columns] + extras_ds
        extra = ["NR_TURNO"] if "NR_TURNO" in bloco.columns else []
        mun = bloco["NM_MUNICIPIO"].map(norm)
        mascaras = {"ESTADO_SP": pd.Series(True, index=bloco.index),
                    "RMSP": mun.isin(RMSP),
                    "CAPITAL_SP": mun == CAPITAL}

        for esc, masc in mascaras.items():
            sub = bloco[masc]
            if sub.empty:
                continue
            geo = [[c for c in g if c in sub.columns] for g in GEO_ESCOPO[esc]]
            geo = [g for g in geo if g]
            for nomeg, chaves in montar_grupos(dims, extra, geo).items():
                if chaves:
                    parcial = sub.groupby(chaves, dropna=False)[medidas].sum().reset_index()
                else:
                    parcial = sub[medidas].sum().to_frame().T
                acumulado[esc].setdefault(nomeg, []).append(parcial)
        log(f"  ... bloco: {len(bloco):,} linhas SP | RMSP {int(mascaras['RMSP'].sum()):,} | capital {int(mascaras['CAPITAL_SP'].sum()):,}")

    resultados = {}
    for esc, grupos in acumulado.items():
        resultados[esc] = {}
        for nomeg, partes in grupos.items():
            df = pd.concat(partes, ignore_index=True)
            chaves = [c for c in df.columns if not c.startswith("QT_")]
            df = df.groupby(chaves, dropna=False).sum().reset_index() if chaves else df.sum().to_frame().T
            resultados[esc][nomeg] = adicionar_taxas(df)
    return resultados, colunas_arquivo


def adicionar_taxas(df):
    cols = set(df.columns)
    if {"QT_ABSTENCAO", "QT_APTOS"} <= cols:
        df["PCT_ABSTENCAO"] = (100 * df["QT_ABSTENCAO"] / df["QT_APTOS"].replace(0, pd.NA)).round(2)
    if {"QT_COMPARECIMENTO", "QT_APTOS"} <= cols:
        df["PCT_COMPARECIMENTO"] = (100 * df["QT_COMPARECIMENTO"] / df["QT_APTOS"].replace(0, pd.NA)).round(2)
    for c in list(cols):
        m = re.match(r"QT_ABSTENCAO_(.+)", c)
        if m and f"QT_COMPARECIMENTO_{m.group(1)}" in cols:
            tot = df[c] + df[f"QT_COMPARECIMENTO_{m.group(1)}"]
            df[f"PCT_ABSTENCAO_{m.group(1)}"] = (100 * df[c] / tot.replace(0, pd.NA)).round(2)
    base = "QT_ELEITORES_PERFIL" if "QT_ELEITORES_PERFIL" in cols else ("QT_APTOS" if "QT_APTOS" in cols else None)
    if base:
        for c in [x for x in cols if x.startswith("QT_ELEITORES_") and x != base]:
            df[f"PCT_{c[3:]}"] = (100 * df[c] / df[base].replace(0, pd.NA)).round(2)
        # peso de cada linha no total do escopo (por turno, se houver)
        if "NR_TURNO" in df.columns:
            total = df.groupby("NR_TURNO")[base].transform("sum")
        else:
            total = df[base].sum()
        df["PCT_DO_TOTAL"] = (100 * df[base] / total).round(3)
    return df


def rankings(res):
    """Perfis com maior e menor abstenção (a partir do perfil completo)."""
    out = {}
    df = res.get("6_PERFIL_COMPLETO")
    if df is None or "PCT_ABSTENCAO" not in df.columns:
        return out
    df = df[df["QT_APTOS"] >= MIN_ELEITORES_RANKING]
    if "NR_TURNO" in df.columns:
        df = df[df["NR_TURNO"].astype(str) == "1"]
    out["7_RANKING_MAIOR_ABSTENCAO"] = df.sort_values("PCT_ABSTENCAO", ascending=False).head(100)
    out["7_RANKING_MENOR_ABSTENCAO"] = df.sort_values("PCT_ABSTENCAO").head(100)
    for geo in ["5_NM_MUNICIPIO", "5_NR_ZONA", "5_NM_MUNICIPIO_NR_ZONA"]:
        g = res.get(geo)
        if g is not None and "PCT_ABSTENCAO" in g.columns:
            if "NR_TURNO" in g.columns:
                g = g[g["NR_TURNO"].astype(str) == "1"]
            out[f"7_RANKING_ABSTENCAO_{geo[2:]}"] = g.sort_values("PCT_ABSTENCAO", ascending=False)
    return out


def main():
    dicionario = []
    todos = {esc: {} for esc in GEO_ESCOPO}
    for nome, url in BASES.items():
        caminho = baixar(nome, url)
        if not caminho:
            dicionario.append({"base": nome, "status": "NÃO BAIXADA", "url": url, "colunas": ""})
            continue
        try:
            res, cols = processar(nome, caminho)
            for esc, tabelas in res.items():
                tabelas.update(rankings(tabelas))
                pasta = f"{OUT}/{esc}/{nome}"
                os.makedirs(pasta, exist_ok=True)
                for nomeg, df in tabelas.items():
                    df.to_csv(f"{pasta}/{nomeg}.csv", sep=";", index=False,
                              encoding="utf-8-sig", decimal=",")
                todos[esc][nome] = tabelas
            dicionario.append({"base": nome, "status": "OK", "url": url, "colunas": ", ".join(cols or [])})
        except Exception as e:
            log(f"  !! Erro em {nome}: {e}")
            dicionario.append({"base": nome, "status": f"ERRO: {e}", "url": url, "colunas": ""})
        finally:
            if os.path.exists(caminho):
                os.remove(caminho)

    for esc, bases in todos.items():
        xlsx = f"{OUT}/RESUMO_{ANO}_{esc}.xlsx"
        with pd.ExcelWriter(xlsx, engine="openpyxl") as w:
            pd.DataFrame(dicionario).to_excel(w, sheet_name="LEIA-ME", index=False)
            indice, n = [], 0
            for base, tabelas in bases.items():
                for nomeg, df in tabelas.items():
                    if len(df) > LIMITE_EXCEL:
                        indice.append({"aba": "(só CSV)", "base": base, "tabela": nomeg, "linhas": len(df)})
                        continue
                    n += 1
                    aba = f"{n:03d}"
                    df.to_excel(w, sheet_name=aba, index=False)
                    indice.append({"aba": aba, "base": base, "tabela": nomeg, "linhas": len(df)})
            pd.DataFrame(indice).to_excel(w, sheet_name="INDICE", index=False)
        log(f"Gerado {xlsx}")


if __name__ == "__main__":
    main()
