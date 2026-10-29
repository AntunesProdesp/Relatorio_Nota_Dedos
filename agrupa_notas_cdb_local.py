# -*- coding: utf-8 -*-
"""
Agrupamento dos consolidados diarios CDB_LOCAL_AAAAMMDD.

O que faz
---------
1. Procura na pasta de origem todos os arquivos CDB_LOCAL_AAAAMMDD
   (.xlsx, .xlsm, .xlsb, .xls ou .csv).
2. Le so as colunas necessarias (bem mais rapido do que abrir o arquivo inteiro).
3. Monta uma linha por coleta com: CLIENTE, NOME_POSTO, FINALIDADE, LOCALIDADE,
   HOSTNAME, IP_ESTACAO (dadosColeta.ipEstacao), CPF_CIDADAO,
   ID_COLETA, ANO, MES_ANO, DATA e as colunas dos 10 dedos
   ("1 - polegar da mao direita" ... "10 - minimo da mao esquerda").
   A DATA vem da coluna dataInclusao (ver ORIGEM_DATA).
4. Filtra o periodo pela dataInclusao (opcional: DATA_INICIO / DATA_FIM).
5. Salva um arquivo por cliente e por mes em .csv (e em .xlsx, se couber no limite do Excel).

Uso
---
    python agrupa_notas_cdb_local.py
    python agrupa_notas_cdb_local.py "C:\\outra\\pasta"     (opcional: outra pasta de origem)

Requisitos
----------
    pip install pandas openpyxl
    pip install python-calamine     (opcional: deixa a leitura dos .xlsx bem mais rapida)
"""

import os
import re
import sys
import time
import unicodedata

import pandas as pd

# ============================== CONFIGURACAO ==============================
PASTA_ORIGEM = r"T:\Corporativo\Bases\Setoriais\BI - Coleta Biométrica Unificada\CBU\Dashboard_CBU\CDB LOCAL\2026"

# Onde salvar o resultado. None = subpasta "_AGRUPADO" dentro da pasta de origem.
PASTA_SAIDA = r"T:\Corporativo\Bases\Setoriais\BI - Coleta Biométrica Unificada\CBU\Relatorios_Automatizados"
NOME_SAIDA = "AGRUPADO_CDB_LOCAL"

GERAR_XLSX = True          # alem do .csv, gera .xlsx (se couber em 1.048.575 linhas)
BUSCAR_SUBPASTAS = False   # True = procura os CDB_LOCAL tambem nas subpastas

SEPARAR_POR_CLIENTE = True   # True = um arquivo por CLIENTE (AGRUPADO_CDB_LOCAL_<CLIENTE>.csv)
GERAR_ARQUIVO_GERAL = False  # True = gera tambem o arquivo com todos os clientes juntos
SEPARAR_POR_MES = True       # True = um arquivo por mes da DATA (..._AAAA-MM.csv); combina com o cliente

# De onde vem a DATA do agrupamento:
#   "inclusao" -> coluna dataInclusao (se vazia/invalida, usa a data do nome do arquivo)
#   "coleta"   -> coluna dadosColeta.data (se vazia/invalida, usa a data do nome do arquivo)
#   "arquivo"  -> data do nome do arquivo (CDB_LOCAL_20261023 -> 23/10/2026)
ORIGEM_DATA = "inclusao"

# Filtro de periodo pela DATA (dataInclusao). None = sem limite. Formato DD/MM/AAAA.
DATA_INICIO = None   # ex.: "01/10/2026"
DATA_FIM = None      # ex.: "31/10/2026"

# Se a mesma idColeta aparecer em mais de um consolidado, mantem so a do arquivo
# mais recente (evita a mesma coleta em duas linhas).
REMOVER_COLETAS_DUPLICADAS = True

FUSO_HORARIO = "America/Sao_Paulo"  # usado so quando a data vem com fuso (ex.: ...T14:33:00Z)
# ==========================================================================


def normalizar(nome):
    """' dadosColeta.data ' -> 'DADOSCOLETA.DATA' (sem acento, sem espaco sobrando, maiusculo)."""
    nome = unicodedata.normalize("NFKD", str(nome)).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", nome).strip().upper()


# Colunas dos dedos, com o nome exato que vai no resultado
COLUNAS_DEDOS = [
    "1 - polegar da mao direita",
    "2 - indicador da mao direita",
    "3 - medio da mao direita",
    "4 - anular da mao direita",
    "5 - minimo da mao direita",
    "6 - polegar da mao esquerda",
    "7 - indicador da mao esquerda",
    "8 - medio da mao esquerda",
    "9 - anular da mao esquerda",
    "10 - minimo da mao esquerda",
]
DEDOS_NORM = {normalizar(c): c for c in COLUNAS_DEDOS}

# nome no resultado -> nome da coluna no consolidado (normalizado: sem acento, maiusculo)
COLUNAS_ORIGEM = {
    "CLIENTE": "CLIENTE",
    "NOME_POSTO": "NOME FANTASIA",
    "FINALIDADE": "FINALIDADE",
    "LOCALIDADE": "LOCALIDADE",
    "HOSTNAME": "HOSTNAME",
    "IP_ESTACAO": "DADOSCOLETA.IPESTACAO",
    "CPF_CIDADAO": "CPF",
    "ID_COLETA": "IDCOLETA",
    "DATA_COLETA": "DADOSCOLETA.DATA",
    "DATA_INCLUSAO": "DATAINCLUSAO",
}
# ORIGEM_DATA -> (coluna normalizada, nome original para as mensagens)
COLUNAS_DATA = {
    "inclusao": ("DATAINCLUSAO", "dataInclusao"),
    "coleta": ("DADOSCOLETA.DATA", "dadosColeta.data"),
}
CHAVES = ["CLIENTE", "NOME_POSTO", "FINALIDADE", "LOCALIDADE", "HOSTNAME", "IP_ESTACAO", "CPF_CIDADAO", "ID_COLETA", "ANO", "MES_ANO", "DATA"]
COLUNAS_RESULTADO = CHAVES + COLUNAS_DEDOS

COLUNAS_LIDAS = set(COLUNAS_ORIGEM.values()) | set(DEDOS_NORM)

PADRAO_ARQUIVO = re.compile(r"^CDB_LOCAL_(\d{8})\.(xlsx|xlsm|xlsb|xls|csv)$", re.IGNORECASE)
LIMITE_EXCEL = 1_048_575


# ------------------------------------------------------------------ leitura
def listar_arquivos(pasta):
    """Retorna [(data_do_arquivo, caminho)] em ordem de data, um arquivo por data."""
    if BUSCAR_SUBPASTAS:
        caminhos = [os.path.join(r, f) for r, _, fs in os.walk(pasta) for f in fs]
    else:
        caminhos = [os.path.join(pasta, f) for f in os.listdir(pasta)]

    por_data = {}
    for caminho in caminhos:
        m = PADRAO_ARQUIVO.match(os.path.basename(caminho))
        if not m or not os.path.isfile(caminho):
            continue
        data = pd.to_datetime(m.group(1), format="%Y%m%d", errors="coerce")
        if pd.isna(data):
            continue
        # dois arquivos da mesma data (ex.: .csv e .xlsx) -> fica o modificado por ultimo
        atual = por_data.get(data)
        if atual is None or os.path.getmtime(caminho) > os.path.getmtime(atual):
            por_data[data] = caminho
    return sorted(por_data.items())


def ler_arquivo(caminho):
    """Le so as colunas usadas. Retorna (DataFrame com nomes normalizados, colunas que faltaram)."""
    filtro = lambda c: normalizar(c) in COLUNAS_LIDAS  # noqa: E731

    if caminho.lower().endswith(".csv"):
        with open(caminho, "rb") as f:
            cabecalho = f.readline().decode("latin-1")
        sep = max([";", ",", "\t", "|"], key=cabecalho.count)
        try:
            df = pd.read_csv(caminho, sep=sep, usecols=filtro, dtype=str, encoding="utf-8-sig")
        except UnicodeDecodeError:
            df = pd.read_csv(caminho, sep=sep, usecols=filtro, dtype=str, encoding="latin-1")
    else:
        try:
            df = pd.read_excel(caminho, usecols=filtro, dtype=str, engine="calamine")
        except (ImportError, ValueError):
            df = pd.read_excel(caminho, usecols=filtro, dtype=str)

    df.columns = [normalizar(c) for c in df.columns]
    df = df.loc[:, ~df.columns.duplicated()]
    faltando = sorted(COLUNAS_LIDAS - set(df.columns))
    for c in faltando:
        df[c] = None
    return df, faltando


# ----------------------------------------------------------- tratamentos
def texto(serie):
    s = serie.astype(object).where(serie.notna(), "")
    return s.astype(str).str.strip()


def texto_id(serie):
    """Remove o '.0' que o Excel coloca em numeros lidos como texto."""
    return texto(serie).str.replace(r"\.0+$", "", regex=True)


def limpar_cpf(serie):
    """Deixa so os digitos e completa os zeros a esquerda (11 digitos)."""
    cpf = texto_id(serie).str.replace(r"\D", "", regex=True)
    return cpf.where((cpf == "") | (cpf.str.len() > 11), cpf.str.zfill(11))


def numerico_se_possivel(serie):
    """Se todos os valores preenchidos da coluna forem numeros, devolve como numero
    (para somar/filtrar no Excel e no Power BI); senao, mantem o texto original."""
    preenchido = serie != ""
    if not preenchido.any():
        return serie
    num = pd.to_numeric(serie.str.replace(",", ".", regex=False), errors="coerce")
    if num[preenchido].isna().any():
        return serie
    if (num.dropna() % 1 == 0).all():
        return num.astype("Int64")
    return num


def converter_data(serie):
    """Converte dadosColeta.data para data, aceitando os formatos mais comuns:
    2026-10-23 14:33:00 | 2026-10-23T17:33:00.000Z | 23/10/2026 14:33 | 46318.6 (numero do Excel)."""
    s = texto(serie)
    out = pd.Series(pd.NaT, index=s.index, dtype="datetime64[ns]")

    iso = s.str.match(r"^\d{4}-\d{2}-\d{2}")
    com_fuso = iso & s.str.contains(r"(?:Z|[+-]\d{2}:?\d{2})$")
    sem_fuso = iso & ~com_fuso
    br = s.str.match(r"^\d{1,2}/\d{1,2}/\d{4}")
    numero = s.str.fullmatch(r"\d{4,6}(?:[.,]\d+)?")
    resto = (s != "") & ~iso & ~br & ~numero

    if com_fuso.any():
        try:
            d = pd.to_datetime(s[com_fuso], errors="coerce", utc=True, format="ISO8601")
        except (TypeError, ValueError):
            d = pd.to_datetime(s[com_fuso], errors="coerce", utc=True)
        out[com_fuso] = d.dt.tz_convert(FUSO_HORARIO).dt.tz_localize(None)
    if sem_fuso.any():
        out[sem_fuso] = pd.to_datetime(s[sem_fuso].str[:10], format="%Y-%m-%d", errors="coerce")
    if br.any():
        parte = s[br].str.extract(r"^(\d{1,2})/(\d{1,2})/(\d{4})")
        out[br] = pd.to_datetime(
            parte[2] + "-" + parte[1].str.zfill(2) + "-" + parte[0].str.zfill(2),
            format="%Y-%m-%d", errors="coerce",
        )
    if numero.any():
        n = pd.to_numeric(s[numero].str.replace(",", ".", regex=False), errors="coerce")
        out[numero] = pd.to_datetime(n, unit="D", origin="1899-12-30", errors="coerce")
    if resto.any():
        out[resto] = pd.to_datetime(s[resto], errors="coerce", dayfirst=True)

    # descarta datas absurdas (vao para a data do arquivo)
    out = out.where((out >= "2000-01-01") & (out <= "2100-12-31"))
    return out.dt.normalize()


def processar(data_arquivo, caminho):
    """Uma linha por coleta, com as chaves e as 10 colunas dos dedos."""
    df, faltando = ler_arquivo(caminho)
    avisos = []
    if faltando:
        nomes = [DEDOS_NORM.get(c, c) for c in faltando]
        avisos.append("colunas nao encontradas: " + ", ".join(nomes))

    r = pd.DataFrame(index=df.index)
    r["CLIENTE"] = texto(df["CLIENTE"])
    r["NOME_POSTO"] = texto(df["NOME FANTASIA"])
    r["FINALIDADE"] = texto(df["FINALIDADE"])
    r["LOCALIDADE"] = texto(df["LOCALIDADE"])
    r["HOSTNAME"] = texto(df["HOSTNAME"])
    r["IP_ESTACAO"] = texto(df["DADOSCOLETA.IPESTACAO"])
    r["CPF_CIDADAO"] = limpar_cpf(df["CPF"])
    r["ID_COLETA"] = texto_id(df["IDCOLETA"])

    if ORIGEM_DATA == "arquivo":
        data = pd.Series(data_arquivo, index=df.index)
    else:
        col, nome_col = COLUNAS_DATA[ORIGEM_DATA]
        data = converter_data(df[col])
        sem_data = int(data.isna().sum())
        if sem_data and col not in faltando:
            avisos.append(f"{sem_data} linha(s) sem {nome_col} valida -> usada a data do arquivo")
        data = data.fillna(data_arquivo)

    r["ANO"] = data.dt.year.astype(int)
    r["MES_ANO"] = data.dt.strftime("%m/%Y")
    r["DATA"] = data

    for norm, nome in DEDOS_NORM.items():
        r[nome] = texto_id(df[norm])

    return r, avisos


# ------------------------------------------------------------------ main
def nome_arquivo(texto_cliente):
    """'Secretaria da Saúde / SP' -> 'SECRETARIA_DA_SAUDE_SP' (seguro para nome de arquivo)."""
    nome = re.sub(r"[^A-Z0-9]+", "_", normalizar(texto_cliente)).strip("_")
    return nome[:80] or "SEM_CLIENTE"


def filtrar_periodo(resultado):
    """Mantem so as linhas com DATA (dataInclusao) entre DATA_INICIO e DATA_FIM."""
    if DATA_INICIO:
        resultado = resultado[resultado["DATA"] >= pd.to_datetime(DATA_INICIO, dayfirst=True)]
    if DATA_FIM:
        resultado = resultado[resultado["DATA"] <= pd.to_datetime(DATA_FIM, dayfirst=True)]
    return resultado.reset_index(drop=True)


def salvar(resultado, pasta_saida, nome_saida=NOME_SAIDA):
    os.makedirs(pasta_saida, exist_ok=True)
    gerados = []

    csv = os.path.join(pasta_saida, nome_saida + ".csv")
    try:
        resultado.to_csv(csv, sep=";", index=False, encoding="utf-8-sig", date_format="%d/%m/%Y")
        gerados.append(csv)
    except PermissionError:
        print(f"  !! Nao consegui gravar {csv} (arquivo aberto?). Feche e rode de novo.")

    if GERAR_XLSX:
        if len(resultado) > LIMITE_EXCEL:
            print(f"  .xlsx nao gerado: {len(resultado):,} linhas passam do limite do Excel (use o .csv).")
        else:
            xlsx = os.path.join(pasta_saida, nome_saida + ".xlsx")
            col_data = resultado.columns.get_loc("DATA") + 1
            try:
                with pd.ExcelWriter(xlsx, engine="openpyxl") as w:
                    resultado.to_excel(w, index=False, sheet_name="AGRUPADO")
                    ws = w.sheets["AGRUPADO"]
                    for (cel,) in ws.iter_rows(min_row=2, min_col=col_data, max_col=col_data):
                        cel.number_format = "DD/MM/YYYY"
                    ws.freeze_panes = "A2"
                    ws.auto_filter.ref = ws.dimensions
                gerados.append(xlsx)
            except PermissionError:
                print(f"  !! Nao consegui gravar {xlsx} (arquivo aberto?). Feche e rode de novo.")
    return gerados


def salvar_por_mes(resultado, pasta_saida, nome_saida):
    """Um arquivo por mes (DATA): <nome_saida>_AAAA-MM.csv / .xlsx."""
    gerados = []
    meses = resultado["DATA"].dt.strftime("%Y-%m")
    for mes, parte in resultado.groupby(meses, sort=True):
        print(f"    {mes}: {len(parte):,} linha(s)")
        gerados += salvar(parte.reset_index(drop=True), pasta_saida, f"{nome_saida}_{mes}")
    return gerados


def main():
    pasta = sys.argv[1] if len(sys.argv) > 1 else PASTA_ORIGEM
    pasta_saida = PASTA_SAIDA or os.path.join(pasta, "_AGRUPADO")
    inicio = time.time()

    if not os.path.isdir(pasta):
        print(f"Pasta nao encontrada: {pasta}\n(o drive T: esta mapeado nesta maquina?)")
        sys.exit(1)

    arquivos = listar_arquivos(pasta)
    if not arquivos:
        print(f"Nenhum arquivo CDB_LOCAL_AAAAMMDD encontrado em: {pasta}")
        sys.exit(1)

    print(f"{len(arquivos)} consolidado(s) encontrado(s) em {pasta}\n")
    partes, erros = [], []
    for i, (data_arq, caminho) in enumerate(arquivos, 1):
        nome = os.path.basename(caminho)
        t = time.time()
        try:
            parte, avisos = processar(data_arq, caminho)
            partes.append(parte)
            print(f"[{i}/{len(arquivos)}] {nome}: {len(parte):,} coletas ({time.time() - t:.1f}s)")
            for a in avisos:
                print(f"      aviso: {a}")
        except Exception as e:  # um arquivo com problema nao derruba o resto
            erros.append(nome)
            print(f"[{i}/{len(arquivos)}] {nome}: ERRO -> {type(e).__name__}: {e}")

    if not partes:
        print("\nNenhum arquivo pode ser lido.")
        sys.exit(1)

    base = pd.concat(partes, ignore_index=True)

    if REMOVER_COLETAS_DUPLICADAS:
        dup = (base["ID_COLETA"] != "") & base.duplicated("ID_COLETA", keep="last")
        if dup.any():
            print(f"\n{int(dup.sum()):,} coleta(s) repetida(s) entre arquivos -> mantida a do arquivo mais recente.")
            base = base[~dup]

    # agrupamento: uma linha por combinacao de chaves + valores dos dedos (linhas identicas viram uma so)
    resultado = (
        base[COLUNAS_RESULTADO]
        .drop_duplicates()
        .sort_values(["DATA", "CLIENTE", "NOME_POSTO", "FINALIDADE", "CPF_CIDADAO", "ID_COLETA"])
        .reset_index(drop=True)
    )
    for c in COLUNAS_DEDOS:
        resultado[c] = numerico_se_possivel(resultado[c])

    if DATA_INICIO or DATA_FIM:
        antes = len(resultado)
        resultado = filtrar_periodo(resultado)
        print(f"\nFiltro dataInclusao {DATA_INICIO or '...'} a {DATA_FIM or '...'}: "
              f"{antes - len(resultado):,} linha(s) fora do periodo removida(s).")

    print(f"\nLinhas no resultado: {len(resultado):,}")
    if resultado.empty:
        print("Nada para salvar.")
        sys.exit(1)

    gerados = []
    if GERAR_ARQUIVO_GERAL or not (SEPARAR_POR_CLIENTE or SEPARAR_POR_MES):
        if SEPARAR_POR_MES:
            gerados += salvar_por_mes(resultado, pasta_saida, NOME_SAIDA)
        else:
            gerados += salvar(resultado, pasta_saida)
    if SEPARAR_POR_CLIENTE:
        print(f"\nSeparando por cliente ({resultado['CLIENTE'].nunique()} cliente(s)):")
        for cliente, parte in resultado.groupby("CLIENTE", sort=True):
            nome = f"{NOME_SAIDA}_{nome_arquivo(cliente)}"
            print(f"  {cliente or '(sem cliente)'}: {len(parte):,} linha(s)")
            if SEPARAR_POR_MES:
                gerados += salvar_por_mes(parte, pasta_saida, nome)
            else:
                gerados += salvar(parte.reset_index(drop=True), pasta_saida, nome)
    elif SEPARAR_POR_MES and not GERAR_ARQUIVO_GERAL:
        gerados += salvar_por_mes(resultado, pasta_saida, NOME_SAIDA)

    print()
    for g in gerados:
        print(f"Salvo: {g}")
    if erros:
        print(f"\nArquivos com erro ({len(erros)}): " + ", ".join(erros))
    print(f"\nConcluido em {time.time() - inicio:.1f}s")


if __name__ == "__main__":
    main()
