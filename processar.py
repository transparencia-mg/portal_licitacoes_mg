import copy
import io
import math
import re
import subprocess
from pathlib import Path

import pandas as pd
import yaml  # pip install pyyaml

# ==========================
# CONFIGURAÇÃO
# ==========================

REPO = Path(__file__).parent
UPLOAD = REPO / "upload"
DATASET = REPO / "dataset"
DATA = DATASET / "data"                      # onde ficam os CSVs lidos pelo index.html
DATAPACKAGE = DATASET / "datapackage.yaml"   # ajuste se o yaml estiver em outro lugar

anos = [2022, 2023, 2024, 2025, 2026]
ABAS = ["licitacoes", "item"]
TITULOS = {"licitacoes": "Licitações", "item": "Itens"}

# Limite por arquivo (margem de segurança abaixo de 10 MB).
# Vale tanto para o .xlsx quanto para o .csv equivalente.
LIMITE_BYTES = int(9.5 * 1024 * 1024)


# ==========================
# PARTICIONAMENTO
# ==========================

def gerar_xlsx_e_medir_csv(df):
    """Gera o .xlsx em memória e mede quanto o mesmo trecho ocupa em CSV."""
    buffer = io.BytesIO()
    df.to_excel(buffer, index=False)
    tamanho_csv = len(df.to_csv(sep=";", index=False).encode("utf-8-sig"))
    return buffer.getvalue(), tamanho_csv


def dividir_em_partes(df):
    """
    Devolve uma lista de arquivos .xlsx (bytes). Cada parte fica abaixo do
    limite tanto em xlsx quanto em csv. Se uma parte ainda passar, é
    dividida de novo.
    """
    dados, tamanho_csv = gerar_xlsx_e_medir_csv(df)
    maior = max(len(dados), tamanho_csv)

    if maior <= LIMITE_BYTES or len(df) <= 1:
        return [dados]

    n_partes = math.ceil(maior * 1.1 / LIMITE_BYTES)   # +10% de folga
    linhas_por_parte = math.ceil(len(df) / n_partes)

    partes = []
    for inicio in range(0, len(df), linhas_por_parte):
        partes.extend(dividir_em_partes(df.iloc[inicio:inicio + linhas_por_parte]))
    return partes


def nomes_das_partes(aba, ano, quantidade):
    """licitacoes2026 (arquivo único) ou licitacoes2026_1, licitacoes2026_2..."""
    if quantidade == 1:
        return [f"{aba}{ano}"]
    return [f"{aba}{ano}_{i}" for i in range(1, quantidade + 1)]


def remover_antigos(pasta, aba, ano, extensao, manter=()):
    """
    Apaga arquivos de execuções anteriores (aba+ano, com ou sem _N),
    exceto os que serão gerados agora (manter).
    """
    if not pasta.exists():
        return
    padrao = re.compile(rf"^{aba}{ano}(_\d+)?\.{extensao}$")
    for f in pasta.glob(f"{aba}{ano}*.{extensao}"):
        if padrao.match(f.name) and f.stem not in manter:
            f.unlink()
            print(f"Removido antigo: {f.relative_to(REPO)}")


# ==========================
# DATAPACKAGE.YAML
# ==========================

class IndentDumper(yaml.SafeDumper):
    """Mantém o estilo do yaml original (itens de lista indentados)."""
    def increase_indent(self, flow=False, indentless=False):
        return super().increase_indent(flow, False)


def atualizar_datapackage(pacote, aba, ano, nomes):
    """
    Substitui no yaml os recursos de aba+ano (com ou sem _N) por um recurso
    por parte, reaproveitando o schema do recurso que já existia.
    """
    recursos = pacote.setdefault("resources", [])
    padrao_ano = re.compile(rf"^{aba}{ano}(_\d+)?$")
    padrao_aba = re.compile(rf"^{aba}\d{{4}}(_\d+)?$")

    posicoes = [i for i, r in enumerate(recursos) if padrao_ano.match(r.get("name", ""))]

    if posicoes:
        modelo = recursos[posicoes[0]]
        posicao = posicoes[0]
    else:
        # ano novo: usa como modelo o último recurso da mesma aba
        mesma_aba = [i for i, r in enumerate(recursos) if padrao_aba.match(r.get("name", ""))]
        if not mesma_aba:
            print(f"Aviso: nenhum recurso '{aba}' no datapackage para servir de modelo.")
            return
        modelo = recursos[mesma_aba[-1]]
        posicao = mesma_aba[-1] + 1

    novos = []
    for i, nome in enumerate(nomes, 1):
        r = copy.deepcopy(modelo)
        r["name"] = nome
        r["path"] = f"data/{nome}.csv"
        r["title"] = f"{TITULOS[aba]} {ano}" + (f" - Parte {i}" if len(nomes) > 1 else "")
        r.pop("hash", None)    # valores do arquivo antigo não valem mais
        r.pop("bytes", None)
        novos.append(r)

    recursos[:] = [r for i, r in enumerate(recursos) if i not in posicoes]
    recursos[posicao:posicao] = novos


# ==========================
# PROCESSAMENTO
# ==========================

pacote = None
if DATAPACKAGE.exists():
    with open(DATAPACKAGE, encoding="utf-8") as f:
        pacote = yaml.safe_load(f)
else:
    print(f"Aviso: {DATAPACKAGE} não encontrado; o yaml não será atualizado.")

arquivos_gerados = 0

for ano in anos:

    arquivo = UPLOAD / f"LICITACOES_PORTAL{ano}.xlsx"

    if not arquivo.exists():
        print(f"Arquivo não encontrado: {arquivo.name}")
        continue

    print(f"Processando {arquivo.name}...")

    abas = pd.read_excel(
        arquivo,
        sheet_name=None,
        header=1      # remove primeira linha
    )

    for aba in ABAS:

        if aba not in abas:
            print(f"Aba '{aba}' não encontrada em {arquivo.name}")
            continue

        df = abas[aba]

        # remove primeira coluna
        df = df.iloc[:, 1:]

        # remove linhas totalmente vazias
        df = df.dropna(how="all")

        # remove colunas totalmente vazias
        df = df.dropna(axis=1, how="all")

        partes = dividir_em_partes(df)
        nomes = nomes_das_partes(aba, ano, len(partes))

        # limpa saídas antigas: xlsx em upload/ e CSVs que não serão mais gerados
        remover_antigos(UPLOAD, aba, ano, "xlsx")
        remover_antigos(DATA, aba, ano, "csv", manter=nomes)

        for nome, dados in zip(nomes, partes):
            saida = UPLOAD / f"{nome}.xlsx"
            saida.write_bytes(dados)
            print(f"Gerado: {saida.name} ({len(dados) / 1024 / 1024:.2f} MB)")
            arquivos_gerados += 1

        if pacote is not None:
            atualizar_datapackage(pacote, aba, ano, nomes)

    # remove arquivo original do ano
    arquivo.unlink()
    print(f"Removido: {arquivo.name}")

if pacote is not None and arquivos_gerados > 0:
    with open(DATAPACKAGE, "w", encoding="utf-8") as f:
        yaml.dump(pacote, f, Dumper=IndentDumper, sort_keys=False,
                  allow_unicode=True, default_flow_style=False)
    print(f"Atualizado: {DATAPACKAGE.relative_to(REPO)}")

# ==========================
# GIT
# ==========================

def git(*args):
    """Roda um comando git e mostra o erro se falhar (antes o erro ficava escondido)."""
    r = subprocess.run(["git", *args], cwd=REPO, capture_output=True, text=True)
    if r.stdout.strip():
        print(r.stdout.strip())
    if r.returncode != 0:
        print(f"ERRO no 'git {' '.join(args)}' (código {r.returncode}):")
        print(r.stderr.strip())
    return r


if arquivos_gerados > 0:

    caminhos = [p.relative_to(REPO).as_posix() for p in (UPLOAD, DATASET) if p.exists()]

    status = git("status", "--porcelain", "--", *caminhos)

    if status.returncode != 0:
        print("Não foi possível verificar o repositório; nada foi enviado ao GitHub.")

    elif not status.stdout.strip():
        print("Nenhuma alteração encontrada.")

    else:
        # -A registra também os arquivos apagados (partes antigas)
        ok = (
            git("add", "-A", "--", *caminhos).returncode == 0
            and git("commit", "-m", "Atualização automática portal licitações").returncode == 0
            and git("push").returncode == 0
        )
        print("GitHub atualizado com sucesso." if ok else "Falha ao enviar ao GitHub (veja o erro acima).")
