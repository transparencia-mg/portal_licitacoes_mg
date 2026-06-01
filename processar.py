import pandas as pd
from pathlib import Path
import subprocess
import os

REPO = Path(__file__).parent
UPLOAD = REPO / "upload"

anos = [2022, 2023, 2024, 2025, 2026]

arquivos_gerados = 0

for ano in anos:

    arquivo = UPLOAD / f"NOVACONSULTALICITACOES{ano}.xlsx"

    if not arquivo.exists():
        print(f"Arquivo não encontrado: {arquivo.name}")
        continue

    print(f"Processando {arquivo.name}...")

    abas = pd.read_excel(
        arquivo,
        sheet_name=None,
        header=1      # remove primeira linha
    )

    for aba in ["licitacoes", "item"]:

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

        saida = UPLOAD / f"{aba}{ano}.xlsx"

        df.to_excel(
            saida,
            index=False
        )

        print(f"Gerado: {saida.name}")

        arquivos_gerados += 1

    # remove arquivo original do ano
    os.remove(arquivo)
    print(f"Removido: {arquivo.name}")

# ==========================
# GIT
# ==========================

if arquivos_gerados > 0:

    resultado = subprocess.run(
        ["git", "status", "--porcelain"],
        cwd=REPO,
        capture_output=True,
        text=True
    )

    if resultado.stdout.strip():

        subprocess.run(
            ["git", "add", "upload"],
            cwd=REPO
        )

        subprocess.run(
            [
                "git",
                "commit",
                "-m",
                "Atualização automática portal licitações"
            ],
            cwd=REPO
        )

        subprocess.run(
            ["git", "push"],
            cwd=REPO
        )

        print("GitHub atualizado com sucesso.")

    else:
        print("Nenhuma alteração encontrada.")