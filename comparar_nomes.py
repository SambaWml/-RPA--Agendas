import pandas as pd
from pathlib import Path
from gerar_evidencia import gerar_relatorio


BASE_DIR = Path(r"C:\Users\wesml\OneDrive\Documentos\comparação")


def encontrar_arquivo(pattern: str) -> Path | None:
    """Encontra o primeiro .xlsx que combina com o padrão (ex: 'buyers', 'exhibitors')."""
    pattern = pattern.lower()
    candidatos = []
    for f in BASE_DIR.glob("*.xlsx"):
        if pattern in f.name.lower():
            candidatos.append(f)
    return candidatos[0] if candidatos else None


def encontrar_pasta(pattern: str) -> Path | None:
    """Encontra a primeira pasta que contenha o padrão no nome."""
    pattern = pattern.lower()
    candidatos = []
    for p in BASE_DIR.iterdir():
        if p.is_dir() and pattern in p.name.lower():
            candidatos.append(p)
    return candidatos[0] if candidatos else None


def normalizar_nome(nome) -> str:
    if pd.isna(nome):
        return ""
    if not isinstance(nome, str):
        nome = str(nome)
    
    nome = nome.lower()
    if nome.startswith("agenda_"):
        nome = nome[len("agenda_"):]
        
    chars = []
    for char in nome:
        if char.isalnum():
            chars.append(char)
        else:
            chars.append(' ')
    nome = "".join(chars)
    
    return " ".join(nome.split())


def detectar_coluna_nome(df: pd.DataFrame, tipo: str) -> str:
    """
    Tenta detectar automaticamente a coluna de nome.
    tipo: 'buyer' ou 'exhibitor' (só para ajudar nos pesos).
    """
    colunas = [str(c) for c in df.columns]

    # palavras-chave de empresa têm peso maior que nome genérico
    kw_empresa = ["company", "agency", "supplier", "hotel", "exhibitor", "venue", "organiz"]
    if tipo == "buyer":
        kw_tipo = ["buyer", "host", "agency"]
    else:
        kw_tipo = ["exhibitor", "supplier", "hotel", "venue"]
    kw_nome = ["name", "full name", "nome", "contact", "contact name"]

    candidatos_score = []

    for col in colunas:
        col_lower = col.lower()
        score = 0

        # pontos por conter palavras de empresa (peso alto)
        for kw in kw_empresa:
            if kw in col_lower:
                score += 4

        # pontos por conter palavras do tipo específico
        for kw in kw_tipo:
            if kw in col_lower:
                score += 3

        # pontos por conter palavras genéricas de nome (peso baixo)
        for kw in kw_nome:
            if kw in col_lower:
                score += 1

        # se a coluna é de texto em muitos valores, também conta
        serie = df[col]
        num_notna = serie.notna().sum()
        if num_notna > 0:
            amostra = serie.dropna().head(20)
            qtde_str = amostra.apply(lambda v: isinstance(v, str)).sum()
            if qtde_str / max(len(amostra), 1) > 0.5:
                score += 1

        candidatos_score.append((score, col))

    candidatos_score.sort(reverse=True)  # maior score primeiro

    melhor_score, melhor_col = candidatos_score[0]
    print(f"   -> Coluna detectada para {tipo}: '{melhor_col}' (score={melhor_score})")
    return melhor_col


def coletar_nomes_de_pasta(raiz: Path) -> set[str]:
    """
    Varre todos os arquivos dentro de 'raiz' e pega o nome do arquivo sem extensão.
    """
    nomes = set()
    for path in raiz.rglob("*"):
        if path.is_file():
            nomes.add(normalizar_nome(path.stem))
    return {n for n in nomes if n}  # remove strings vazias


def comparar(tipo: str, nomes_excel: set[str], nomes_pastas: set[str]):
    somente_excel = sorted(nomes_excel - nomes_pastas)
    somente_pastas = sorted(nomes_pastas - nomes_excel)
    em_ambos = sorted(nomes_excel & nomes_pastas)

    print(f"\n===== {tipo.upper()} =====")
    print(f"Total no Excel: {len(nomes_excel)}")
    print(f"Total nas pastas: {len(nomes_pastas)}")
    print(f"Em ambos: {len(em_ambos)}")
    print(f"Somente no Excel: {len(somente_excel)}")
    print(f"Somente nas pastas: {len(somente_pastas)}")

    linhas = []
    for nome in somente_excel:
        linhas.append({"tipo": tipo, "nome": nome, "origem": "excel"})
    for nome in somente_pastas:
        linhas.append({"tipo": tipo, "nome": nome, "origem": "pastas"})
    for nome in em_ambos:
        linhas.append({"tipo": tipo, "nome": nome, "origem": "ambos"})
    return linhas


def processar(tipo: str):
    # tipo: 'buyer' ou 'exhibitor'
    print(f"\n>>> Processando {tipo}s")

    arq = encontrar_arquivo(tipo)
    if not arq:
        print(f"   [ERRO] Não encontrei arquivo .xlsx com '{tipo}' no nome na pasta.")
        return []

    print(f"   Arquivo Excel: {arq.name}")

    df = pd.read_excel(arq)

    col_nome = detectar_coluna_nome(df, tipo)
    nomes_excel = {normalizar_nome(v) for v in df[col_nome].dropna()}
    nomes_excel = {n for n in nomes_excel if n}  # remove vazios

    pasta = encontrar_pasta(tipo)
    if not pasta:
        print(f"   [ERRO] Não encontrei pasta extraída que contenha '{tipo}' no nome.")
        return []

    print(f"   Pasta agendas: {pasta.name}")

    nomes_pastas = coletar_nomes_de_pasta(pasta)

    return comparar(tipo, nomes_excel, nomes_pastas)


def main():
    print(f"Base: {BASE_DIR}")

    rel_buyers = processar("buyer")
    rel_exhibitors = processar("exhibitor")

    relatorio = rel_buyers + rel_exhibitors
    if not relatorio:
        print("\nNenhum dado para salvar (provavelmente faltou Excel ou pasta).")
        return

    saida = BASE_DIR / "relatorio_comparacao_nomes.csv"
    pd.DataFrame(relatorio).to_csv(saida, index=False, encoding="utf-8-sig")
    print(f"\nRelatório salvo em: {saida}")

    # Gera automaticamente a evidência e o bug report
    gerar_relatorio()


if __name__ == "__main__":
    main()

