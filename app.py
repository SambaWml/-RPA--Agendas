import io
import zipfile
import tempfile
import shutil
from datetime import datetime
from pathlib import Path

import pandas as pd
from flask import Flask, render_template, request, jsonify, send_file

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024  # 200 MB


# ── Lógica de comparação ─────────────────────────────────────────────────────

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
            chars.append(" ")
    return " ".join("".join(chars).split())


def detectar_coluna_nome(df: pd.DataFrame) -> str:
    """Retorna a coluna mais provável de conter o nome da empresa/agência."""
    # Palavras que indicam coluna de empresa (peso alto)
    kw_empresa = ["company", "agency", "supplier", "hotel", "exhibitor", "venue", "organiz"]
    # Palavras genéricas de nome (peso menor)
    kw_nome = ["name", "nome", "contact"]

    scores = []
    for col in df.columns:
        col_lower = str(col).lower()
        score = 0
        score += sum(4 for kw in kw_empresa if kw in col_lower)
        score += sum(1 for kw in kw_nome if kw in col_lower)

        amostra = df[col].dropna().head(20)
        if len(amostra):
            pct_str = amostra.apply(lambda v: isinstance(v, str)).sum() / len(amostra)
            score += int(pct_str > 0.5)
        scores.append((score, str(col)))
    scores.sort(reverse=True)
    return scores[0][1]


def nomes_do_excel(arquivo_excel, coluna: str | None) -> tuple[set[str], str, list[str]]:
    """Lê o Excel e devolve (conjunto de nomes normalizados, coluna usada, lista de colunas)."""
    df = pd.read_excel(arquivo_excel)
    colunas = [str(c) for c in df.columns]
    col = coluna if coluna and coluna in colunas else detectar_coluna_nome(df)
    nomes = {normalizar_nome(v) for v in df[col].dropna()}
    nomes = {n for n in nomes if n}
    return nomes, col, colunas


def nomes_do_zip(arquivo_zip) -> set[str]:
    """Extrai ZIP num temp dir e coleta stems de todos os arquivos."""
    tmp = tempfile.mkdtemp()
    try:
        with zipfile.ZipFile(arquivo_zip) as zf:
            zf.extractall(tmp)
        nomes = set()
        for path in Path(tmp).rglob("*"):
            if path.is_file():
                n = normalizar_nome(path.stem)
                if n:
                    nomes.add(n)
        return nomes
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def nomes_da_pasta(pasta: Path) -> set[str]:
    nomes = set()
    for path in pasta.rglob("*"):
        if path.is_file():
            n = normalizar_nome(path.stem)
            if n:
                nomes.add(n)
    return nomes


def comparar(nomes_excel: set[str], nomes_arquivos: set[str]) -> dict:
    so_excel = sorted(nomes_excel - nomes_arquivos)
    so_arquivos = sorted(nomes_arquivos - nomes_excel)
    em_ambos = sorted(nomes_excel & nomes_arquivos)
    return {
        "total_excel": len(nomes_excel),
        "total_arquivos": len(nomes_arquivos),
        "em_ambos": len(em_ambos),
        "so_excel": so_excel,
        "so_arquivos": so_arquivos,
        "divergencias": len(so_excel) + len(so_arquivos),
    }


def gerar_evidencia_md(resultado: dict, col_usada: str, nome_excel: str, nome_zip: str) -> str:
    agora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    so_excel = resultado["so_excel"]
    so_arquivos = resultado["so_arquivos"]
    divergencias = resultado["divergencias"]
    status = "✅ CONCLUÍDO (100%)" if divergencias == 0 else "⚠️ DIVERGÊNCIA ENCONTRADA"

    linhas = [
        f"# Relatório de Comparação de Agendas",
        f"*Gerado em: {agora}*\n",
        f"## Resumo Geral\n",
        f"- **Status:** {status}",
        f"- Arquivo Excel: `{nome_excel}` (coluna: `{col_usada}`)",
        f"- Arquivo ZIP: `{nome_zip}`",
        f"- Total no Excel: `{resultado['total_excel']}`",
        f"- Total nas Pastas: `{resultado['total_arquivos']}`",
        f"- Correspondências: `{resultado['em_ambos']}`",
    ]
    if divergencias:
        linhas.append(f"- Divergências: `{divergencias}`\n")
        linhas.append("## Detalhamento de Divergências\n")
        linhas.append("| Nome Identificado | Onde foi encontrado? |")
        linhas.append("| :--- | :--- |")
        for n in so_excel:
            linhas.append(f"| `{n}` | Somente no Excel |")
        for n in so_arquivos:
            linhas.append(f"| `{n}` | Somente nas Pastas (PDF) |")
    else:
        linhas.append("\nNenhuma divergência encontrada. Todos os arquivos conferem com o Excel.")

    linhas.append("\n---")
    linhas.append("*Este relatório serve como evidência de conferência entre os nomes listados na planilha e os arquivos gerados nas pastas.*")
    return "\n".join(linhas)


def gerar_bug_report_md(resultado: dict, col_usada: str, nome_excel: str, nome_zip: str) -> str:
    agora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    so_excel = resultado["so_excel"]
    so_arquivos = resultado["so_arquivos"]
    divergencias = resultado["divergencias"]

    linhas = [
        f"# Bug Report - Divergências de Agendas",
        f"*Data da Ocorrência: {agora}*\n",
        f"## Descrição",
        f"Foram detectadas inconsistências entre a planilha de controle e os arquivos gerados.\n",
        f"## Detalhes Técnicos",
        f"- **Total de Inconsistências:** `{divergencias}`",
        f"- **Arquivo Excel:** `{nome_excel}` (coluna: `{col_usada}`)",
        f"- **Arquivo ZIP:** `{nome_zip}`\n",
        f"## Itens para Revisão\n",
        f"| Nome no Sistema | Problema Detectado |",
        f"| :--- | :--- |",
    ]
    for n in so_excel:
        linhas.append(f"| `{n}` | Faltando arquivo PDF na pasta |")
    for n in so_arquivos:
        linhas.append(f"| `{n}` | Arquivo PDF existe mas não está no Excel |")

    linhas += [
        "\n## Sugestão de Ação",
        "1. Verificar se o nome no Excel possui caracteres especiais que impedem a criação correta do arquivo.",
        "2. Validar se o arquivo PDF correspondente foi gerado com um nome diferente do esperado.",
        "3. Caso o arquivo esteja faltando, realizar a regeneração da agenda específica.",
    ]
    return "\n".join(linhas)


# ── Rotas ─────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/colunas", methods=["POST"])
def colunas():
    """Recebe o Excel e devolve lista de colunas para o usuário escolher."""
    arq = request.files.get("excel")
    if not arq:
        return jsonify({"erro": "Arquivo Excel não enviado."}), 400
    try:
        df = pd.read_excel(arq)
        colunas = [str(c) for c in df.columns]
        sugerida = detectar_coluna_nome(df)
        return jsonify({"colunas": colunas, "sugerida": sugerida})
    except Exception as e:
        return jsonify({"erro": str(e)}), 500


@app.route("/comparar", methods=["POST"])
def comparar_route():
    excel_file = request.files.get("excel")
    outros_file = request.files.get("arquivos")  # ZIP
    coluna = request.form.get("coluna") or None

    if not excel_file or not outros_file:
        return jsonify({"erro": "Envie o arquivo Excel e o ZIP com os arquivos."}), 400

    nome_excel = excel_file.filename
    nome_zip = outros_file.filename

    try:
        nomes_ex, col_usada, _ = nomes_do_excel(excel_file, coluna)
        nomes_arq = nomes_do_zip(outros_file)
        resultado = comparar(nomes_ex, nomes_arq)
        resultado["coluna_usada"] = col_usada

        # Gera os markdowns e devolve no JSON para download no browser
        resultado["evidencia_md"] = gerar_evidencia_md(resultado, col_usada, nome_excel, nome_zip)
        if resultado["divergencias"] > 0:
            resultado["bug_report_md"] = gerar_bug_report_md(resultado, col_usada, nome_excel, nome_zip)

        return jsonify(resultado)
    except zipfile.BadZipFile:
        return jsonify({"erro": "O segundo arquivo precisa ser um ZIP válido."}), 400
    except Exception as e:
        return jsonify({"erro": str(e)}), 500


if __name__ == "__main__":
    print("Abrindo interface em http://localhost:5000")
    app.run(debug=True, port=5000)
