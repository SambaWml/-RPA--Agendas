# Comparador de Agendas

Ferramenta criada para conferir se todos os nomes listados na planilha de controle realmente têm um arquivo de agenda gerado na pasta — e vice-versa.

Surgiu de uma necessidade real: depois de gerar centenas de PDFs de agendas pro evento, precisávamos ter certeza de que nenhum havia ficado pra trás (ou sido gerado com nome errado). A comparação é feita pelo nome da empresa, normalizando acentos e caracteres especiais pra evitar falsos positivos.

---

## Como usar

### Interface web (recomendado)

```bash
pip install flask pandas openpyxl
py app.py
```

Acesse **http://localhost:5000**, faça upload da planilha Excel e do ZIP com os PDFs. O sistema detecta automaticamente qual coluna usar, mas você pode trocar se quiser.

No final aparece o resultado e botões pra baixar o **Relatório de Evidência** e o **Bug Report** (quando há divergências).

### Script direto (linha de comando)

```bash
pip install pandas openpyxl
py comparar_nomes.py
```

Espera encontrar na mesma pasta:
- Um `.xlsx` com `buyers` no nome
- Um `.xlsx` com `exhibitors` no nome
- Uma pasta extraída com `buyers` no nome
- Uma pasta extraída com `exhibitors` no nome

Gera o `relatorio_comparacao_nomes.csv` e chama o `gerar_evidencia.py` automaticamente, que cria o `RELATORIO_EVIDENCIA.md` e o `BUG_REPORT.md` se houver divergências.

---

## Arquivos do projeto

| Arquivo | O que faz |
| --- | --- |
| `app.py` | Backend da interface web (Flask) |
| `comparar_nomes.py` | Script de linha de comando |
| `gerar_evidencia.py` | Gera os relatórios `.md` a partir do CSV |
| `templates/index.html` | Interface web |

---

## Como a comparação funciona

Os nomes são normalizados antes de comparar: tudo vira minúsculo, acentos e caracteres especiais viram espaço, espaços duplos são removidos. Assim `Fioraé Travel` e `fiorae travel` são considerados o mesmo nome — o que evita divergências falsas por causa de formatação.

A coluna de comparação é detectada automaticamente pela planilha:
- Buyers → `Agency - Company`
- Exhibitors → `Company Name`

---

## Dependências

```
flask
pandas
openpyxl
```
