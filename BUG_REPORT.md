# 🐛 Bug Report - Divergências de Agendas
*Data da Ocorrência: 16/03/2026 17:45:28*

## 📝 Descrição
Foram detectadas inconsistências entre as planilhas de controle (Excel) e os arquivos gerados (Pastas/PDFs).

## 🔍 Detalhes Técnicos
- **Total de Inconsistências:** `2`
- **Ambiente:** Produção (Duco Italy 2026)

## 📋 Itens para Revisão
| Tipo | Nome no Sistema | Problema Detectado |
| :--- | :--- | :--- |
| Buyer | `fioraé travel` | Faltando arquivo PDF na pasta |
| Buyer | `fiora travel` | Arquivo PDF existe mas não está no Excel |

## ✅ Sugestão de Ação
1. Verificar se o nome no Excel possui caracteres especiais que impedem a criação correta do arquivo.
2. Validar se o arquivo PDF correspondente foi gerado com um nome diferente do esperado.
3. Caso o arquivo esteja faltando, realizar a regeneração da agenda específica.
