import pandas as pd
from datetime import datetime

def gerar_relatorio():
    csv_path = 'relatorio_comparacao_nomes.csv'
    try:
        df = pd.read_csv(csv_path)
    except Exception as e:
        print(f"Erro ao ler o CSV: {e}")
        return

    data_atual = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    
    with open('RELATORIO_EVIDENCIA.md', 'w', encoding='utf-8') as f:
        f.write(f"# 📊 Relatório de Comparação de Agendas - Duco Italy 2026\n")
        f.write(f"*Gerado em: {data_atual}*\n\n")
        
        # --- RESUMO GERAL ---
        f.write("## 📌 Resumo Geral\n\n")
        
        for tipo in ['buyer', 'exhibitor']:
            df_tipo = df[df['tipo'] == tipo]
            total_excel = len(df_tipo[df_tipo['origem'].isin(['excel', 'ambos'])])
            total_pastas = len(df_tipo[df_tipo['origem'].isin(['pastas', 'ambos'])])
            em_ambos = len(df_tipo[df_tipo['origem'] == 'ambos'])
            so_excel = len(df_tipo[df_tipo['origem'] == 'excel'])
            so_pastas = len(df_tipo[df_tipo['origem'] == 'pastas'])
            
            status = "✅ CONCLUÍDO (100%)" if so_excel == 0 and so_pastas == 0 else "⚠️ DIVERGÊNCIA ENCONTRADA"
            
            f.write(f"### 🔹 {tipo.upper()}S\n")
            f.write(f"- **Status:** {status}\n")
            f.write(f"- Total no Excel: `{total_excel}`\n")
            f.write(f"- Total nas Pastas: `{total_pastas}`\n")
            f.write(f"- Correspondências: `{em_ambos}`\n")
            
            if so_excel > 0 or so_pastas > 0:
                f.write(f"- Divergências: `{so_excel + so_pastas}`\n")
            f.write("\n")

        # --- DETALHAMENTO DE DIVERGÊNCIAS ---
        f.write("## ⚠️ Detalhamento de Divergências\n\n")
        divergencias = df[df['origem'] != 'ambos']
        
        if divergencias.empty:
            f.write("Nenhuma divergência encontrada. Todos os arquivos conferem com o Excel.\n")
        else:
            f.write("| Tipo | Nome Identificado | Onde foi encontrado? |\n")
            f.write("| :--- | :--- | :--- |\n")
            for _, row in divergencias.iterrows():
                origem_formatada = "Somente no Excel" if row['origem'] == 'excel' else "Somente nas Pastas (PDF)"
                f.write(f"| {row['tipo'].capitalize()} | `{row['nome']}` | {origem_formatada} |\n")
        
        f.write("\n---\n")
        f.write("*Este relatório serve como evidência de conferência entre os nomes listados nas planilhas de controle e os arquivos PDFs gerados nas respectivas pastas.*")

    print("Relatório 'RELATORIO_EVIDENCIA.md' gerado com sucesso!")

    # --- GERAÇÃO DE BUG REPORT ---
    if not divergencias.empty:
        with open('BUG_REPORT.md', 'w', encoding='utf-8') as f_bug:
            f_bug.write(f"# 🐛 Bug Report - Divergências de Agendas\n")
            f_bug.write(f"*Data da Ocorrência: {data_atual}*\n\n")
            
            f_bug.write("## 📝 Descrição\n")
            f_bug.write("Foram detectadas inconsistências entre as planilhas de controle (Excel) e os arquivos gerados (Pastas/PDFs).\n\n")
            
            f_bug.write("## 🔍 Detalhes Técnicos\n")
            f_bug.write("- **Total de Inconsistências:** `{}`\n".format(len(divergencias)))
            f_bug.write("- **Ambiente:** Produção (Duco Italy 2026)\n\n")
            
            f_bug.write("## 📋 Itens para Revisão\n")
            f_bug.write("| Tipo | Nome no Sistema | Problema Detectado |\n")
            f_bug.write("| :--- | :--- | :--- |\n")
            
            for _, row in divergencias.iterrows():
                if row['origem'] == 'excel':
                    problema = "Faltando arquivo PDF na pasta"
                else:
                    problema = "Arquivo PDF existe mas não está no Excel"
                f_bug.write(f"| {row['tipo'].capitalize()} | `{row['nome']}` | {problema} |\n")
            
            f_bug.write("\n## ✅ Sugestão de Ação\n")
            f_bug.write("1. Verificar se o nome no Excel possui caracteres especiais que impedem a criação correta do arquivo.\n")
            f_bug.write("2. Validar se o arquivo PDF correspondente foi gerado com um nome diferente do esperado.\n")
            f_bug.write("3. Caso o arquivo esteja faltando, realizar a regeneração da agenda específica.\n")
            
        print("Bug Report 'BUG_REPORT.md' gerado com sucesso!")
    else:
        # Se não houver divergências, deleta o bug report antigo se existir
        import os
        if os.path.exists('BUG_REPORT.md'):
            os.remove('BUG_REPORT.md')
            print("Nenhuma divergência encontrada. Bug report anterior removido.")

if __name__ == "__main__":
    gerar_relatorio()
