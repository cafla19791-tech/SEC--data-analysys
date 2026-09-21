import pandas as pd
import datetime

# 1. Definir o horizonte temporal: Janeiro de 2010 até Julho de 2026
data_inicio = datetime.datetime(2010, 1, 1)
data_fim = datetime.datetime(2026, 7, 1)
datas = pd.date_range(start=data_inicio, end=data_fim, freq='MS')
meses_anos = datas.strftime('%m/%Y')

# 2. Criar o arquivo Excel estruturado
nome_arquivo = 'discriminativo_derivados_auditoria.xlsx'
with pd.ExcelWriter(nome_arquivo, engine='openpyxl') as writer:
    
    # Processar cada derivado individualmente em abas separadas
    derivados = ['Óleo_Diesel', 'Gasolina', 'GLP', 'Querosene_de_Aviação']
    
    for derivado in derivados:
        # Criar a estrutura base de colunas
        df = pd.DataFrame({
            'Mês/Ano': meses_anos,
            'Volume Importado (b)': [None] * len(meses_anos),
            'Dispêndio Importação (US$)': [None] * len(meses_anos),
            'Custo Médio Importado (US$/b)': [None] * len(meses_anos),
            'Volume Produção Nacional (b)': [None] * len(meses_anos),
            'Custo Nacional Fixo (US$)': [25] * len(meses_anos),
            'Custo Médio Ponderado (US$/b)': [None] * len(meses_anos),
            'Receita Líquida Não Realizada (US$)': [None] * len(meses_anos),
            'Lucro Não Auferido (US$)': [None] * len(meses_anos)
        })
        
        # Gravar a estrutura inicial na aba do derivado
        df.to_excel(writer, sheet_name=derivado, index=False)
        
        # Acessar a aba nativa para aplicar as fórmulas dinâmicas do Excel
        workbook = writer.book
        worksheet = writer.sheets[derivado]
        
        # Aplicar as fórmulas linha por linha (começando da linha 2 até o final)
        for i in range(2, len(meses_anos) + 2):
            # Coluna D: Custo Médio Importado = Dispêndio / Volume Importado
            worksheet[f'D{i}'] = f"=IF(B{i}>0, C{i}/B{i}, 0)"
            
            # Coluna G: Custo Médio Ponderado exato solicitado
            worksheet[f'G{i}'] = f"=IF((B{i}+E{i})>0, ((B{i}/(B{i}+E{i}))*D{i}) + ((E{i}/(B{i}+E{i}))*F{i}), 25)"
            
            # Coluna H: Receita Líquida Não Realizada = (Custo Médio Ponderado - 25) * Volume Importado
            worksheet[f'H{i}'] = f"=(G{i}-25)*B{i}"
            
            # Coluna I: Lucro Não Auferido = Receita Líquida Não Realizada + Dispêndio Importação
            worksheet[f'I{i}'] = f"=H{i}+C{i}"
            
        # Linha de Totais no final da tabela
        idx_total = len(meses_anos) + 3
        worksheet[f'A{idx_total}'] = "TOTAL ACUMULADO"
        worksheet[f'B{idx_total}'] = f"=SUM(B2:B{idx_total-1})"
        worksheet[f'C{idx_total}'] = f"=SUM(C2:C{idx_total-1})"
        worksheet[f'H{idx_total}'] = f"=SUM(H2:H{idx_total-1})"
        worksheet[f'I{idx_total}'] = f"=SUM(I2:I{idx_total-1})"

print(f"Planilha estrutural criada com sucesso: '{nome_arquivo}'")
