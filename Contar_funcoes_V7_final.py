import pandas as pd
from collections import Counter
import re  # Biblioteca para expressões regulares

def contar_funcoes(script):
    """
    Conta a ocorrência de funções em um script de acordo com uma lista pré-definida,
    ignorando diferenças entre maiúsculas e minúsculas e considerando-as apenas como palavras isoladas.
    """
    # Lista de funções a serem buscadas
    funcoes = [
    "Abs", "AddDays","As", "AddMonths", "AddWeeks", "AddYears", "Age", "Alt", "ApplyMap", "AutoNumber", 
    "AutonumberHash128", "AutonumberHash256", "Avg", "Ceil", "Chr", "Class", "Coalesce", "Concat", 
    "Count", "Date", "Date#", "Day", "DayName", "DayOfWeek", "DayOfYear", "Dual", "Evaluate", "Exists", 
    "Exp", "FieldCount", "FieldName", "FieldNumber", "FieldValue", "FileBaseName", "FileDir", "FileName", 
    "FilePath", "FileSize", "FileTime", "FirstSortedValue", "Floor", "Frac", "Hash128", "Hash160", "Hash256", 
    "Hour", "If", "InDay", "InMonth", "InWeek", "InYear", "Index", "Integer", "Interval", "Interval#", 
    "IntervalMatch", "IsNull", "IsNum", "IsText", "KeepChar", "Len", "Left", "Log", "Log10", "Lookup", 
    "Lower", "LTrim", "MakeDate", "MakeTime", "MakeTimestamp", "MapSubstring", "Match", "Max", "Mid", 
    "Min", "Minute", "Mod", "Month", "MonthEnd", "MonthName", "MonthStart", "Null", "Num", "Only", 
    "OSUser", "Peek", "Pick", "Pi", "Pow", "Previous", "PurgeChar", "QvdCreateTime", "QvdNoOfRecords", 
    "QvdTableName", "Rand", "Replace", "Right", "Round", "RTrim", "ScriptError", "ScriptErrorCount", 
    "Second", "Sign", "Sqrt", "SubField", "SubStringCount", "Sum", "Text", "Time", "Time#", "Timestamp", 
    "Timestamp#", "Trim", "TypeOf", "Upper", "Week", "WeekDay", "WeekEnd", "WeekName", "WeekStart", 
    "WildMatch", "Year", "YearEnd", "YearName", "YearStart", "Then", "Else", "For", "End If", "Loop", 
    "While", "Until", "Exit Script", "Sub", "Call SubName", "Load", "Where", "Group By", "Order By", 
    "Distinct", "Join", "Inner Join", "Left Join", "Right Join", "Outer Join", "Concatenate", 
    "NoConcatenate", "Keep", "Mapping Load", "ApplyMap", "Binary", "Directory", "Set", "Drop Table", 
    "Drop Field", "Rename Field", "Rename Fields", "Qualify", "Unqualify", "Section Access", "Section Application"
    ]







    
    # Normalizar para minúsculas e inicializar contador
    funcoes_lower = [func.lower() for func in funcoes]
    script_lower = script.lower()
    contador = Counter({func: 0 for func in funcoes_lower})
    
    # Criar uma expressão regular para cada função
    for func in funcoes_lower:
        regex = rf"\b{func}\b"  # Verifica a função como palavra isolada
        contador[func] += len(re.findall(regex, script_lower))
    
    return contador

def gerar_excel_e_html(script, file_name="funcoes.xlsx", html_name="funcoes.html"):
    """
    Gera arquivos Excel e HTML com o resumo da contagem das funções usadas no script.
    """
    contador = contar_funcoes(script)
    
    # Criar DataFrame com as contagens
    df_geral = pd.DataFrame(list(contador.items()), columns=["Função", "Contagem"])
    
    # Filtrar funções com contagem > 0
    resumo_df = df_geral[df_geral["Contagem"] > 0].copy()
    total_funcoes_utilizadas = len(resumo_df)
    quantidade_geral_de_uso = resumo_df["Contagem"].sum()
    total_funcoes_analisadas = len(df_geral)
    funcoes_utilizadas = ", ".join(resumo_df["Função"].tolist())
    
    # Adicionar linha total no resumo
    resumo_df.loc["Total"] = ["Total Geral", quantidade_geral_de_uso]
    
    # Salvar a planilha Excel
    with pd.ExcelWriter(file_name, engine="xlsxwriter") as writer:
        df_geral.to_excel(writer, sheet_name="Funções Gerais", index=False)
        resumo_df.to_excel(writer, sheet_name="Resumo", index=False)
    
    # Criar o HTML
    html_content = f"""
    <html>
    <head>
        <title>Guimetria - Contagem de Funções</title>
        <style>
            body {{
                font-family: Arial, sans-serif;
                margin: 20px;
            }}
            h1 {{
                color: #333;
            }}
            table {{
                width: 80%;
                border-collapse: collapse;
                margin: 20px 0;
            }}
            th, td {{
                border: 1px solid #ccc;
                padding: 8px;
                text-align: left;
            }}
            th {{
                background-color: #f4f4f4;
            }}
            tr:nth-child(even) {{
                background-color: #f9f9f9;
            }}
            tr:hover {{
                background-color: #f1f1f1;
            }}
        </style>
    </head>
    <body>
		<div class="header">
        <img src="/Análise_de_Funções/falqon.png" alt="Logo" class="logo"> <!-- Insira a logo aqui -->
        </div>
        <h1>Resumo - Sistema de Contagem de Funções</h1>
        <div>
            <p><strong>Total de funções analisadas:</strong> {total_funcoes_analisadas}</p>
            <p><strong>Total de funções utilizadas:</strong> {total_funcoes_utilizadas}</p>
            <p><strong>Quantidade geral de uso:</strong> {quantidade_geral_de_uso}</p>
            <p><strong>Funções utilizadas:</strong> {funcoes_utilizadas}</p>
        </div>
        <h1>Detalhes das Funções</h1>
        {df_geral.to_html(index=False)}
        <h2>Resumo das Funções Utilizadas</h2>
        {resumo_df.to_html(index=False)}
    </body>
    </html>
    """
    
    # Salvar o HTML
    with open(html_name, "w", encoding="utf-8") as f:
        f.write(html_content)

# Exemplo de uso
script_exemplo = """

 

"""
gerar_excel_e_html(script_exemplo)