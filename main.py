import pandas as pd
import duckdb as db
import xlwings as xw
import tempfile
import os
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
import Atulizador
def inserir_dados_duckdb(df, conn):
    try:
        # Registrar DataFrame como uma tabela temporária
        conn.register('df_tratado', df)
        
        # Criar tabela permanente e inserir dados tratados nela
        conn.execute('CREATE TABLE dados_tratados AS SELECT * FROM df_tratado')
        
    except db.DuckDBError as e:
        print(f"Erro ao inserir dados no DuckDB: {e}")
        raise
    except Exception as e:
        print(f"Erro inesperado: {e}")
        raise

def tratar_dados(df):
    # Renomear colunas, remover espaços em branco e tratar dados ausentes
    novos_nomes_colunas = {
        'Título': 'Title',
        'Resíduo Sólido Urbano': 'RSU',
        'RSU Entrada Unidade': 'RSU_Ent_Unit',
        'Resíduo Sólido Urbano Tratado': 'RSU_Tratado',
        'RSU Resultado Unidade': 'RSU_Result_Unit',
        'CBSI para retroalimentação': 'CBSI_retroalimentacao',
        'CBSI Unidade': 'CBSI_Unit',
        'CBSI Final': 'CBSI_Final',
        'Nome responsavel': 'Nome_responsavel',
        'Assinatura': 'Assinatura',
        'Descrição': 'Descricao',
        'Foto': 'Foto',
        'Audio': 'Audio',
        'Data Base': 'Data_Base',
        'DataAtual': 'Data_Atual',
        'Dias': 'Dias',
        'dias_rel': 'dias_rel',
        'dia_semana': 'dia_semana',
        'diasok': 'diaSok',
        'Tipo de Item': 'Tipo_Item',
        'Caminho': 'Caminho'
    }
    df = df.rename(columns=novos_nomes_colunas)
    
    df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
    
    df = df.fillna({
        'Title': '01/01/1999',
        'RSU': 0.0,
        'RSU_Ent_Unit': 'null',
        'RSU_Tratado': 0.0,
        'RSU_Result_Unit': 'null',
        'CBSI_retroalimentacao': 0.0,
        'CBSI_Unit': 'null',
        'CBSI_Final': 0.0,
        'Nome_responsavel': 'null',
        'Assinatura': 'null',
        'Descricao': 'null',
        'Foto': 'null',
        'Audio': 'null',
        'Data_Base': '01/01/1999',
        'Data_Atual': '01/01/1999',
        'Dias': 0.0,
        'dias_rel': 'null',
        'dia_semana': 'null',
        'diaSok': 'null',
        'Tipo_Item': 'null',
        'Caminho': 'null'
    })
    
   
# Supondo que df seja seu DataFrame
    
    #df['Title'] = pd.to_datetime(df['Title'], errors='coerce')
    df['Data_Base'] = pd.to_datetime(df['Data_Base'], errors='coerce')
    df['Data_Atual'] = pd.to_datetime(df['Data_Atual'], errors='coerce')

    

    df['Title'] = pd.to_datetime(df['Title'], format='%d/%m/%Y').dt.strftime('%Y-%m-%d')
    df['Data_Base'] = df['Data_Base'].dt.strftime('%Y-%m-%d %H:%M:%S')
    df['Data_Atual'] = df['Data_Atual'].dt.strftime('%Y-%m-%d %H:%M:%S')
    
    print ('\n',df,'\n')

    df['RSU'] = df['RSU'].apply(lambda x: str(x).replace(',', '.') if isinstance(x, str) else x)
    df['RSU_Tratado'] = df['RSU_Tratado'].apply(lambda x: str(x).replace(',', '.') if isinstance(x, str) else x)
    df['CBSI_retroalimentacao'] = df['CBSI_retroalimentacao'].apply(lambda x: str(x).replace(',', '.') if isinstance(x, str) else x)
    df['CBSI_Final'] = df['CBSI_Final'].apply(lambda x: str(x).replace(',', '.') if isinstance(x, str) else x)
    df['Dias'] = df['Dias'].apply(lambda x: str(x).replace(',', '.') if isinstance(x, str) else x)
    
    return df

def data_import():
    # Caminho para o arquivo .iqy
    iqy_file_path = r'C:\Users\user\Desktop\Projetos\Will\query.iqy'

    # Criar um arquivo Excel temporário
    temp_excel_path = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False).name

    try:
        # Abrir o arquivo .iqy no Excel usando xlwings
        wb = xw.Book(iqy_file_path)

        # Ativar a conexão e atualizar os dados
        wb.api.RefreshAll()

        # Salvar como um arquivo Excel temporário
        wb.save(temp_excel_path)
        wb.close()

        # Ler os dados do arquivo Excel temporário para um DataFrame pandas
        df = pd.read_excel(temp_excel_path)

        return df

    except Exception as e:
        print(f"Erro ao importar dados do arquivo .iqy: {e}")
        return None

    finally:
        # Fechar o arquivo Excel e remover o arquivo temporário
        try:
            os.remove(temp_excel_path)
        except Exception as e:
            print(f"Erro ao remover arquivo temporário: {e}")

def salvar_resultado_em_pdf(result, file_path, logo_path=None, marca_path=None):
    
    c = canvas.Canvas(file_path, pagesize=letter)
    width, height = letter

    #Entrada das imagens do logo  

    if logo_path:
        logo = ImageReader(logo_path)
        c.drawImage(logo, 80, height -120, width=450, preserveAspectRatio=True, mask='auto')
    
    if marca_path:
        marca = ImageReader(marca_path)
        c.drawImage(marca, 400, height -950, width=180, preserveAspectRatio=True, mask='auto')

       # Adicionar marca d'água
    watermark_path = r'C:\Users\user\Desktop\Projetos\Will\Fundo_reciclagem.png'
    if os.path.exists(watermark_path):
        watermark = ImageReader(watermark_path)
        c.drawImage(watermark, 100, 100, width=400, height=300, preserveAspectRatio=True, mask='auto')
    

    c.setFont("Helvetica-Bold", 12)
    c.drawString(230, height - 150, "Relatório Diário de Insumos .")

    c.setFont("Helvetica", 16)
    y = height - 350
    for column, value in result.items():
        if column == "Resíduo Sólido Urbano":
            c.drawString(30, y, f"Resíduo Sólido Urbano:                                    {int(value)} quilogramas")
        elif column == "Resíduo Sólido Urbano Tratado":
            y -= 20
            c.drawString(30, y, f"Resíduo Sólido Urbano Tratado:                      {int(value)} quilogramas")
        elif column == "CBSI para retroalimentação":
            y -= 20
            c.drawString(30, y, f"CBSI para retroalimentação:                             {int(value)} quilogramas")
        elif column == "CBSI final":
            y -= 20
            c.drawString(30, y, f"CBSI final:                                                          {int(value)} quilogramas")
        y -= 20
        if y < 40:  # Se a página estiver cheia, criar uma nova
            c.showPage()
            y = height - 40

    c.save()

def main():
    # atulizador
    Atulizador.main()
    dados = data_import()
    print ('\n',dados,'\n')
    if dados is not None:
        df_tratado = tratar_dados(dados)
        
        # Conectar ao DuckDB
        conn = db.connect(database=':memory:', read_only=False)
        
        try:
            # Inserir dados tratados no DuckDB
            inserir_dados_duckdb(df_tratado, conn)
            
            # Exemplo de SELECT no DuckDB
            result = conn.execute("""
                SELECT 
                    CAST(RSU AS INTEGER) * 1000 AS "Resíduo Sólido Urbano",
                    CAST(RSU_Tratado AS INTEGER) * 1000 AS "Resíduo Sólido Urbano Tratado",
                    CAST(CBSI_retroalimentacao AS INTEGER) * 1000 AS "CBSI para retroalimentação",
                    CAST(CBSI_Final AS INTEGER) * 1000 AS "CBSI final"
                FROM dados_tratados
                WHERE Title = (SELECT MAX(Title) FROM dados_tratados);
            """).fetchdf().iloc[0]
            
            print("Dados selecionados do DuckDB:")
            print(result)
            
            # Salvar resultado do SELECT em PDF
            pdf_file_path = "resultado_select.pdf"
            logo_path = r'C:\Users\user\Desktop\Projetos\Will\logo_01.PNG'  
            marca_path =  r'C:\Users\user\Desktop\Projetos\Will\Ass_Construpro.jpg'
            salvar_resultado_em_pdf(result, pdf_file_path, logo_path, marca_path)
            print(f"Resultado do SELECT salvo em {pdf_file_path}")
        
        finally:
            # Fechar conexão com DuckDB
            conn.close()

if __name__ == "__main__":
    main()
