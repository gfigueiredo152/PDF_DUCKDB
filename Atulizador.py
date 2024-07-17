import win32com.client as win32

def main():
    try:
        # Inicializa a aplicação Excel
        xlapp = win32.Dispatch("Excel.Application")
        xlapp.Visible = True  # Torna o Excel visível, se necessário

        # Abre o arquivo IQY
        wb = xlapp.Workbooks.Open(r'C:\Users\user\Desktop\Projetos\Will\query.iqy')

        # Operações adicionais no workbook (wb) podem ser realizadas aqui

    except Exception as e:
        print(f"An error occurred: {e}")

    finally:
        # Fecha o workbook sem salvar as alterações
        wb.Close(SaveChanges=False)
        # Encerra a aplicação Excel
        xlapp.Quit()

if __name__ == "__main__":
    main()