from scripts.database import connect_to_db, query_to_dataframe, close_connection
from scripts.email_sender import send_emails
from config.email_config import EMAIL_LIST
def main():
    # Estabelecer Conexão
    conn = connect_to_db()
    try:
        # Executar Consulta e Processar Dados
        df = query_to_dataframe(conn)
        
        # Definir caminho de destino
        output_path = r'C:\Users\paulocardoso\OneDrive - ARCO Educacao\Área de Trabalho\Docs_Paulo\Paulo\Pessoal\Python\Projects\data\FaturadosDia.xlsx'
        
        # Salvar o DataFrame em um arquivo Excel
        df.to_excel(output_path, index=False)
        
        # Enviar email com o arquivo
        email_list = EMAIL_LIST
        send_emails(email_list, output_path)

    except Exception as e:
        print(f"Erro ao executar a consulta ou salvar o arquivo: {e}")
    
    finally:
        # Fechar Conexão
        close_connection(conn)

if __name__ == "__main__":
    main()
