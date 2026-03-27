import os
from playwright.sync_api import sync_playwright

# Caminhos baseados na sua estrutura de Analista na Diálogo
FILE_FILA = r"C:\Notas_Fiscais_SP\fila_processamento.txt"
SAVE_PATH = r"C:\Notas_Fiscais_SP\PDFs_Baixados"

def baixar_notas():
    if not os.path.exists(FILE_FILA):
        print("Fila vazia ou não encontrada. Verifique o Outlook.")
        return
    
    if not os.path.exists(SAVE_PATH):
        os.makedirs(SAVE_PATH)
    
    with open(FILE_FILA, "r", encoding="utf-8") as f:
        linhas = f.readlines()

    if not linhas:
        print("Nenhum link para processar.")
        return

    with sync_playwright() as p:
        # headless=False permite que você veja o robô trabalhando nas 7 notas
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()
        
        for linha in linhas:
            if "|" in linha:
                numero, link = linha.strip().split("|")
                print(f"Baixando NF {numero}...")
                
                try:
                    page = context.new_page()
                    page.goto(link, timeout=60000)
                    
                    with page.expect_download() as download_info:
                        page.get_by_text("Download NFS-e").click()
                    
                    download = download_info.value
                    download.save_as(os.path.join(SAVE_PATH, f"{numero}.pdf"))
                    page.close()
                    print(f"Sucesso: {numero}.pdf salvo.")
                except Exception as e:
                    print(f"Erro ao baixar a nota {numero}: {e}")
            
        browser.close()
    
    # Após processar as 2.000 notas ou as 7 de teste, deletamos a fila
    os.remove(FILE_FILA)

if __name__ == "__main__":
    baixar_notas()