import subprocess

def close_libreoffice():
    try:
        subprocess.run(['pkill', 'soffice.bin'])
        print("LibreOffice foi fechado com sucesso.")
    except Exception as e:
        print("Ocorreu um erro ao fechar o LibreOffice:", e)

if __name__ == "__main__":
    close_libreoffice()
