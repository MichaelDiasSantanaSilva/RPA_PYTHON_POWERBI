import subprocess
import time
import pyautogui
from datetime import datetime
import os

# --- Configurações iniciais ---
# Caminho do executável do Power BI — substitua pelo seu caminho local
power_bi_exe = r"C:\Caminho\Para\Seu\PowerBI.exe"

# Lista dos arquivos PBIX que você quer atualizar — substitua pelos seus arquivos
arquivos_pbix = [
    r"C:\Caminho\Para\Relatorio1.pbix",
    r"C:\Caminho\Para\Relatorio2.pbix",
    r"C:\Caminho\Para\Relatorio3.pbix",
    # Adicione mais conforme precisar
]

# Imagens para reconhecimento na tela — coloque as suas imagens na pasta do script
imagem_icone_janela_powerbi = "janela_powerbi.png"
imagem_carregando = "carregando_modelo_dados.png"
imagem_atualizar = "atualizar.png"
imagem_janela_atualizacao = "janela_atualizacao.png"
imagem_trabalhando_nisso = "trabalhando_nisso.png"
imagem_publicar = "publicar.png"
imagem_admin_homolog = "homologacao.png"
imagem_botao_selecionar_cancelar = "selecionar.png"
imagem_botao_substituir = "substituir.png"
imagem_botao_entendi = "entendi.png"
imagem_erro_atualizacao = "erro_atualizacao.png"
imagem_botao_fechar = "botao_fechar.png"
imagem_botao_fechar_janela = "botao_fechar_janela.png"
imagem_erro_bloqueio_consultas = "erro_bloqueio_consultas.png"

# --- Funções de apoio ---
def localizar_imagem_seguro(imagem, confidence=0.7):
    try:
        return pyautogui.locateOnScreen(imagem, confidence=confidence)
    except Exception:
        return None

def localizar_centro_imagem_seguro(imagem, confidence=0.7):
    try:
        return pyautogui.locateCenterOnScreen(imagem, confidence=confidence)
    except Exception:
        return None

def esperar_imagem_sumir_com_delay_ate_sumir(imagem, confidence=0.7, delay_apos_sumir=5):
    print(f"⏳ Esperando a imagem '{imagem}' sumir da tela...")
    while True:
        if not localizar_imagem_seguro(imagem, confidence=confidence):
            print(f"✅ Imagem '{imagem}' sumiu.")
            time.sleep(delay_apos_sumir)
            return True
        time.sleep(8)

def encontrar_e_esperar_estavel(imagem, timeout_segundos=300, estabilidade_tempo=2, confidence=0.7):
    start_time = time.time()
    pos_anterior = None
    tempo_estavel = 0
    while time.time() - start_time < timeout_segundos:
        pos = localizar_centro_imagem_seguro(imagem, confidence=confidence)
        if pos:
            if pos == pos_anterior:
                tempo_estavel += 0.8
                if tempo_estavel >= estabilidade_tempo:
                    return pos
            else:
                pos_anterior = pos
                tempo_estavel = 0
        else:
            pos_anterior = None
            tempo_estavel = 0
        time.sleep(0.8)
    return None

def verificar_erro_atualizacao():
    return localizar_imagem_seguro(imagem_erro_atualizacao, confidence=0.9)

 """ Processo principal """
arquivos_publicados = []

for arquivo_pbix in arquivos_pbix:
    print(f"\n🚀 Processando: {arquivo_pbix}\n")

    subprocess.Popen([power_bi_exe, arquivo_pbix])
    time.sleep(10)

    location_atualizar = encontrar_e_esperar_estavel(imagem_atualizar, 600, 5)
    if not location_atualizar:
        print("❌ Botão 'Atualizar' não encontrado.")
        continue

    if localizar_imagem_seguro(imagem_carregando, confidence=0.7):
        esperar_imagem_sumir_com_delay_ate_sumir(imagem_carregando, confidence=0.7, delay_apos_sumir=3)

    pyautogui.moveTo(location_atualizar)
    pyautogui.click()
    time.sleep(5)

    erro_bloqueio = localizar_imagem_seguro(imagem_erro_bloqueio_consultas, confidence=0.85)
    if erro_bloqueio:
        print("⚠️ Erro de bloqueio detectado. Tentando novamente...")
        pos_fechar = localizar_centro_imagem_seguro(imagem_botao_fechar, confidence=0.9)
        if pos_fechar:
            pyautogui.moveTo(pos_fechar)
            pyautogui.click()
            time.sleep(3)
            pyautogui.click(location_atualizar)
            time.sleep(10)

    if verificar_erro_atualizacao():
        print("⚠️ Erro na atualização, tentando novamente...")
        pos = localizar_centro_imagem_seguro(imagem_botao_fechar, confidence=0.9)
        if pos:
            pyautogui.moveTo(pos)
            pyautogui.click()
            time.sleep(3)
            pyautogui.click(location_atualizar)
            time.sleep(10)

    print("🕵️ Aguardando janela de atualização aparecer...")
    start = time.time()
    while time.time() - start < 180:
        if localizar_imagem_seguro(imagem_janela_atualizacao, confidence=0.7):
            break
        time.sleep(2)
    else:
        print("❌ Timeout na janela de atualização.")
        continue

    esperar_imagem_sumir_com_delay_ate_sumir(imagem_janela_atualizacao, confidence=0.6, delay_apos_sumir=3)

    # Revalida erro de bloqueio até 3 tentativas
    for tentativa in range(3):
        erro_bloqueio_apos_atualizacao = localizar_imagem_seguro(imagem_erro_bloqueio_consultas, confidence=0.85)
        if not erro_bloqueio_apos_atualizacao:
            break
        print(f"⚠️ Erro de bloqueio persistente. Tentativa {tentativa+1}/3. Reiniciando atualização...")
        pos_fechar = localizar_centro_imagem_seguro(imagem_botao_fechar, confidence=0.9)
        if pos_fechar:
            pyautogui.moveTo(pos_fechar)
            pyautogui.click()
            time.sleep(3)
            pyautogui.click(location_atualizar)
            time.sleep(10)

    pyautogui.hotkey("ctrl", "s")
    time.sleep(3)
    esperar_imagem_sumir_com_delay_ate_sumir(imagem_trabalhando_nisso, confidence=0.7, delay_apos_sumir=3)

    location = encontrar_e_esperar_estavel(imagem_publicar, 120, 1)
    if not location:
        continue
    pyautogui.moveTo(location)
    pyautogui.click()
    time.sleep(3)

    location = encontrar_e_esperar_estavel(imagem_admin_homolog, 120, 1)
    if not location:
        continue
    pyautogui.moveTo(location)
    pyautogui.click()

    box = localizar_imagem_seguro(imagem_botao_selecionar_cancelar, confidence=0.9)
    if not box:
        continue

    pyautogui.moveTo(box.left + 25, box.top + box.height // 2)
    pyautogui.click()
    time.sleep(2)

    pos = encontrar_e_esperar_estavel(imagem_botao_substituir, 60)
    if pos:
        pyautogui.moveTo(pos)
        pyautogui.click()
        time.sleep(2)

    pos = encontrar_e_esperar_estavel(imagem_botao_entendi, 120, 1)
    if pos:
        pyautogui.moveTo(pos)
        pyautogui.click()
        time.sleep(2)

    pos = encontrar_e_esperar_estavel(imagem_botao_fechar_janela, 60, confidence=0.9)
    if pos:
        pyautogui.moveTo(pos)
        pyautogui.click()
        time.sleep(8)
    else:
        print("⚠️ Botão fechar não encontrado.")

    arquivos_publicados.append(arquivo_pbix)

""" Finalização """
print("\n🌟 Todos os arquivos foram processados.")
print("\n📦 Relatórios publicados com sucesso:")

with open("log.txt", "a", encoding="utf-8") as log_file:
    for arquivo in arquivos_publicados:
        nome = os.path.basename(arquivo)
        data_hora = datetime.now().strftime("%d/%m/%Y %H:%M")
        linha = f"{nome} - {data_hora}"
        print(f"✅ {linha}")
        log_file.write(f"{linha}\n")
