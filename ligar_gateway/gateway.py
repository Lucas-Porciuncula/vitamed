import pyautogui
import os
import time
import pygetwindow as gw

pyautogui.FAILSAFE = True  # move mouse pro canto pra parar

# =========================
# CONFIG
# =========================
APP_PATH = r"C:\Users\laporciuncula\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\On-premises data gateway (personal mode).lnk"

IMG_ENTRAR = r"C:\Users\laporciuncula\Documents\python\ligar_gateway\entrar.png"
IMG_EMAIL = r"C:\Users\laporciuncula\Documents\python\ligar_gateway\email.png"
IMG_AVANCAR = r"C:\Users\laporciuncula\Documents\python\ligar_gateway\avancar.png"
IMG_COMERCIAL = r"C:\Users\laporciuncula\Documents\python\ligar_gateway\comercial.png"
IMG_FECHAR = r"C:\Users\laporciuncula\Documents\python\ligar_gateway\fechar.png"

EMAIL = "comercialbi@vitalicenciamento.onmicrosoft.com"


# =========================
# FUNÇÕES
# =========================
def log(msg):
    print(f"[{time.strftime('%H:%M:%S')}] {msg}")


def esperar_imagem(path, confidence=0.8, timeout=20):
    """Espera imagem aparecer com timeout"""
    inicio = time.time()
    while time.time() - inicio < timeout:
        try:
            pos = pyautogui.locateCenterOnScreen(path, confidence=confidence)
            if pos:
                return pos
        except:
            pass
        time.sleep(0.5)
    return None


def clicar_imagem(path, confidence=0.8, timeout=20):
    pos = esperar_imagem(path, confidence, timeout)
    if pos:
        pyautogui.click(pos)
        return True
    return False


def mover_janela_tela1():
    """Força janela pra tela principal"""
    time.sleep(3)

    # método 1: atalho
    pyautogui.hotkey('win', 'shift', 'left')
    time.sleep(1)

    # método 2: posição absoluta
    try:
        janelas = gw.getWindowsWithTitle("On-premises")
        if janelas:
            janela = janelas[0]
            janela.activate()
            janela.moveTo(0, 0)
            janela.resizeTo(1200, 800)
    except:
        pass


# =========================
# EXECUÇÃO
# =========================
log("Abrindo aplicação...")
os.startfile(APP_PATH)

time.sleep(5)
mover_janela_tela1()

# =========================
# LOGIN
# =========================
log("Procurando botão entrar...")
if clicar_imagem(IMG_ENTRAR, confidence=0.9):
    log("Entrar encontrado")

    if clicar_imagem(IMG_EMAIL):
        log("Campo email encontrado")

        pyautogui.write(EMAIL, interval=0.05)

        if clicar_imagem(IMG_AVANCAR):
            log("Clicou em avançar")

            log("Procurando conta comercial...")
            if clicar_imagem(IMG_COMERCIAL, timeout=30):
                log("Login finalizado com sucesso ✅")
                time.sleep(7)
                if clicar_imagem(IMG_FECHAR, timeout=30):
                    log("Fechou janela de sucesso")
                else:
                    log("Botão fechar não encontrado ❌")
            else:
                log("Conta comercial NÃO encontrada ❌")

        else:
            log("Botão avançar não encontrado ❌")

    else:
        log("Campo email não encontrado ❌")

else:
    log("Botão entrar não encontrado ❌")