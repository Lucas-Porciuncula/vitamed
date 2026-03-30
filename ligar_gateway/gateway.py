import pyautogui
import os
from time import sleep

os.startfile(
    r"C:\Users\laporciuncula\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\On-premises data gateway (personal mode).lnk"
)
pasta = r"C:\Users\laporciuncula\Downloads\Ligar Gateway"
sleep(5)
while True:
    try:
        imagem = pyautogui.locateCenterOnScreen( r"C:\Users\laporciuncula\Documents\python\ligar_gateway\entrar.png", confidence=0.9)
    except:
        imagem = None
    if imagem:
        sleep(0.4)
        pyautogui.click(imagem)
        sleep(0.4)

        imagem_email = pyautogui.locateCenterOnScreen( r"C:\Users\laporciuncula\Documents\python\ligar_gateway\email.png", confidence=0.8)
        if imagem_email:
            sleep(0.4)
            pyautogui.click(imagem_email)
            sleep(0.4)

            pyautogui.write("comercialbi@vitalicenciamento.onmicrosoft.com")
            sleep(1)

            imagem_avancar = pyautogui.locateCenterOnScreen( r"C:\Users\laporciuncula\Documents\python\ligar_gateway\avancar.png", confidence=0.8)
            if imagem_avancar:
                sleep(0.4)
                pyautogui.click(imagem_avancar)
                sleep(0.4)

                while True:
                    try:
                        imagem_comercial = pyautogui.locateOnScreen(
                             r"C:\Users\laporciuncula\Documents\python\ligar_gateway\comercial.png", confidence=0.8
                        )
                    except:
                        imagem_comercial = None

                    if imagem_comercial:
                        sleep(0.4)
                        pyautogui.click(imagem_comercial)
                        break
