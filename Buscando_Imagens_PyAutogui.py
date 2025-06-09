# Buscando itens em tela web quando não conseguimos acessar os elementos HTML

import pyautogui


    #-----------------------------------------------------#
                    ### MÁQUINA VIRTUAL ###
    #-----------------------------------------------------#

    # Item que deve ser buscado na tela printada
    menu_geral = r"C:\Users\Users\Users\telas_que_vou_buscar\menu_geral.PNG"

    # Tela atual printada
    pyautogui.screenshot("my_screenshot_nova.png")

    # Busca do item na tela printada com confiança de 68%
    buscar_menu_geral = pyautogui.locateOnScreen(menu_geral, confidence=0.677)

    if buscar_menu_geral:
        print(f"Elemento encontrado: {buscar_menu_geral}")
        centro_menu = pyautogui.center(buscar_menu_geral)
        print(f"Centro do elemento: {centro_menu}")
        pyautogui.click(centro_menu)
    else:
        print("Elemento menu_geral não encontrado na tela.")
    

    