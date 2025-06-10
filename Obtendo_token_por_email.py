# código para automação web que depende de inserção de token recebido por e-mail para logar no site.

from selenium import webdriver
import win32com.client as win32

# Função para obter o token de autenticação
def obter_token(sender, subject):
    try:
        outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6)
        deleted_items = outlook.GetDefaultFolder(3)

        while True:
            messages = inbox.Items.Restrict(f"[SenderEmailAddress]='{sender}' AND [Subject]='{subject}'")
            if messages.Count > 0:
                messages.Sort("[ReceivedTime]", True)
                latest_message = messages[0]
                token = latest_message.Body.strip()[-6:]
                latest_message.Move(deleted_items)
                return token
            else:
                time.sleep(2)
    except Exception as e:
        print(f"Erro ao acessar o Outlook: {e}")
        return None


# Esperar a página de autenticação e inserir o token
token = obter_token(sender_email, email_subject)
   if token:
       wait.until(EC.presence_of_element_located((By.XPATH, {'html_element'}))).send_keys(token)
       driver.find_element(By.XPATH, {'html_element'}).click()
   else:
       print("Token de autenticação não encontrado.")

    