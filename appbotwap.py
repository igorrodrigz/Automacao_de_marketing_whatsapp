import time
import pywhatkit
import openpyxl
import emoji

# Variaveis

smiley = emoji.emojize(':smiling_face_with_smiling_eyes:')
handy = emoji.emojize(':hand_with_fingers_splayed:')
sparky = emoji.emojize(':sparkles:')
chaty = emoji.emojize(':thought_balloon:')

mensagem_saudacao = emoji.emojize(f"""Oi, tudo bom? {smiley} \nAqui é o Igor da MR Solutions, *especializados em melhorar o atendimento* ao cliente em clínicas médicas. Podemos ajudar sua equipe de recepcionistas a proporcionar uma experiência mais eficiente e Acolhedora para os pacientes. 
\nPodemos conversar sobre como podemos ser úteis para vocês? {sparky} \n_Confira nosso Site: bit.ly/MRsolutions_""")
mensagem_linkada = emoji.emojize(f"Com o nosso *super time de atendimento* a MR Solutions pode simplificar o suporte ao seu paciente. Confira nosso método eficiente de comunicação, desenvolvido especialmente para o atendimento de clínicas. "
                                 f" {chaty}_Estamos à disposição para esclarecer dúvidas e discutir possíveis soluções para sua empresa._{handy}")


# Carregar a planilha com os contatos
lista_contatos = openpyxl.load_workbook('contatos_teste.xlsx')
planilha = lista_contatos['Plan1']

# Iterar sobre as linhas da planilha e enviar mensagens
for linha in planilha.iter_rows(min_row=2, min_col=1, values_only=True):
    numero_cliente = linha[0]  # Primeira coluna, índice 0

    # Verificar se o número de telefone não é nulo
    if numero_cliente:
        # Enviar a mensagem de saudação
        pywhatkit.sendwhatmsg_instantly(numero_cliente, mensagem_saudacao, 25
                                        , tab_close=True)  # Não fechar a aba do navegador
        time.sleep(35)  # Aguardar um intervalo entre as mensagens

        # Enviar a segunda mensagem com a imagem e o link
        pywhatkit.sendwhats_image(numero_cliente,
                                  "C:\\Users\\morei\\PycharmProjects\\pythonProject1\\img_Mr_Solutions_Atendimento.jpeg",
                                  mensagem_linkada,
                                  25,  # Tempo de espera padrão de 40 segundos
                                  True,  # Não fechar outras abas do navegador
                                  35)  # Tempo de fechamento padrão de 35 segundos após o envio
    else:
        print("Número de telefone inválido ou vazio.")
