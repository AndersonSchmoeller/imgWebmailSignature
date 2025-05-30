from html2image import Html2Image
import cv2
import numpy as np
import os
import pyad.adquery
import win32api
import logging
logging.basicConfig(level=logging.DEBUG)

# Crie uma instância do Html2Image
hti = Html2Image()
hti.browser.use_new_headless = None
hti.browser.print_command = True         # prints the command line used to perform the screenshot

phoneImg = 'C:\\PATH\\TO\\IMAGE\\phone.png'
cellImg =  'C:\\PATH\\TO\\IMAGE\\cell.png'
mailImg =  'C:\\PATH\\TO\\IMAGE\\mail.png'
logoImg =  'C:\\PATH\\TO\\IMAGE\\logo.png'

# Função para criar imagem personalizada para cada usuário
def criar_imagem_para_usuario(samAccountName, displayName, telephoneNumber, mobile, email):
	if not telephoneNumber:
		telefone = '''
		'''
	else:
		telefone = fr'''
		<tr>
			<td><img src='{phoneImg}' width=16 height=16 /></td><td style='color: #727272; padding-left: 5px; margin-top:10px;'><a href='tel:{telephoneNumber}' style='font-size: 10px; text-decoration: none; color: #727272; font-family: Arial, Helvetica, sans-serif;'>{telephoneNumber}</a></td>
		</tr>
		'''
	if not mobile:
		celular = '''
		'''
	else:
		celular = fr'''
		<tr>
            <td><img src='{cellImg}' width=16 height=16 /></td><td style='color: #727272; padding-left: 3px; margin-top:10px;'><a href='tel:{mobile}' style='font-size: 10px; text-decoration: none; color: #727272; font-family: Arial, Helvetica, sans-serif;'>{mobile}</a></td>
        </tr>
		'''
    # Gerar HTML personalizado
	html = fr'''
<html>
    <head>

    </head>
    <body>
        <table style='background-color: white;' width='auto' height='117'>
            <tr>
                <td width=auto height=6 colspan=3 valign=top style='width:auto; border:none;'></td>
            </tr>
            <!--
                amarelo #fad902
                azul #081186
            -->
            <tr>
                <td width=20>
                </td>
                <td width=auto valign=top style='width: auto;border:none; padding-left: 7px; padding-right: 7px;'>
                    <div style='margin-top: 6px; font-size: 13px; color: #081186; font-family: Aptos ,Arial, Helvetica, sans-serif; font-weight: bold;'><span>{displayName}</span></div>
                    <table style='margin-top: 7px; border-spacing: 7px'>
						<!-- Descomentar linha -->
						<!-- 
						<tr>
							<td colspan=2 style='color: #727272; font-size: 10px; padding-left: 0px; margin-top:0px; font-family: "Times New Roman", Arial, Helvetica, sans-serif;'>Diretor Executivo</td>
                        </tr>
                         -->
                        {telefone}
                        {celular}   
                        <tr>
                            <td><img src='{mailImg}' width=16 height=16 /></td><td style='padding-left: 3px; margin-top:10px; font-size: 12px; color: #081186; font-family: Arial, Helvetica, sans-serif;'><a href='mailto:{email}' style='font-size: 10px; color: #081186; text-decoration: none;'>{email}</a></td>
                        </tr>
                    </table>
                </td>
                <td width=26 style='background-color: white;'>
                </td>
                <td width=auto style='border:none; border-left:solid windowtext 1.0pt; border-color: #a5a5a5; padding-left: 0px; padding-right: 0px;'></td>
                <td width=33 style='background-color: white;'>
                </td>
                <td><img src='{logoImg}' width='104' height='104'>
                </td>
                <td width=20>
                </td>
            </tr>
            <tr>
                <td width=auto height=6 colspan=7 valign=top style='text-align: center; width:auto; border:none; background:rgba(255, 255, 255, 0.315);'>
                    	<div style='margin-top:10px;'><span style="font-style: italic; color: #081186; font-size: 10px; font-family: Aptos, Arial, Helvetica, sans-serif; font-weight: bold;">
		    SLOGAN PHRASE
      			</div>
                </td>
            </tr>
        
        </table>
    </body>
</html>"
'''


	css = """ """
	
	
	print("[DEBUG] Iniciando geração da imagem...")
	# Converta a string HTML em imagem
	imagem = hti.screenshot(html_str=html, css_str=css, save_as='imagem.png')
	print("[DEBUG] Geração finalizada.")

	print('CARREGA A IMAGEM')
	# Carregar a imagem com fundo transparente
	imagem = cv2.imread('imagem.png', cv2.IMREAD_UNCHANGED)

	# Verificar se a imagem possui um canal alfa (transparência)
	if imagem.shape[2] == 4:
		# Extrair o canal alfa (transparência) da imagem
		_, _, _, alpha = cv2.split(imagem)

		# Encontrar os contornos do conteúdo não transparente
		contornos, _ = cv2.findContours(alpha, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

		# Encontrar o maior contorno (o contorno do conteúdo)
		maior_contorno = max(contornos, key=cv2.contourArea)

		# Calcular o retângulo delimitador do conteúdo
		x, y, w, h = cv2.boundingRect(maior_contorno)
		
		print('RECORTA A IMAGEM')
		# Recortar apenas o conteúdo da imagem
		conteudo = imagem[y:y+h, x:x+w]

		print('SALVA IMAGEM')
		# Salvar a imagem recortada e redimensionada
		cv2.imwrite(f'C:\\PATH\\TO\\SAVE\\IMAGE\\{samAccountName}.jpg', conteudo)
	else:
		print("A imagem não possui um canal alfa (transparência).")

samAccountName = 'imgFileName'
displayName = 'Firstname Lastname'
telephoneNumber = '(47) 1234-5678'
mobile = '(47) 91234-5678'
email = 'mail@domain.com'
	
	# Chamar a função para criar a imagem para o usuário atual
criar_imagem_para_usuario(samAccountName, displayName, telephoneNumber, mobile, email)
