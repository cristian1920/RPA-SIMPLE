import datetime
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email import encoders
from email.mime.text import MIMEText
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import os
import smtplib
from PIL import Image
import urllib
import urllib.request
from selenium import webdriver
import time
from datetime import datetime
from PyPDF2.merger import PdfFileMerger
from PyPDF2.pdf import PdfFileReader
import img2pdf


class InicioGeneral():

     def inicio():
          ruta="img/"
          driver = webdriver.Chrome(executable_path=r'./chromedriver.exe')        
          driver.maximize_window()
          driver.get('https://www.arus.com.co/')
          try:
               WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//img[contains(@alt,'Home personas arus')]")))
               print ("cargo  primera foto")
          except:
               print ("El elemento no cargo")
          url1 = driver.find_element_by_xpath("//img[contains(@alt,'Home personas arus')]").get_attribute("src")
          urllib.request.urlretrieve(url1, "img/home.png")
          time.sleep(2)
          driver.get('https://www.arus.com.co/suaporte/procesos-empresariales/administra')
          try:
               WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//img[contains(@alt,'MicrosoftTeams-image (38) (1)')]")))
               print ("cargo  segunda foto")
          except:
               print ("El elemento no cargo")
          url2 = driver.find_element_by_xpath("//img[contains(@alt,'MicrosoftTeams-image (38) (1)')]").get_attribute("src")
          urllib.request.urlretrieve(url2, "img/mst.png")
          InicioGeneral.CombinarArchivo(ruta)

     def CombinarArchivo(ruta):
          print("combina las imagenes")
          with open("Imagenes.pdf", "wb") as documento:
                    documento.write(img2pdf.convert(ruta+'home.png', ruta+'mst.png'))
          InicioGeneral.correo(ruta)



     def correo(ruta): 
          print("enviando el correo")  
          fecha = datetime.now()
          fecha = str(fecha)
          anio = datetime.strptime(fecha,'%Y-%m-%d %H:%M:%S.%f').strftime('%Y')
          mes = datetime.strptime(fecha,'%Y-%m-%d %H:%M:%S.%f').strftime('%m')
          dia = datetime.strptime(fecha,'%Y-%m-%d %H:%M:%S.%f').strftime('%d')   
          remitente = 'taydes.clondono@iumafis.edu.co'
          destinatario = ['camilo1211@hotmail.com','taydes.clondono@iumafis.edu.co','cristian131411@gmail.com']
          # Creamos el objeto mensaje
          mensaje = MIMEMultipart()
          # Establecemos los atributos del mensaje
          mensaje['From'] = remitente
          mensaje['To'] = ", ".join(destinatario)
          mensaje['Subject'] = 'PRUEBA DE DESARROLLO PARA ARUS'
          fp=open('Imagenes.pdf',"rb")
          adjunto= MIMEBase('multipart','encripte')
          adjunto.set_payload(fp.read())
          fp.close()
          encoders.encode_base64(adjunto)
          adjunto.add_header('Content-Disposition','attachment',filename='Imagenes.pdf')
          mensaje.attach(adjunto)
          # Creamos el cuerpo del mensaje
          cuerpo = f'Bello,{dia} del mes {mes} de {anio} <br><br>Señores<br><b>Arus</b><br>Prueba desarrollo<br>Cristian londoño meneses<br>Bello<br><br><br>Asunto: Prueba desarrollo<br><br>Cordial saludo,<br><br>  Se adjunta pdf con las imagenes obtenidas de la pagina arus combinados en un solo pdf.<br><br>Cualquier información adicional con gusto será atendida en el correo electrónico: <a href="mailto:cristian131411@gmail.com">cristian131411@gmail.com.</a><br>Muchas gracias.<br>Quedo atento.<br>Feliz día.'
          # Y lo agregamos al objeto mensaje como objeto MIME de tipo texto
          mensaje.attach(MIMEText(cuerpo, 'html'))
          # Creamos la conexión con el servidor
          sesion_smtp = smtplib.SMTP('smtp.office365.com', 587)
          # Ciframos la conexión
          sesion_smtp.starttls()
          # Iniciamos sesión en el servidor
          sesion_smtp.login('taydes.clondono@iumafis.edu.co','C1020464889*')
          # Convertimos el objeto mensaje a texto
          texto = mensaje.as_string()
          # Enviamos el mensaje
          sesion_smtp.sendmail(remitente, destinatario, texto)
          # Cerramos la conexión
          sesion_smtp.quit()
          print('Mensaje enviado Exitosamente')
