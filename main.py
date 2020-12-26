import shutil
import os
from datetime import datetime, timedelta
import time
from selenium import webdriver

Hoy = datetime.today()
FechaCompleta = datetime.strftime(Hoy, '%Y%m%d')
Anio = datetime.strftime(Hoy, '%Y')
FechaMes = datetime.strftime(Hoy, '%B')


def mes(valor):
    return {
        'January': 'Enero',
        'February': 'Febrero',
        'March': 'Marzo',
        'April': 'Abril',
        'May': 'Mayo',
        'June': 'Junio',
        'July': 'Julio',
        'August': 'Agosto',
        'September': 'Septiembre',
        'October': 'Octubre',
        'November': 'Noviembre',
        'December': 'Diciembre'
    }[valor]


Mes = mes(FechaMes)

try:
    driver = webdriver.Chrome('chromedriver.exe')
    urlVtr = 'https://vtrchile.sharepoint.com/sites/Cobranzas/Documentos%20compartidos/Forms/AllItems.aspx?'
    driver.get(urlVtr)
    time.sleep(10)

    # Inicio Login
    userLogin = driver.find_element_by_xpath('//*[@id="i0116"]')
    userLogin.send_keys('rgaete@vtr.cl')

    btnSig = driver.find_element_by_id('idSIButton9')
    btnSig.click()
    time.sleep(5)
    passLogin = driver.find_element_by_xpath('//*[@id="i0118"]')
    passLogin.send_keys('vtr.2020')
    time.sleep(5)

    btnLogin = driver.find_element_by_id('idSIButton9')
    btnLogin.click()

    btnIni = driver.find_element_by_id('idBtn_Back')
    btnIni.click()
    # Fin Login

    urlFijo = 'https://vtrchile.sharepoint.com/sites/Cobranzas/Documentos%20compartidos/Forms/AllItems.aspx?' \
              'FolderCTID=0x012000AD7F8A229C34AE4682AB9CFFD4B40872&sortField=Modified&isAscending=false&id=' \
              '%2Fsites%2FCobranzas%2FDocumentos%20compartidos%2FPagos%20Plataformas%2FPagos%20Fijo%2F' + Anio + '' \
              '%2F' + Mes + '%2FPagos%5FFijo%5F' + FechaCompleta + '%2Etxt&parent=%2Fsites%2FCobranzas%2FDocumentos' \
              '%20compartidos%2FPagos%20Plataformas%2FPagos%20Fijo%2F' + Anio + '%2F' + Mes + ''

    urlMovil = 'https://vtrchile.sharepoint.com/sites/Cobranzas/Documentos%20compartidos/Forms/AllItems.aspx?' \
               'FolderCTID=0x012000AD7F8A229C34AE4682AB9CFFD4B40872&sortField=Modified&isAscending=false&id=' \
               '%2Fsites%2FCobranzas%2FDocumentos%20compartidos%2FPagos%20Plataformas%2FPagos%20M%C3%B3vil' \
               '%2F' + Anio + '%2F' + Mes + '%2FPagos%5FMovil%5F' + FechaCompleta + '%2Etxt&parent=%2Fsites%2FCobranzas%2FDocumentos' \
               '%20compartidos%2FPagos%20Plataformas%2FPagos%20M%C3%B3vil%2F' + Anio + '%2F' + Mes + ''

    urlAndes = 'https://vtrchile.sharepoint.com/sites/Cobranzas/Documentos%20compartidos/Forms/AllItems.aspx?' \
               'FolderCTID=0x012000AD7F8A229C34AE4682AB9CFFD4B40872&sortField=Modified&isAscending=false&id=' \
               '%2Fsites%2FCobranzas%2FDocumentos%20compartidos%2FPagos%20Plataformas%2FPagos%20Andes%2F' + Anio + '' \
               '%2F' + Mes + '%2FPagos%5FAndes%5F' + FechaCompleta + '%2Etxt&parent=%2Fsites%2FCobranzas%2FDocumentos%20compartidos' \
               '%2FPagos%20Plataformas%2FPagos%20Andes%2F' + Anio + '%2F' + Mes + ''

    urlBotonDescarga = '//*[@id="appRoot"]/div/div[2]/div/div/div/div/div/div/div/div/div[1]/div[4]/button'

    driver.get(urlFijo)
    time.sleep(20)
    btnDescargar = driver.find_element_by_xpath(urlBotonDescarga)
    print(btnDescargar)
    btnDescargar.click()
    time.sleep(5)

    driver.get(urlMovil)
    time.sleep(15)
    btnDescargar = driver.find_element_by_xpath(urlBotonDescarga)
    btnDescargar.click()
    time.sleep(5)

    driver.get(urlAndes)
    time.sleep(15)
    btnDescargar = driver.find_element_by_xpath(urlBotonDescarga)
    btnDescargar.click()
    time.sleep(20)

    # Copiar archivos a la carpeta de prometeo
    nombreFijo = 'Pagos_Fijo_' + FechaCompleta + '.txt'
    nombreMovil = 'Pagos_Movil_' + FechaCompleta + '.txt'
    nombreAndes = 'Pagos_Andes_' + FechaCompleta + '.txt'

    urlOrigen = 'C:/Users/dmorir/Downloads/'
    urlDestino = '//Prometeo/TMP2/VTR/INPUT/'
    urlCargados = '//Prometeo/TMP2/VTR/CARGADOS/'

    def movearchivo(nombrearchivo):
        if os.path.exists(urlCargados + nombrearchivo):
            print('El archivo ' + nombrearchivo + ' ya se ha cargado.')
        else:
            if os.path.exists(urlDestino + nombrearchivo):
                print('El archivo ' + nombrearchivo + ' ya existe en la ruta')
            else:
                shutil.move(urlOrigen + nombrearchivo, urlDestino)

    movearchivo(nombreFijo)
    movearchivo(nombreMovil)
    movearchivo(nombreAndes)

    def removearchivo(nombrearchivo):
        if os.path.exists(urlOrigen + nombrearchivo):
            os.remove(urlOrigen + nombrearchivo)
            print('El archivo ' + nombrearchivo + ' fue eliminado')
        else:
            print('El archivo ' + nombrearchivo + ' no existe en la ruta')

except ValueError as e:
    print('error', e)
finally:
    removearchivo(nombreFijo)
    removearchivo(nombreMovil)
    removearchivo(nombreAndes)
    driver.close()
    driver.quit()
