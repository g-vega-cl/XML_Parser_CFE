import subprocess
import sys

"""
def pip_install(package):
    subprocess.check_call([sys.executable, "-m", "pip", "install", package])


pip_install("pandas")
pip_install("selenium")
pip_install("webdriver_manager")
pip_install("openpyxl")
"""

import os
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
import time
from calendar import monthrange
import datetime
import random
import pandas as pd
import tkinter
from tkinter import messagebox

if __name__ == "__main__":

	JS_value = sys.argv[1]

	folderString = JS_value[0] + ":\\" + JS_value[3:]

	chromeOptions = webdriver.ChromeOptions()
	prefs = {"download.default_directory" : folderString}

	chromeOptions.add_experimental_option("prefs",prefs)
	time.sleep(.5)
	driver = webdriver.Chrome(ChromeDriverManager().install(), chrome_options=chromeOptions)
	time.sleep(1)
	driver.get("https://app.cfe.mx/Aplicaciones/CCFE/Recibos/Consulta/login.aspx?ReturnUrl=%2FAplicaciones%2FCCFE%2FRecibos%2FConsulta")
	time.sleep(1)

	while True:
		time.sleep(3)
		try:
			facturas_id =  driver.find_element_by_id("ctl00_PHContenidoPag_gvFacturasUsuario")
			facturas_pdf = driver.find_elements_by_class_name("pdf")
			break
		except:
			continue
	

	for i in range(len(facturas_pdf)):
		PDF_ID = facturas_pdf[i].get_attribute("id")
		XML_ID = PDF_ID[:-3] + "XML"
		current_xml = driver.find_element_by_id(XML_ID)
		current_xml.click()
		time.sleep(3)

	time.sleep(10)

	driver.quit()

