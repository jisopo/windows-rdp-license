import time
import urllib
import random
import os.path

from selenium.webdriver.support.ui import Select
from selenium import webdriver
from smb.SMBHandler import SMBHandler

FILE_SERVER_ADDRESS = 'file-server'
FILE_SERVER_USER_NAME = 'jisopo'
FILE_SERVER_PATH_FILE_PATH = 'smb://{}/incoming/{}/product_id.txt'.format(FILE_SERVER_ADDRESS, FILE_SERVER_USER_NAME)
FILE_SERVER_SERVER_ACTIVATION_KEY_FILE_PATH = 'smb://{}/incoming/{}/server_activation_key.txt'.format(FILE_SERVER_ADDRESS, FILE_SERVER_USER_NAME)
FILE_SERVER_CLIENTS_ACTIVATION_KEY_FILE_PATH = 'smb://{}/incoming/{}/client_license_activation_key.txt'.format(FILE_SERVER_ADDRESS, FILE_SERVER_USER_NAME)
MICROSOFT_ACTIVATION_URI = 'https://activate.microsoft.com/'

LANGUAGE_LIST_ID = "Content_ddlLanguageList"
BUTTON_NEXT_ID = "Content_btnNext"
COMPANY_NAME_ID = "Content_companyInfo_txtCompanyName"
COMPANY_COUNTRY_ID = "Content_companyInfo_ddlCountry"
COMPANY_PRODUCT_CODE_ID = "Content_pidControl_txtProductID"
KEY_FIELD_ID = "Content_lblLSID"
KEY_FIELD2_ID = "Content_lblLKPID"
INSTALL_CLIENT_LICENSES_RADIO_BUTTON_ID = "Content_rblActivate_1"
PRODUCT_NAME_KEY_ID = "Content_lsIDControl_txtLSID"
LICENSE_TYPE_ID = "Content_lpInfoControl_ddlLicenseProgram"
PRODUCT_TYPE_ID = "Content_prodType_ddlProductType"
PREFERED_LICENSE_COUNT_ID = "Content_txtQuantity"
LICENSE_NUMBER_AGREEMENT_ID = "Content_txtAgreementNumber"

PREFERED_LANGUAGE = "Russian"
PREFERED_COMPANY_NAME = "OOO Universe"
PREFERED_COMPANY_COUNTRY = "Египет"
PREFERED_LICENSE_TYPE = "Соглашение Enterprise agreement"
# Для активации лицензии на 2012 сервере поменять цифры с 2016 на 2012 в полях PREREFED_CALS_LICENSE_TYPE_DEVICE и PREREFED_CALS_LICENSE_TYPE_USER
# select_by_index - value = 020
PREREFED_CALS_LICENSE_TYPE_DEVICE = "Windows Server 2016 Клиентская лицензия доступа к службам удаленных рабочих столов \"на устройство\""
# select_by_index - value = 012
PREREFED_CALS_LICENSE_TYPE_USER = "Windows Server 2016 Клиентская лицензия доступа к службам удаленных рабочих столов \"на пользователя\""
PREFERED_LICENSE_COUNT = "10"
SELECT_RANDOM_COUNTRY = False
SELECT_ACTIVATION_BY_DEVICES = True
# дополнительный - 3325596
LICENSE_AGREEMENT_NUMBER= "5296992"

TEMP_SERVER_ACTIVATION_SERIAL_NUMBER_FILE = "server_product_key.txt"
TEMP_CLIENTS_ACTIVATION_SERIAL_NUMBER_FILE = "clients_license_product_key.txt"
PRODUCT_KEY_1 = ""
PRODUCT_KEY_2 = ""

SELECTED_COUNTRY_INDEX = -1

# секунды
FILE_SERVER_FILE_WAIT_INTERVAL = 2

opener = urllib.request.build_opener(SMBHandler)

def GetSerial():
    # TODO: добавить возможность чтения файла с диска
    print("Ожидаем серийный номер с файл сервера..")
    fh = None
    while(True):
        try:
            fh = opener.open(FILE_SERVER_PATH_FILE_PATH)
            data = fh.read()
            fh.close()
            serial = data.decode('utf8')
            print("Серийный номер получен")
            return serial
        except urllib.error.URLError as e:
            time.sleep(FILE_SERVER_FILE_WAIT_INTERVAL)

def UploadSerialNumer(serial_number, local_file_name, network_path):
    with open(local_file_name, "w", encoding='utf8') as f:
        f.write(serial_number)

    file_fh = open(local_file_name, 'rb')
    print("Загружаем серийный номер на сервер")
    fh = opener.open(network_path, data = file_fh)
    print("Готово")
    return


def SetUpServerLicense():
    # можно запускать в режиме --headless чтобы не было открытого окна браузера
    # но так нагляднее
    driver = webdriver.Chrome()
    driver.get(MICROSOFT_ACTIVATION_URI)

    lang_box = driver.find_element_by_id(LANGUAGE_LIST_ID)
    lang_box_selector = Select(lang_box)
    lang_box_selector.select_by_visible_text(PREFERED_LANGUAGE)

    next_button = driver.find_element_by_id(BUTTON_NEXT_ID)
    next_button.click()

    company_name = driver.find_element_by_id(COMPANY_NAME_ID)
    company_name.send_keys(PREFERED_COMPANY_NAME)

    company_country = driver.find_element_by_id(COMPANY_COUNTRY_ID)
    select_company_country = Select(company_country)
    country_items_count = len(select_company_country.options)

    if SELECT_RANDOM_COUNTRY:
        global SELECTED_COUNTRY_INDEX
        SELECTED_COUNTRY_INDEX = random.randint(0, country_items_count)
        select_company_country.select_by_index(SELECTED_COUNTRY_INDEX)
    else:
        select_company_country.select_by_visible_text(PREFERED_COMPANY_COUNTRY)

    company_product_code = driver.find_element_by_id(COMPANY_PRODUCT_CODE_ID)

    product_key = GetSerial()

    company_product_code.send_keys(product_key)

    next_button = driver.find_element_by_id(BUTTON_NEXT_ID)
    next_button.click()

    next_button = driver.find_element_by_id(BUTTON_NEXT_ID)
    next_button.click()

    serial_number = driver.find_element_by_id(KEY_FIELD_ID)
    global PRODUCT_KEY_1
    PRODUCT_KEY_1 = serial_number.text
    UploadSerialNumer(PRODUCT_KEY_1, TEMP_SERVER_ACTIVATION_SERIAL_NUMBER_FILE, FILE_SERVER_SERVER_ACTIVATION_KEY_FILE_PATH)

    driver.quit()

    return

def SetUpCALsLicense():
    global PRODUCT_KEY_1
    global SELECTED_COUNTRY_INDEX
    if PRODUCT_KEY_1 == "" and os.path.isfile(TEMP_SERVER_ACTIVATION_SERIAL_NUMBER_FILE):
        with open(TEMP_SERVER_ACTIVATION_SERIAL_NUMBER_FILE) as f:
            PRODUCT_KEY_1 = f.readline()

    driver = webdriver.Chrome()
    driver.get(MICROSOFT_ACTIVATION_URI)
    lang_box = driver.find_element_by_id(LANGUAGE_LIST_ID)
    select = Select(lang_box)
    select.select_by_visible_text(PREFERED_LANGUAGE)

    install_client_licenses = driver.find_element_by_id(INSTALL_CLIENT_LICENSES_RADIO_BUTTON_ID)
    install_client_licenses.click()

    next_button = driver.find_element_by_id(BUTTON_NEXT_ID)
    next_button.click()

    product_info_key = driver.find_element_by_id(PRODUCT_NAME_KEY_ID)
    product_info_key.send_keys(PRODUCT_KEY_1)

    company_name = driver.find_element_by_id(COMPANY_NAME_ID)
    company_name.send_keys(PREFERED_COMPANY_NAME)

    company_country = driver.find_element_by_id(COMPANY_COUNTRY_ID)
    select_company_country = Select(company_country)
    country_items_count = len(select_company_country.options)

    if SELECT_RANDOM_COUNTRY:
        select_company_country.select_by_index(SELECTED_COUNTRY_INDEX)
    else:
        select_company_country.select_by_visible_text(PREFERED_COMPANY_COUNTRY)

    license_type = driver.find_element_by_id(LICENSE_TYPE_ID)
    license_type_selector = Select(license_type)
    license_type_selector.select_by_visible_text(PREFERED_LICENSE_TYPE)

    next_button = driver.find_element_by_id(BUTTON_NEXT_ID)
    next_button.click()

    lcsa_product_type = driver.find_element_by_id(PRODUCT_TYPE_ID)
    lcsa_product_type_selector = Select(lcsa_product_type)
    if(SELECT_ACTIVATION_BY_DEVICES):
        lcsa_product_type_selector.select_by_visible_text(PREREFED_CALS_LICENSE_TYPE_DEVICE)
    else:
        lcsa_product_type_selector.select_by_visible_text(PREREFED_CALS_LICENSE_TYPE_USER)

    licenses_count = driver.find_element_by_id(PREFERED_LICENSE_COUNT_ID)
    licenses_count.send_keys(PREFERED_LICENSE_COUNT)

    license_agreement_number = driver.find_element_by_id(LICENSE_NUMBER_AGREEMENT_ID)
    license_agreement_number.send_keys(LICENSE_AGREEMENT_NUMBER)

    next_button = driver.find_element_by_id(BUTTON_NEXT_ID)
    next_button.click()

    next_button = driver.find_element_by_id(BUTTON_NEXT_ID)
    next_button.click()

    # TODO: навернуть проверки если страница не доступа и тд
    serial_number = driver.find_element_by_id(KEY_FIELD2_ID)
    global PRODUCT_KEY_2
    PRODUCT_KEY_2 = serial_number.text
    UploadSerialNumer(PRODUCT_KEY_2, TEMP_CLIENTS_ACTIVATION_SERIAL_NUMBER_FILE, FILE_SERVER_CLIENTS_ACTIVATION_KEY_FILE_PATH)

    driver.quit()

    return

SetUpServerLicense()
SetUpCALsLicense()
print("все готово")