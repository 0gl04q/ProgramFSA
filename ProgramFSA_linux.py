import os
import time

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.action_chains import ActionChains

from openpyxl import load_workbook


# Запуск программы через Chrome
def chrome(dir_drive):
    # Подключение Chrome
    service = Service(executable_path=dir_drive)
    driver = webdriver.Chrome(service=service)

    # Уменьшение масштаба
    driver.get("chrome://settings/")
    driver.execute_script("chrome.settingsPrivate.setDefaultZoom(0.5)")
    return main(driver)


# Запуск программы через FireFox
def firefox(dir_drive, head):
    # Подключение опций
    options = Options()
    if head:
        options.headless = True
        options.add_argument('--window-size=1920,1920')
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-gpu')

    # Подключение FireFox
    service = Service(executable_path=dir_drive)
    driver = webdriver.Firefox(service=service, options=options)

    # Уменьшение масштаба страницы FireFox
    driver.get("about:preferences")
    driver.execute_script("arguments[0].click();", driver.find_element(By.XPATH, "//*[@id='defaultZoom']"))
    ActionChains(driver).click(driver.find_element(By.XPATH, "//*[@value='50']")).perform()
    return main(driver)


# функция работы программы
def main(driver):
    driver.get("https://support.fsa.gov.ru")  # Сайт ФСА

    folder = r'/opt/ProgramFSA/File'  # Объявление папки для работы с файлами

    for file in os.listdir(folder):  # Перебор файлов со сведениями
        name_company = file[:3]
        print(file)

        # Проверка папки для скриншотов + создание
        if not os.path.exists(fr'/opt/ProgramFSA/Screenshot/{file}'):
            os.mkdir(fr'/opt/ProgramFSA/Screenshot/{file}')

        # Открытие файла Excel
        wb = load_workbook(folder + "/" + file)
        sheet = wb.active

        # Ожидание окончания загрузки сайта
        try:
            WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH, "//div[@data-t='for_metrology']")))
        except:
            # получение доступа к элементу сайта
            driver.execute_script("arguments[0].click();",
                                  driver.find_element(By.XPATH, "//div[@data-t='for_metrology']"))

        # проход по строкам начиная с 5
        nom_str = 5
        str_enumeration = sheet.cell(row=nom_str, column=1).value
        while str_enumeration is not None:
            # Открытие формы для заполнения данных
            driver.execute_script("arguments[0].click();", driver.find_element(By.XPATH, "//*[@id='metrology-report']"))

            # Выбор формы для поверки
            Select(driver.find_element(By.NAME, "alType")).select_by_index(3)

            # Проверка наименования организации и внесение рег номера аккредитации
            if name_company == "АТМ":
                Select(driver.find_element(By.XPATH,
                                           "//*[@id='metrologyReportForm']/div[2]/div/div[1]/select")).select_by_index(
                    2)
                driver.find_element(By.XPATH, "//*[@id='metrologyReportForm']/div[2]/div/div[1]/div/input").send_keys(
                    sheet.cell(row=1, column=1).value)
            else:
                Select(driver.find_element(By.XPATH,
                                           "//*[@id='metrologyReportForm']/div[2]/div/div[1]/select")).select_by_index(
                    0)
                driver.find_element(By.XPATH, "//*[@id='metrologyReportForm']/div[2]/div/div[1]/div/input").send_keys(
                    sheet.cell(row=1, column=1).value)

            # Внесение сведений в форму
            driver.find_element(By.XPATH, "//*[@id='metrologyReportForm']/div[3]/div/div/input").send_keys(
                sheet.cell(row=1, column=2).value)

            driver.find_element(By.XPATH, "//*[@id='measurementsForm']/div[1]/div/div/input").send_keys(
                str(sheet.cell(row=nom_str, column=1).value) + " " + str(sheet.cell(row=nom_str,
                                                                                    column=2).value) + " Зав.№" + str(
                    sheet.cell(row=nom_str, column=3).value))

            driver.find_element(By.XPATH, "//*[@id='measurementsForm']/div[2]/div/div/input").send_keys(
                sheet.cell(row=nom_str, column=4).value.strftime('%Y-%m-%d'))

            driver.find_element(By.XPATH, "//*[@id='measurementsForm']/div[3]/div/div/input").send_keys(
                sheet.cell(row=nom_str, column=5).value)

            driver.find_element(By.XPATH, "//*[@id='measurementsForm']/div[4]/div/div/input").send_keys(
                str(sheet.cell(row=nom_str, column=1).value) + " " + str(sheet.cell(row=nom_str,
                                                                                    column=2).value) + " Зав.№" + str(
                    sheet.cell(row=nom_str, column=3).value))

            driver.find_element(By.XPATH, "//*[@id='measurementsForm']/div[5]/div/div/input").send_keys(
                sheet.cell(row=nom_str, column=7).value + " " + sheet.cell(row=nom_str, column=8).value)

            driver.find_element(By.XPATH, "//*[@id='measurementsForm']/div[6]/div[1]/input").send_keys(
                sheet.cell(row=nom_str, column=9).value)
            driver.find_element(By.XPATH, "//*[@id='measurementsForm']/div[6]/div[2]/input").send_keys(
                sheet.cell(row=nom_str, column=10).value)
            driver.find_element(By.XPATH, "//*[@id='measurementsForm']/div[6]/div[3]/input").send_keys(
                sheet.cell(row=nom_str, column=11).value)

            # Создание скриншота
            driver.save_screenshot('/opt/ProgramFSA/Screenshot/' + file + '/' + str(
                nom_str - 4) + ' ' + str(sheet.cell(row=nom_str, column=4).value.strftime('%d.%m.%Y')) + ' ' +
                                   name_company.replace(' ', '') + ".png")

            print('Сохранение скриншота заполненного счетчика, строка:', nom_str - 4)

            # Проверка отправки формы
            check_str = driver.find_element(By.XPATH,
                                            "//*[@id='metrologyReportForm']/div[3]/div/div/input").get_attribute(
                "value")

            i = len(check_str)

            while i > 0:
                i = len(
                    driver.find_element(By.XPATH, "//*[@id='metrologyReportForm']/div[3]/div/div/input").get_attribute(
                        "value"))
                if i > 0:
                    driver.execute_script("arguments[0].click();",
                                          driver.find_element(By.XPATH, "//*[@id='metrology-report-submit']"))
                    time.sleep(20)

            # Увеличение счетчика строки для прохода по строкам
            nom_str = nom_str + 1
            str_enumeration = sheet.cell(row=nom_str, column=1).value
            time.sleep(20)

        # Закрытие файла Excel
        wb.close()

    # Закрытие драйвера
    driver.close()


if __name__ == '__main__':

    # Инициализация программы
    init_prog = input('Использовать предустановленные данные?(Да-1, Нет-2): ')
    if init_prog == 2:
        dit_drive = input(r'Введите путь к драйверу("/opt/ProgramFSA/geckodriver"): ')
        brow = input('Введите предпочитаемый браузер(chrome-1, firefox-2): ')
        if brow == '2':
            head_change = input('Выберите режим(Headless-1; Head-2): ')
            if head_change == '1':
                firefox(dit_drive, head=True)
            elif head_change == '2':
                firefox(dit_drive, head=False)
            else:
                print('Неверно выбран режим')
        elif brow == '1':
            chrome(dit_drive)
        else:
            print('Неверно выбран вариант')
    elif init_prog == '1':
        dit_driver = '/opt/ProgramFSA/geckodriver'
        firefox(dit_driver, head=True)

    print('Конец работы программы')
