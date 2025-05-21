"""
Отчет по событиям: выгрузка данных из PERCo, формирование Excel и отправка по email.
"""

import calendar
import smtplib
import sys
import datetime
import logging
import os
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate
from os.path import basename

import win32com.client
import xlsxwriter
from validate_email import validate_email
from dotenv import load_dotenv

# Настройка логирования
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    handlers=[
        logging.FileHandler('percopy.log', encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)

load_dotenv()

perco_server = os.getenv('PERCO_SERVER')
perco_port = int(os.getenv('PERCO_PORT'))
perco_username = os.getenv('PERCO_USERNAME')
perco_password = os.getenv('PERCO_PASSWORD')

smtp_server = os.getenv('SMTP_SERVER')
smtp_port = int(os.getenv('SMTP_PORT'))
smtp_username = os.getenv('SMTP_USERNAME')
smtp_password = os.getenv('SMTP_PASSWORD')

send_from = os.getenv('SEND_FROM')

mail_subject = os.getenv('MAIL_SUBJECT')
mail_text = os.getenv('MAIL_TEXT')

filename = os.getenv('FILENAME')

msxml = win32com.client.DispatchEx('MSXML2.DOMDocument.3.0')
perco = win32com.client.DispatchEx('PERCo_S20_SDK.ExchangeMain')

id_event = os.getenv('ID_EVENT')
id_resource = os.getenv('ID_DEVICE')


def period_begin_end_calculate() -> tuple:
    """
    Функция для вычисления начала и конца периода.
    Вычисляет начало и конец предыдущего месяца, если текушая дата до 15 числа.
    Вычисляет начало и конец текущего месяца, если текушая дата после 15 числа.
    :return: tuple (beginperiod, endperiod)
    """
    today = datetime.date.today()
    if today.day <= 15:
        first = today.replace(day=1)
        last_month = first - datetime.timedelta(days=1)
        last_day = calendar.monthrange(int(last_month.strftime("%Y")),
                                       int(last_month.strftime("%m")))[1]
        beginperiod = '01.' + str(last_month.strftime("%m")) + '.' + str(last_month.strftime("%Y"))
        endperiod = str(last_day) + '.' + \
                    str(last_month.strftime("%m")) + '.' + \
                    str(last_month.strftime("%Y"))
    else:
        beginperiod = '01.' + str(today.strftime("%m")) + '.' + str(today.strftime("%Y"))
        endperiod = str(calendar.monthrange(int(today.strftime("%Y")),
                                            int(today.strftime("%m")))[1]) + '.' + \
                     str(today.strftime("%m")) + '.' + \
                     str(today.strftime("%Y"))
    return beginperiod, endperiod


def send_mail(email: str) -> bool:
    """
    Отправка отчета на почту
    :return: True, если успешно, иначе False
    """
    logging.info('Отправка отчета на почту: %s', email)
    try:
        msg = MIMEMultipart()
        msg['From'] = send_from
        msg['To'] = COMMASPACE.join(email)
        msg['Date'] = formatdate(localtime=True)
        msg['Subject'] = mail_subject + ' ' + str(datetime.date.today())

        msg.attach(MIMEText(mail_text + ' ' + str(datetime.date.today()), 'plain'))

        with open(filename, "rb") as fil:
            part = MIMEApplication(
                fil.read(),
                Name=basename(filename)
            )
        part['Content-Disposition'] = f'attachment; filename="{basename(filename)}"'
        msg.attach(part)

        smtp = smtplib.SMTP(smtp_server, smtp_port)
        smtp.starttls()
        smtp.login(smtp_username, smtp_password)
        smtp.sendmail(send_from, email, msg.as_string())
        smtp.close()

        return True

    except (OSError, IOError, smtplib.SMTPException) as e:
        logging.error('Error sending email: %s', e)
        return False


def make_xml_for_get_data() -> win32com.client.CDispatch:
    """
    Создание xml для получения данных из perco
    :return: XML документ
    """
    beginperiod, endperiod = period_begin_end_calculate()
    header = msxml.createProcessingInstruction('xml',
                                               'version="1.0" encoding="UTF-8" standalone="yes"')
    msxml.appendChild(header)
    elem = msxml.createElement('documentrequest')
    elem.setAttribute('type', 'regevents')
    node = msxml.appendChild(elem)
    elem = msxml.createElement('eventsreport')
    elem.setAttribute('beginperiod', beginperiod)
    elem.setAttribute('endperiod', endperiod)
    elem.setAttribute('beginperiodtime', '00:00:00')
    elem.setAttribute('endperiodtime', '23:59:59')
    elem.setAttribute('id_resource_internal', id_resource)
    elem.setAttribute('id_event_internal', id_event)

    node = node.appendChild(elem)
    elem_root = msxml.createElement('events')
    node.appendChild(elem_root)

    return msxml


def get_xml_data() -> win32com.client.CDispatch:
    """
    Получение xml данных из perco
    :return:
    """
    getdata_xml = make_xml_for_get_data()

    if perco.SetConnect(perco_server, perco_port, perco_username, perco_password) != 0:
        error = win32com.client.DispatchEx('MSXML2.DOMDocument.3.0')
        perco.GetErrorDescription(error)
        logging.error('Не удалось подключиться к серверу PERCo. Ошибка: %s', error.xml)
        return getdata_xml

    v = perco.CheckVersion()
    if v[0] != 0:
        logging.error('Версия библиотеки SDK и сервера не совпадают.')
        perco.DisConnect()
        sys.exit(1)

    if perco.GetData(getdata_xml) != 0:
        error = win32com.client.DispatchEx('MSXML2.DOMDocument.3.0')
        perco.GetErrorDescription(error)
        response = error.xml
        logging.error('Ошибка получения данных: %s', response)
        sys.exit(1)
    perco.DisConnect()

    if getdata_xml.GetElementsByTagName('documentrequest').length == 0:
        logging.error('Не удалось получить данные из сервера PERCo.')
        sys.exit(1)
    # save xml
    with open('getdata.xml', 'w', encoding='utf-8') as f:
        f.write(getdata_xml.xml)
        f.close()
        logging.info('XML данные успешно сохранены в файл: %s', 'getdata.xml')

    return getdata_xml


def save_data_to_xlsx(xml: win32com.client.CDispatch) -> bool:
    """
    Сохранение данных в формате xlsx
    :param xml: XML данные
    :return: Успех операции
    """
    try:
        workbook = xlsxwriter.Workbook(filename)
        worksheet = workbook.add_worksheet('Отчет')
    except (xlsxwriter.exceptions.XlsxWriterException, OSError, IOError) as e:
        logging.error('Error creating Excel workbook: %s', e)
        return False

    worksheet.write(0, 0, 'ФИО')
    worksheet.write(0, 1, 'Дата')
    worksheet.write(0, 2, 'Подразделение')

    try:
        events = xml.getElementsByTagName('event')
    except AttributeError as e:
        logging.error('Attribute error getting events: %s', e)
        return False

    if events.length == 0:
        logging.warning('Нет данных для сохранения.')
        return False
    for i in range(events.length):
        event = events[i]
        try:
            worksheet.write(i + 1, 0, event.getAttribute('f_fio'))
            worksheet.write(i + 1, 1, event.getAttribute('f_date_ev'))
            worksheet.write(i + 1, 2, event.getAttribute('f_name_subdiv'))
        except AttributeError as e:
            logging.error('Attribute error writing to Excel: %s', e)
            continue
    try:
        workbook.close()
    except (xlsxwriter.exceptions.XlsxWriterException, OSError, IOError) as e:
        logging.error('Error closing Excel workbook: %s', e)
        return False

    return True


def main():
    """
    Основная функция, которая выполняет следующие действия:
    1. Проверяет, что передан email адрес в параметрах скрипта.
    2. Проверяет корректность email адреса.
    3. Получает данные из XML.
    4. Сохраняет данные в формате XLSX.
    5. Отправляет отчет на почту.
    :return: None
    """

    xml = get_xml_data()
    if xml is None:
        logging.error('Получен пустой xml.')
        return

    if not save_data_to_xlsx(xml):
        logging.error('Не удалось сохранить данные в xlsx.')
        return

    if len(sys.argv) < 2:
        logging.error('Необходимо ввести email адрес в параметрах \
            скрипта для отправки отчета на email.')
        return

    if not validate_email(sys.argv[1]):
        logging.error('Invalid email address.')
        return

    if not send_mail(sys.argv[1]):
        logging.error('Не удалось отправить отчет на почту.')
        return
    logging.info('Отчет отправлен на почту: %s', sys.argv[1])


if __name__ == '__main__':
    main()
