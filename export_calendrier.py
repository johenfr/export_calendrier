#!/usr/bin/env python

"""
@File    :   export_calendrier_gab.py
@Time    :   2023/09/23 11:38:01
@Version :   1.0
@Desc    :   recuperation du calendrier sur cyu.fr
"""
import locale
import os
import datetime
import logging
import time
import pickle
import subprocess

import keepassxc_proxy_client
import keepassxc_proxy_client.protocol

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options

from openpyxl import Workbook
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.worksheet.worksheet import Worksheet


class Dict2ClassEmpty(object):
    def login(self):
        return ''
    
    def password(self):
        return ''
    
    
class Dict2Class(Dict2ClassEmpty):

    def __init__(self, my_dict):
        for key in my_dict:
            setattr(self, key, my_dict[key])


def get_credential(url):
    """
    get_credential from keepass
    :param url:
    :return:
    """
    # get single credential
    connection = keepassxc_proxy_client.protocol.Connection()
    connection.connect()
    _home = os.environ.get('HOME', '')
    if not os.path.exists(os.path.join(_home, '.ssh', 'python_keepassxc')):
        connection.associate()
        name, public_key = connection.dump_associate()
        print("Got connection named '", name, "' with key", public_key)
        with open(os.path.join(_home, '.ssh', 'python_keepassxc'), 'wb') as fic:
            fic.write(public_key)
    with open(os.path.join(_home, '.ssh', 'python_keepassxc'), 'rb') as fic:
        public_key = fic.read()
    connection.load_associate('python', public_key)
    connection.test_associate()
    credentials = connection.get_logins(url)
    if credentials:
        credential = Dict2Class(credentials[0])
        logging.debug(credential.login)
        return credential
    else:
        return Dict2ClassEmpty()


if __name__ == '__main__':
    logging.basicConfig(encoding='utf-8', level=logging.INFO)
    my_creds = get_credential("https://cyu.fr/")
    locale.setlocale(locale.LC_ALL, 'fr_FR.utf8')
    options = Options()
    options.binary_location = '/usr/lib/firefox/firefox'

    jours = list(datetime.date(2001, 1, i).strftime('%A') for i in range(1, 7))
    start_date = datetime.date.today()
    if start_date.weekday() != 0:
        start_date += datetime.timedelta(days=7 - start_date.weekday())
    dates = jours
    for ind_i, _ in enumerate(jours):
        n_date = start_date + datetime.timedelta(days=ind_i)
        dates[ind_i] += n_date.strftime(' %d/%m')
    logging.debug(dates)
    if os.path.exists('e_d_t.dat'):
        # vérification de la date d'export
        now = datetime.datetime.now()
        debut = now - datetime.timedelta(days=6)
        st = os.stat('e_d_t.dat')
        date_f = datetime.datetime.fromtimestamp(st.st_mtime)
        if date_f < debut:
            os.remove('e_d_t.dat')
    if os.path.exists('e_d_t.dat'):
        # export encore valide
        with open('e_d_t.dat', 'rb') as driver_dat:
            e_d_t = pickle.load(driver_dat)
    else:
        # pas d'export disponible → on récupère depuis cyu.fr
        driver = webdriver.Firefox()
        my_creds = get_credential("https://cyu.fr/")
        my_creds2 = get_credential("https://services-web.cyu.fr/calendar/cal")

        URL = "https://services-web.cyu.fr/calendar/"
        LOGIN_ROUTE = "LdapLogin/"
        DATA_ROUTE = "cal?vt=agendaWeek&dt=%s&et=student&fid0=%s" % (start_date.isoformat(), my_creds2.login)

        driver.get(URL+LOGIN_ROUTE)

        uname = driver.find_element(By.ID, "Name")
        uname.send_keys(my_creds.login)
        upswd = driver.find_element(By.ID, "Password")
        upswd.send_keys(my_creds.password)
        driver.find_element(By.CSS_SELECTOR, ".loginBtn").click()
        driver.get(URL+DATA_ROUTE)
        time.sleep(1)

        e_d_t = [['Horaire', 'Cours / TD', 'Salle', 'Prof.']]
        for ind_i, col in enumerate(driver.find_elements(By.CSS_SELECTOR, ".fc-content-col")):
            e_d_t.append([dates[ind_i], '', '', ''])
            for ind_j, output in enumerate(col.find_elements(By.XPATH, value="(.//*[contains(@class, 'fc-content')])")):
                if output.text:
                    heure_i, sans_lignes = output.text.split('\n', 1)
                    contenu_i, suite_i = sans_lignes.split('\nCH', 1)
                    salle_i = 'CH' + suite_i.split('p\n')[0] + 'p'
                    salle_i = '\n'.join(salle_i.split(' ', 1))
                    salle_i = '\n'.join(salle_i.rsplit(' ', 1))
                    try:
                        prof_i = suite_i.split('p\n')[1]
                    except IndexError:
                        prof_i = ''
                    e_d_t.append(['- %s' % heure_i, contenu_i, salle_i, prof_i])
        driver.close()
        with open('e_d_t.dat', 'wb') as driver_dat:
            pickle.dump(e_d_t, driver_dat)

    wb = Workbook()
    ws = wb.active
    if ws is not None:
        for ligne in e_d_t:
            ws.append(ligne)
        ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True, autoPageBreaks=False)
        dims = {}
        for row in ws.rows:
            for cell in row:
                if cell.value:
                    dims[cell.column] = max([len(ligne) for ligne in str(cell.value).splitlines()]) + 5
        for col, value in dims.items():
            ws.column_dimensions[chr(64 + col)].width = value
        Worksheet.set_printer_settings(ws, paper_size=9, orientation='landscape')
        wb.save("e_d_t.xlsx")
        subprocess.run(['open', "e_d_t.xlsx"], check=False)
