#!/usr/bin/env python

"""
@File    :   export_calendrier.py
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
from openpyxl.styles import PatternFill
from openpyxl.styles import Font

import tkinter.messagebox as msg

etudiant="Gabriel"


class Dict2ClassEmpty(object):
    def login(self):
        return ""

    def password(self):
        return ""


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
    global etudiant
    # get single credential
    connection = keepassxc_proxy_client.protocol.Connection()
    connection.connect()
    _home = os.environ.get("HOME", "")
    if not os.path.exists(os.path.join(_home, ".ssh", "python_keepassxc")):
        connection.associate()
        name, public_key = connection.dump_associate()
        print("Got connection named '", name, "' with key", public_key)
        with open(os.path.join(_home, ".ssh", "python_keepassxc"), "wb") as fic:
            fic.write(public_key)
    with open(os.path.join(_home, ".ssh", "python_keepassxc"), "rb") as fic:
        public_key = fic.read()
    connection.load_associate("python", public_key)
    connection.test_associate()
    credentials = connection.get_logins(url)
    if credentials:
        for gugusse in credentials:
            wanted_etudiant = "e-%s" % etudiant[0].lower()
            if wanted_etudiant in gugusse["login"]:
                credential = Dict2Class(gugusse)
                logging.debug(credential.login)
                return credential
    return Dict2ClassEmpty()


if __name__ == "__main__":
    for choix in ["Gabriel", "Louis-Joseph"]:
        if msg.askyesno("Etudiant?", choix):
            etudiant = choix
            break
    logging.basicConfig(encoding="utf-8", level=logging.INFO)
    my_creds = Dict2ClassEmpty()
    locale.setlocale(locale.LC_ALL, "fr_FR.utf8")
    options = Options()
    options.binary_location = "/usr/lib/firefox/firefox"
    driver = None

    jours = list(datetime.date(2001, 1, i).strftime("%A") for i in range(1, 7))
    start_date = datetime.date.today()
    if start_date.weekday() != 0:
        start_date += datetime.timedelta(days=7 - start_date.weekday())
    dates = jours
    for ind_i, _ in enumerate(jours):
        n_date = start_date + datetime.timedelta(days=ind_i)
        dates[ind_i] += n_date.strftime(" %d/%m")
    logging.debug(dates)
    if os.path.exists("e_d_t_%s.dat" % etudiant.lower()):
        # vérification de la date d'export
        now = datetime.datetime.now()
        debut = now - datetime.timedelta(days=6)
        st = os.stat("e_d_t_%s.dat" % etudiant.lower())
        date_f = datetime.datetime.fromtimestamp(st.st_mtime)
        if date_f < debut:
            os.remove("e_d_t_%s.dat" % etudiant.lower())
    if os.path.exists("e_d_t_%s.dat" % etudiant.lower()):
        # export encore valide
        with open("e_d_t_%s.dat" % etudiant.lower(), "rb") as driver_dat:
            e_d_t = pickle.load(driver_dat)
    else:
        # pas d'export disponible → on récupère depuis cyu.fr
        my_creds = get_credential("https://cyu.fr/")
        my_creds2 = get_credential("https://services-web.cyu.fr/calendar/cal")
        
        driver = webdriver.Firefox()
        URL = "https://services-web.cyu.fr/calendar/"
        LOGIN_ROUTE = "LdapLogin/"
        DATA_ROUTE = "cal?vt=agendaWeek&dt=%s&et=student&fid0=%s" % (start_date.isoformat(), my_creds2.password)

        driver.get(URL + LOGIN_ROUTE)

        uname = driver.find_element(By.ID, "Name")
        uname.send_keys(my_creds.login)
        upswd = driver.find_element(By.ID, "Password")
        upswd.send_keys(my_creds.password)
        driver.find_element(By.CSS_SELECTOR, ".loginBtn").click()
        driver.get(URL + DATA_ROUTE)
        time.sleep(1)

        e_d_t = [["Horaire", "Cours / TD", "Salle", "Prof.", "          "]]
        for ind_i, col in enumerate(driver.find_elements(By.CSS_SELECTOR, ".fc-content-col")):
            e_d_t.append([dates[ind_i], "", "", ""])
            for ind_j, output in enumerate(col.find_elements(By.XPATH, value="(.//*[contains(@class, 'fc-content')])")):
                if output.text:
                    heure_i, sans_lignes = output.text.split("\n", 1)
                    prof_i = ""
                    try:
                        contenu_i, suite_i = sans_lignes.split("\nCH", 1)
                        salle_i = "CH" + suite_i.split("p\n")[0] + "p"
                        salle_i = "\n".join(salle_i.split(" ", 1))
                        salle_i = "\n".join(salle_i.rsplit(" ", 1))
                    except ValueError:
                        try:
                            sans_lignes_1 = sans_lignes.split("\n", 1)[0]
                            sans_lignes_2 = sans_lignes.split(")\n", 1)[1]
                            sans_lignes = "\n".join([sans_lignes_1, sans_lignes_2])
                        except IndexError:
                            pass
                        contenu_i = ""
                        salle_i = ""
                        suite_i = ""
                        for site in ["NEU", "TUR"]:
                            try:
                                contenu_i, suite_i = sans_lignes.split("\n%s" % site, 1)
                                salle_i = site + suite_i.split("p\n")[0] + "p"
                                break
                            except ValueError:
                                continue
                        if not contenu_i:
                            if "à distance" in sans_lignes:
                                reconstruction = sans_lignes.split("\n")
                                reconstruction.pop()
                                prof_i = reconstruction.pop()
                                contenu_i = "\n".join(reconstruction)
                                salle_i = "à la maison"
                            else:
                                contenu_i, suite_i = sans_lignes.split("\n", 1)
                                salle_i = suite_i.split("p\n")[0] + "p"
                    try:
                        if salle_i != "à la maison":
                            prof_i = suite_i.split("p\n")[1].split('\n')[0]
                    except IndexError:
                        pass
                    e_d_t.append(["- %s" % heure_i, contenu_i, salle_i, prof_i, ""])
        with open("e_d_t_%s.dat" % etudiant.lower(), "wb") as driver_dat:
            pickle.dump(e_d_t, driver_dat)

    wb = Workbook()
    ws = wb.active
    blueFill = PatternFill(start_color="FFAAAAFF", end_color="FFAAAAFF", fill_type="solid")
    if ws is None:
        exit(1)
    for ligne in e_d_t:
        ws.append(ligne)
    for row in ws.rows:
        for cell in row:
            cell.font = Font(bold=True)
        break

    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True, autoPageBreaks=False)
    dims = {}
    for row in ws.rows:
        for cell in row:
            if cell.value:
                dims[cell.column] = max([len(ligne) for ligne in str(cell.value).splitlines()]) + 5
            if not row[2].value:
                cell.fill = blueFill
                cell.font = Font(bold=True)
    for col, value in dims.items():
        ws.column_dimensions[chr(64 + col)].width = value
    if etudiant == "Gabriel":
        Worksheet.set_printer_settings(ws, paper_size=9, orientation="landscape")
    else:
        Worksheet.set_printer_settings(ws, paper_size=9, orientation="portrait")
    wb.save("e_d_t_%s.xlsx" % etudiant.lower())
    subprocess.run(["open", "e_d_t_%s.xlsx" % etudiant.lower()], check=False)
    if driver:
        driver.close()
