#!/usr/bin/env python

"""
@File    :   export_calendrier_gab.py
@Time    :   2023/09/23 11:38:01
@Version :   1.0
@Desc    :   recuperation du calendrier sur cyu.fr
"""
import locale
import datetime
import keepasshttp
import logging
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from tabulate import tabulate
from reportlab.platypus import SimpleDocTemplate, Paragraph
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.pagesizes import A4
import tempfile
import win32api


def get_credential(url):
    """
    get_credential from keepass
    :param url:
    :return:
    """
    # get single credential
    credential = keepasshttp.get(url)
    logging.debug(credential.login)
    return credential


if __name__ == '__main__':
    logging.basicConfig(encoding='utf-8', level=logging.INFO)
    locale.setlocale(locale.LC_ALL, 'french_france')
    driver = webdriver.Firefox()

    jours = list(datetime.date(2001, 1, i).strftime('%A') for i in range(1, 6))
    start_date = datetime.date.today()
    if start_date.weekday() != 0:
        start_date += datetime.timedelta(days=7 - start_date.weekday())
    e_d_t = [jours]
    for ind_i, _ in enumerate(jours):
        n_date = start_date + datetime.timedelta(days=ind_i)
        e_d_t[0][ind_i] += n_date.strftime(' %d/%m')
    for ind_i in range(5):
        e_d_t.append([])
        for ind_j in range(len(jours)):
            e_d_t[ind_i+1].append('')

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

    for ind_i, col in enumerate(driver.find_elements(By.CSS_SELECTOR, ".fc-content-col")):
        for ind_j, output in enumerate(col.find_elements(By.XPATH, value="(.//*[contains(@class, 'fc-content')])")):
            if output.text:
                e_d_t[ind_j+1][ind_i] = output.text
    text_edt = tabulate(e_d_t, headers='firstrow')
    logging.debug('\'%s' % text_edt)

    driver.close()

    pdf_file_name = tempfile.mktemp(".pdf")

    styles = getSampleStyleSheet()
    h1 = styles["h1"]
    normal = styles["Code"]
    normal.fontSize = 6

    doc = SimpleDocTemplate(
        pdf_file_name,
        pagesize=(A4[1],A4[0]),
        leftMargin=0, rightMargin=0)

    text = text_edt.splitlines()
    story = []
    for line in text:
        # reportlab expects to see XML-compliant
        #  data; need to escape ampersands space and so on.
        story.append(Paragraph(line.replace(' ', '&nbsp;'), normal))

    doc.build(story)
    win32api.ShellExecute(0, "open", pdf_file_name, None, ".", 0)
