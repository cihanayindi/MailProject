#-------------------KÜTÜPHANE------------------#
#----------------------------------------------#

import sys
from sqlite3 import Cursor
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import *
from AnaSayfaUI import *
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

#---------------UYGULAMA OLUSTUR---------------#
#----------------------------------------------#

Uygulama = QApplication(sys.argv)
penAna = QMainWindow()
ui = Ui_MainWindow()
ui.setupUi(penAna)
penAna.show()

#-------------VERİTABANI OLUŞTUR---------------#
#----------------------------------------------#

import sqlite3
global curs
global conn
conn = sqlite3.connect('mailveritabani.db')
curs = conn.cursor()

curs.execute("CREATE TABLE IF NOT EXISTS mailler(MailAdresi TEXT NOT NULL)")
conn.commit()

#-------------------KAYITLI MAİLLERİ LİSTELE---------------------#
#----------------------------------------------------------------#

def LISTELE():

    ui.tblMailler.clear()
    ui.tblMailler.setHorizontalHeaderLabels(('Mail Adresi',))
    ui.tblMailler.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

    curs.execute("SELECT * FROM mailler")
    for satirIndeks, satirVeri in enumerate(curs):
        for sutunIndeks, sutunVeri in enumerate(satirVeri):
            ui.tblMailler.setItem(satirIndeks, sutunIndeks, QTableWidgetItem(str(sutunVeri)))


LISTELE()

#-------------------BOŞLUKLARI TEMİZLE---------------------#
#----------------------------------------------------------#

def BOSLUKLARITEMIZLE():

    ui.lneKullaniciMailAdresi.clear()
    ui.lneKullaniciMailSifresi.clear()
    ui.lneMailinGidecegiAdres.clear()
    ui.lneMailKonusu.clear()
    ui.txteIleti.clear()

#-------------------KAYITLI MAİL BİLGİLERİNİ DOLDUR---------------------#
#-----------------------------------------------------------------------#

def BILGILERIDOLDUR():

    secili = ui.tblMailler.selectedItems()
    sayi = len(secili)
    mailler = ''
    for numara in range(0,sayi):

        bilgi = secili[numara].text()
        bilgi = str(bilgi)
        mailler += bilgi + ","

    ui.lneMailinGidecegiAdres.setText(mailler)


#-------------------MAİL ADRESİNİ KAYDET---------------------#
#------------------------------------------------------------#

def KAYDET():

    bilgi = ui.lneKaydedilecekMail.text()
    curs.execute("INSERT INTO mailler(MailAdresi) VALUES(?)",(bilgi,))
    conn.commit()
    LISTELE()
    ui.lneKaydedilecekMail.clear()

#-------------------***MAİLİ GÖNDERME***---------------------#
#------------------------------------------------------------#

def GONDER():

    liste = ui.lneMailinGidecegiAdres.text().split(",")
    adet = ui.spnbAdet.text()
    adet = int(adet)
    sayi = 0
    gosterge = 1
    while adet > 0:

        for i in range(0,len(liste)-1):
            sifre = ui.lneKullaniciMailSifresi.text()
            sifre = str(sifre)
            mesaj = MIMEMultipart()
            kullaniciMaili = ui.lneKullaniciMailAdresi.text()
            kullaniciMaili = str(kullaniciMaili)
            mesaj["From"] = kullaniciMaili
            gidecegiMail = liste[i]
            gidecegiMail = str(gidecegiMail)
            mesaj["To"] = gidecegiMail
            mailinKonusu = ui.lneMailKonusu.text()
            mailinKonusu = str(mailinKonusu)
            mesaj["Subject"] = mailinKonusu

            ileti = ui.txteIleti.toPlainText()
            ileti = str(ileti)
            yazi = ileti

            mesaj_govdesi = MIMEText(yazi, "plain")

            mesaj.attach(mesaj_govdesi)

            try:
                mail = smtplib.SMTP("smtp.gmail.com", 587)

                mail.ehlo()

                mail.starttls()

                mail.login(kullaniciMaili, sifre)

                mail.sendmail(mesaj["From"], mesaj["To"], mesaj.as_string())

                gosterge = str(gosterge)
                ui.statusbar.showMessage("Mail başarıyla gönderildi! ({gosterge})", 10000)
                mail.close()

            except Exception as Hata:
                Hata = str(Hata)
                ui.statusbar.showMessage("Mail Gönderilemedi. Hata: " + Hata, 10000)
        gosterge = int(gosterge)
        gosterge += 1
        adet -= 1
        sayi += 1

    sayi *= (len(liste)-1)
    sayi = str(sayi)
    ui.statusbar.showMessage(sayi+" adet mail başarıyla gönderildi.")

#-------------SİLME BUTONU---------------#
#----------------------------------------#

def SILME():
    cevap = QMessageBox.question(penAna, "MAİL SİL", "Kaydı silmek istediğinize emin misiniz?", \
                                 QMessageBox.Yes | QMessageBox.No)

    if cevap == QMessageBox.Yes:
        secili = ui.tblMailler.selectedItems()
        if len(secili) > 1:
            ui.statusbar.showMessage("Birden fazla maili aynı anda silemezsiniz!")
        else:
            silinecek = secili[0].text()
            try:
                curs.execute("DELETE FROM mailler WHERE MailAdresi = ?",(silinecek,))
                conn.commit()
                LISTELE()
                ui.statusbar.showMessage("MAİL SİLME İŞLEMİ BAŞARIYLA GERÇEKLEŞTİ...",10000)
            except Exception as Hata:
                Hata = str(Hata)
                ui.statusbar.showMessage("Şöyle bir hata ile karşılaşıldı: "+Hata,10000)
    else:
        ui.statusbar.showMessage("SİLME İŞLEMİ İPTAL EDİLDİ...", 10000)

#-------------ÇIKIŞ BUTONU---------------#
#----------------------------------------#

def CIKIS():

    cevap = QMessageBox.question(penAna, "ÇIKIŞ", "Programdan çıkmak istediğinize emin misiniz?", \
                                 QMessageBox.Yes | QMessageBox.No)
    if cevap == QMessageBox.Yes:
        conn.close()
        sys.exit(Uygulama.exec_())
    else:
        penAna.show()

#-------------SİNYAL-SLOT---------------#
#---------------------------------------#

ui.btnBosluklariTemizle.clicked.connect(lambda : BOSLUKLARITEMIZLE())
ui.btnBoslukDoldurma.clicked.connect(lambda : BILGILERIDOLDUR())
ui.btnAdresKaydet.clicked.connect(lambda : KAYDET())
ui.btnGonder.clicked.connect(lambda : GONDER())
ui.btnKayitliMailSil.clicked.connect(lambda : SILME())
ui.btnCikis.clicked.connect(lambda : CIKIS())

sys.exit(Uygulama.exec_())