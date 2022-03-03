from PyQt5 import uic

with open('AnaSayfaUI.py', 'w', encoding='utf-8') as fout:

    uic.compileUi('AnaSayfa.ui',fout)
