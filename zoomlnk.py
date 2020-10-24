from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.uic import loadUi
import os
import shutil
import tempfile
from hashlib import md5
import subprocess
import winshell
from win32com.client import Dispatch
from urllib.request import urlopen
import pyperclip
import re

desktop = winshell.desktop()
wDir = os.path.join(os.environ['APPDATA'], 'Zoom', 'bin')
icon = os.path.join(wDir, 'Zoom.exe')
shell = Dispatch('WScript.Shell')

##os.chdir(os.path.dirname(__file__))
__path__ = os.path.dirname(__file__)

app = QApplication([])
window = QWidget()
loadUi(os.path.join(__path__, 'zoomlnk.ui'), window)
window.setAcceptDrops(True) # allow drop on window
window.setWindowTitle('Zoom Shortcut Generator')
window.setWindowIcon(QIcon(os.path.join(__path__, 'zoom.png')))

def showmsgbox(icon, title, message, info, detail, buttons):
    msg = QMessageBox(window)
    msg.setIcon(icon)
    msg.setWindowTitle(title)
    msg.setText(message)
    msg.setInformativeText(info)
    msg.setDetailedText(detail)
    msg.setStandardButtons(buttons)
    msg.show()
    return msg.exec_()

def showinfo(title, message, detail = ''):
    return showmsgbox(QMessageBox.Information, title, message, '', detail,
                      QMessageBox.Ok)

def showerror(title, message, detail = ''):
    return showmsgbox(QMessageBox.Critical, title, message, '', detail,
                      QMessageBox.Ok)

def comboPressed():
    mode = window.modeInput.currentText()
    window.submitButton.setText(mode)
    if mode == 'Create':
        window.urlEnt.setEnabled(True)
    elif mode == 'Delete':
        window.urlEnt.setDisabled(True)

def submit():
    name = window.nameEnt.text()
    url = window.urlEnt.text()
    mode = window.modeInput.currentText().lower()
    
    if not name:
        showerror('Zoom Shortcut Generator', '"Name" field must be filled in')
        return
    if (not url) and mode == 'create':
        showerror('Zoom Shortcut Generator', '"URL" field must be filled in')
        return

    if mode == 'create':
        try:
            url = urlopen(url).url
        except:
            showerror('Zoom Shortcut Generator', '"URL" is invalid')
            return

        try:
            open(os.path.join(desktop, name), 'w').close()
            os.remove(os.path.join(desktop, name))
        except:
            showerror('Zoom Shortcut Generator', '"Name" is invalid')
            return
        
        url = urlopen(url).url
        
        fpath = os.path.join(desktop, md5(name.encode()).hexdigest())
        if not os.path.exists(fpath): os.mkdir(fpath)
        subprocess.call(['attrib', '+H', fpath], creationflags = 0x8000000)
        
        sbtmp = os.path.join(fpath, '.bat')
        svtmp = os.path.join(fpath, '.vbs')
        spath = os.path.join(desktop, f'{name}.lnk')
        
        sbfile = open(sbtmp, 'w')
        svfile = open(svtmp, 'w')
        slfile = shell.CreateShortCut(spath)

        sbfile.write(f'@echo off\n"%APPDATA%\\Zoom\\bin\\Zoom.exe" \
"--url={url}"')
        svfile.write(rf'''Set WshShell = CreateObject("WScript.Shell")
WshShell.Run chr(34) & "{sbtmp}" & Chr(34), 0
Set WshShell = Nothing''')

        slfile.Targetpath = svtmp
        slfile.WorkingDirectory = wDir
        slfile.IconLocation = icon

        sbfile.close()
        svfile.close()
        slfile.save()

        with open(os.path.join(fpath, '.txt'), 'w') as file: file.write(url)

        showinfo('Zoom Shortcut Generator', 'Zoom Shortcut Generated')
        
    elif mode == 'delete':
        path = os.path.join(desktop, md5(name.encode()).hexdigest())
        try:
            os.remove(os.path.join(desktop, f'{name}.lnk'))
        except:
            showerror('Zoom Shortcut Generator', '"Name" is invalid')
            return
            
        shutil.rmtree(path)

        showinfo('Zoom Shortcut Generator', 'Zoom Shortcut Deleted')

def submit2():
        try:
            with open(os.path.join(
                desktop,
                md5(window.nameEnt2.text().encode()).hexdigest(),
                '.txt'
                )) as file:
                url = file.read()
        except FileNotFoundError:
            showerror('Zoom Link Extractor', 'Zoom Shortcut Not Found')
            return
            
        pyperclip.copy(url)
        
        showinfo('Zoom Link Extractor',
                 f'Extracted Link: \n{url}\nCopied to clipboard')

def dragEnterEvent(e): # accept or ignore a drop request
    if e.mimeData().hasUrls():
        if len(e.mimeData().urls()) == 1:
            if e.mimeData().urls()[0].toLocalFile().endswith('.url'):
                e.accept()
                return
            elif os.path.exists(e.mimeData().urls()[0].toLocalFile()):
                e.ignore()
                return
        else:
            e.ignore()
            return

    if e.mimeData().text():
        e.accept()
        return
    e.ignore()

def dropEvent(e):
    if e.mimeData().hasUrls():
        filename = e.mimeData().urls()[0].toLocalFile()
        try:
            with open(filename) as file: data = file.read()
            urls = re.findall(r'URL=(.*)', data)
            if urls:
                url = urls[0]
            else:
                return
            window.nameEnt.setText(os.path.basename(filename)[:-4])
            window.urlEnt.setText(url)
        except UnicodeDecodeError:
            pass
        except (OSError, FileNotFoundError):
            window.urlEnt.setText(e.mimeData().text())
    else:
        window.urlEnt.setText(e.mimeData().text())

def showHelp():
    helpWindow = QDialog(window)
    loadUi(os.path.join(__path__, 'help.ui'), helpWindow)
    helpWindow.setWindowTitle('Help')
    helpWindow.show()
    helpWindow.raise_()

window.modeInput.activated.connect(comboPressed)
window.submitButton.clicked.connect(submit)
window.submitButton2.clicked.connect(submit2)
window.helpButton.clicked.connect(showHelp)
window.dragEnterEvent = dragEnterEvent
window.dropEvent = dropEvent

window.show()
window.raise_()
app.exec_()
