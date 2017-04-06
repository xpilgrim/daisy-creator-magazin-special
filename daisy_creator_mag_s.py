#!/usr/bin/python
# -*- coding: utf-8 -*-

"""
Autor: Joerg Sorge
Distributed under the terms of GNU GPL version 2 or later
Copyright (C) Joerg Sorge joergsorge at googel
2012-06-20

This program is for
- copy mp3 files for processing for DAISY Talking Books
- create DAISY Fileset
For now, it's only possible to create a flat DAISY-Structure

Dieses Programm
- kopiert mp3-Files fuer die Verarbeitung zu Daisy-Buechern
- erzeugt die noetigen Dateien fuer eine Daisy-Struktur.
Bisher kann nur eine DAISY-Ebene erzeugt werden.

Additional python modul necessary:
Zusatz-Modul benoetigt:
python-mutagen
lame
sudo apt-get install python-mutagen lame

Update gui with:
GUI aktualisieren mit:
pyuic4 daisy_creator_mag_s.ui -o daisy_creator_mag_s_ui.py
"""

from PyQt4 import QtGui, QtCore
import sys
import os
import shutil
import datetime
from datetime import timedelta
import ntpath
import subprocess
import string
import re
from mutagen.mp3 import MP3
from mutagen.id3 import ID3
from mutagen.id3 import ID3NoHeaderError
import ConfigParser
import daisy_creator_mag_s_ui


class DaisyCopy(QtGui.QMainWindow, daisy_creator_mag_s_ui.Ui_DaisyMain):
    """
    mainClass
    The second parent must be 'Ui_<obj. name of main widget class>'.
    """

    def __init__(self, parent=None):
        """Settings"""
        super(DaisyCopy, self).__init__(parent)
        # This is because Python does not automatically
        # call the parent's constructor.
        self.setupUi(self)
        # Pass this "self" for building widgets and
        # keeping a reference.
        self.app_debugMod = "yes"
        self.app_bhzItems = ["Zeitschrift"]
        self.app_prevAusgItems = ["10", "11", "12", "22", "23", "24", "25"]
        self.app_currentAusgItems = ["01", "02", "03", "04", "05", "06", "07",
            "08", "09", "10", "11", "12", "13", "14", "15", "16", "17", "18",
            "19", "20", "21", "22", "23", "24", "25"]
        self.app_nextAusgItems = ["01", "02", "03", "04"]
        self.app_bhzPath = QtCore.QDir.homePath()
        self.app_bhzPathDest = QtCore.QDir.homePath()
        self.app_bhzPathMeta = QtCore.QDir.homePath()
        self.app_bhzPathIssueAnnouncement = QtCore.QDir.homePath()
        self.app_bhzPathIntro = QtCore.QDir.homePath()
        # we need ext package lame
        self.app_lame = ""
        self.connectActions()

    def connectActions(self):
        """Define Actions """
        self.toolButtonCopySource.clicked.connect(self.actionOpenCopySource)
        self.toolButtonCopyDest.clicked.connect(self.actionOpenCopyDest)
        self.toolButtonCopyFile1.clicked.connect(self.actionOpenCounterFile)
        self.toolButtonMetaFile.clicked.connect(self.actionOpenMetaFile)
        self.commandLinkButton.clicked.connect(self.actionRunCopy)
        self.commandLinkButtonMeta.clicked.connect(self.metaLoadFile)
        self.commandLinkButtonDaisy.clicked.connect(self.actionRunDaisy)
        self.toolButtonDaisySource.clicked.connect(self.actionOpenDaisySource)
        self.pushButtonClose1.clicked.connect(self.actionQuit)
        self.pushButtonClose2.clicked.connect(self.actionQuit)
        self.pushButtonClose3.clicked.connect(self.actionQuit)

    def readConfig(self):
        """read Config from file"""
        fileExist = os.path.isfile("daisy_creator_mag_s.config")
        if fileExist is False:
            self.showDebugMessage(u"File not exists")
            self.textEdit.append(
                "<font color='red'>"
                + "Config-Datei konnte nicht geladen werden: </font>"
                + "daisy_creator_mag.config")
            return

        config = ConfigParser.RawConfigParser()
        config.read("daisy_creator_mag_s.config")
        self.app_bhzPath = config.get('Ordner', 'BHZ')
        self.app_bhzPathDest = config.get('Ordner', 'BHZ-Ziel')
        self.app_bhzPathMeta = config.get('Ordner', 'BHZ-Meta')
        self.app_bhzPathIssueAnnouncement = config.get('Ordner',
                                                    'BHZ-Ausgabeansage')
        self.app_bhzPathIntro = config.get('Ordner', 'BHZ-Intro')
        self.app_bhzItems = config.get('Blindenhoerzeitschriften',
                                                        'BHZ').split(",")
        self.app_lame = config.get('Programme', 'LAME')

    def readHelp(self):
        """read Readme from file"""
        fileExist = os.path.isfile("README.md")
        if fileExist is False:
            self.showDebugMessage("File not exists")
            self.textEdit.append(
                "<font color='red'>"
                + "Hilfe-Datei konnte nicht geladen werden: </font>"
                + "README.md")
            return

        fobj = open("README.md")
        for line in fobj:
            self.textEditHelp.append(line)
        # set cursor on top of helpfile
        cursor = self.textEditHelp.textCursor()
        cursor.movePosition(QtGui.QTextCursor.Start,
                            QtGui.QTextCursor.MoveAnchor, 0)
        self.textEditHelp.setTextCursor(cursor)
        fobj.close()

    def actionOpenCopySource(self):
        """Source of copy"""
        # QtCore.QDir.homePath()
        dirSource = QtGui.QFileDialog.getExistingDirectory(
                        self,
                        "Quell-Ordner",
                        self.app_bhzPath
                    )
        # Don't attempt to open if open dialog
        # was cancelled away.
        if dirSource:
            self.lineEditCopySource.setText(dirSource)
            self.textEdit.append("Quelle:")
            self.textEdit.append(dirSource)

    def actionOpenDaisySource(self):
        """Source of daisy"""
        dirSource = QtGui.QFileDialog.getExistingDirectory(
                        self,
                        "Quell-Ordner",
                        self.app_bhzPathDest
                    )
        # Don't attempt to open if open dialog
        # was cancelled away.
        if dirSource:
            self.lineEditDaisySource.setText(dirSource)
            self.textEdit.append("Quelle:")
            self.textEdit.append(dirSource)

    def actionOpenCopyDest(self):
        """Destination of Copy"""
        dirDest = QtGui.QFileDialog.getExistingDirectory(
                        self,
                        "Ziel-Ordner",
                        self.app_bhzPath
                    )
        # Don't attempt to open if open dialog
        # was cancelled away.
        if dirDest:
            self.lineEditCopyDest.setText(dirDest)
            self.textEdit.append("Ziel:")
            self.textEdit.append(dirDest)

    def actionOpenCounterFile(self):
        """counter file for renaming of files"""
        file1 = QtGui.QFileDialog.getOpenFileName(
                        self,
                        "Counter-Datei",
                        self.app_bhzPath
                    )
        # Don't attempt to open if open dialog
        # was cancelled away.
        if file1:
            self.lineEditFileCounter.setText(file1)
            self.textEdit.append("Counter-Datei:")
            self.textEdit.append(file1)

    def actionOpenMetaFile(self):
        """Metafile to load"""
        mfile = QtGui.QFileDialog.getOpenFileName(
                        self,
                        "Daisy_Meta",
                        self.app_bhzPathMeta
                    )
        # Don't attempt to open if open dialog
        # was cancelled away.
        if mfile:
            self.lineEditMetaSource.setText(mfile)

    def actionRunCopy(self):
        """main function for copy"""
        if self.lineEditCopySource.text() == "Quell-Ordner":
            errorMessage = u"Quell-Ordner wurde nicht ausgewaehlt.."
            self.showDialogCritical(errorMessage)
            return

        if self.lineEditCopyDest.text() == "Ziel-Ordner":
            errorMessage = u"Ziel-Ordner wurde nicht ausgewaehlt.."
            self.showDialogCritical(errorMessage)
            return

        if self.lineEditFileCounter.text() == "Counter-Datei waehlen":
            if self.comboBoxCopyBhz.currentText() == "ABC_Journal":
                errorMessage = u"Counter-Datei wurde nicht ausgewaehlt.."
                self.showDialogCritical(errorMessage)
                return
        else:
            counterLines = self.readCounterFile()
            nCounterLines = len(counterLines)

        self.showDebugMessage(self.lineEditCopySource.text())
        self.showDebugMessage(self.lineEditCopyDest.text())

        # check for files in source
        try:
            dirsSource = os.listdir(self.lineEditCopySource.text())
        except Exception, e:
            logMessage = u"read_files_from_dir Error: %s" % str(e)
            self.showDebugMessage(logMessage)

        # ceck dir of dest
        if os.path.exists(self.lineEditCopyDest.text()) is False:
            errorMessage = u"Ziel-Ordner existiert nicht.."
            self.showDebugMessage(errorMessage)
            self.showDialogCritical(errorMessage)
            self.lineEditCopyDest.setFocus()
            return

        self.showDebugMessage(dirsSource)
        # add number to counter if copy AusgAnsage and/or Intro
        zCounter = 0
        if self.checkBoxCopyBhzIntro.isChecked():
            zCounter += 1
        if self.checkBoxCopyBhzIssueAnnouncement.isChecked():
            zCounter += 1

        z = 0
        nDirsSource = len(dirsSource)
        if nCounterLines != nDirsSource:
            self.textEdit.append(
                    "<b><font color='red'>"
                    + "Anzahl der mp3-Dateien (" + str(nDirsSource)
                    + ") stimmt nicht mit Counter-Datei ("
                    + str(nCounterLines)
                    + ") ueberein!</font></b>")

        filesOK = self.compareFiles(dirsSource, counterLines)
        if filesOK is None:
            return

        self.textEdit.append("<b>Kopieren:</b>")
        self.showDebugMessage(nDirsSource)
        dirsSource.sort()
        # loop trough files
        for item in dirsSource:
            if (item[len(item) - 4:len(item)] != ".MP3"
                            and item[len(item) - 4:len(item)] != ".mp3"):
                continue

            fileToCopySource = self.lineEditCopySource.text() + "/" + item
            # check if file exist
            fileExist = os.path.isfile(fileToCopySource)
            if fileExist is False:
                self.showDebugMessage("File not exists")
                # change max number and update progress
                nDirsSource = nDirsSource - 1
                pZ = z * 100 / nDirsSource
                self.progressBarCopy.setValue(pZ)
                self.textEdit.append(
                        "<b>Datei konnte nicht kopiert werden: </b>"
                        + fileToCopySource)
                self.showDebugMessage(fileToCopySource)
                self.textEdit.append("<b>Uebersprungen</b>:")
                self.textEdit.append(fileToCopySource)
                continue

            # search filename in counterfile,
            # extract counter and set it on top of the filename
            nFound = 0
            for line in counterLines:
                nLenLine = len(line) - 1
                if line[3:nLenLine] == item:
                    nFound += 1
                    self.showDebugMessage("counterline " + line)
                    self.showDebugMessage(u"counter " + line[0:02])
                    # Counter on top
                    zFileCount = zCounter + int(line[0:02])
                    fileToCopyDest = (self.lineEditCopyDest.text() + "/"
                                + str(zFileCount).zfill(2) + "_" + item)

            # filename not found
            if nFound == 0:
                self.textEdit.append(
                    "<b><font color='red'>"
                    + "Dateiname in Counterdatei nicht gefunden</font></b>:")
                self.textEdit.append(ntpath.basename(str(fileToCopySource)))
                self.textEdit.append("<b>Bearbeitung abgebrochen</b>:")
                return

            self.textEdit.append(ntpath.basename(str(fileToCopyDest)))
            self.showDebugMessage(fileToCopySource)
            self.showDebugMessage(fileToCopyDest)

            # check Bitrate, when necessary recode in new destination
            isChangedAndCopy = self.checkChangeBitrateAndCopy(
                                        fileToCopySource, fileToCopyDest)
            # nothing to do, only copy
            if isChangedAndCopy is None:
                self.copyFile(fileToCopySource, fileToCopyDest)

            self.checkCangeId3(fileToCopyDest)
            z += 1
            self.showDebugMessage(z)
            self.showDebugMessage(nDirsSource)
            pZ = z * 100 / nDirsSource
            self.showDebugMessage(pZ)
            self.progressBarCopy.setValue(pZ)

        self.showDebugMessage(z)

        if self.checkBoxCopyBhzIntro.isChecked():
            self.copyIntro()

        if self.checkBoxCopyBhzIssueAnnouncement.isChecked():
            self.copyIssueAnnouncement()

        # load metadata
        self.lineEditMetaSource.setText(
            self.app_bhzPathMeta + "/Daisy_Meta_"
            + self.comboBoxCopyBhz.currentText())
        self.metaLoadFile()
        # enter path of source and destination
        self.lineEditDaisySource.setText(self.lineEditCopyDest.text())

    def checkPackages(self, package):
        """
        check if package is installed
        needs subprocess, os
        http://stackoverflow.com/
        questions/11210104/check-if-a-program-exists-from-a-python-script
        """
        try:
            devnull = open(os.devnull, "w")
            subprocess.Popen([package], stdout=devnull,
                            stderr=devnull).communicate()
        except OSError as e:
            if e.errno == os.errno.ENOENT:
                errorMessage = (u"Es fehlt das Paket:\n " + package
                                + u"\nZur Nutzung des vollen Funktionsumfanges "
                                + "muss es installiert werden!")
                self.showDialogCritical(errorMessage)
                self.textEdit.append(
                    "<font color='red'>Es fehlt das Paket: </font> " + package)

    def checkFilename(self, fileName):
        """check for spaces and non ascii characters"""
        error = None
        self.textEdit.append(fileName)
        if type(fileName) is str:
            try:
                cfileName = fileName.encode("ascii")
            except Exception, e:
                error = str(e)
        else:
            # maybe fileName could be QString, so we must convert it
            try:
                cfileName = str(fileName)
                self.textEdit.append(cfileName)
            except Exception, e:
                error = str(e)

        if error is not None:
            if (error.find("'ascii' codec can't encode character") != -1
                or
                error.find("'ascii' codec can't decode byte") != -1):
                errorMessage = ("<b>Unerlaubte(s) Zeichen im Dateinamen!</b>")
                self.showDebugMessage(errorMessage)
                self.textEdit.append(errorMessage)
                return None
            else:
                errorMessage = ("<b>Fehler im Dateinamen!</b>")
                self.showDebugMessage(errorMessage)
                self.textEdit.append(errorMessage)
                return None

        if cfileName.find(" ") != -1:
            errorMessage = ("<b>Unerlaubtes Leerzeichen im Dateinamen!</b>")
            self.textEdit.append(errorMessage)
            self.tabWidget.setCurrentIndex(1)
            return None
        return "OK"

    def checkFilenames(self, filesSource):
        for item in filesSource:
            if (item[len(item) - 4:len(item)] != ".MP3"
                                and item[len(item) - 4:len(item)] != ".mp3"):
                continue
            checkOK = self.checkFilename(item)
            if checkOK is None:
                return None
        return "OK"

    def copyFile(self, fileToCopySource, fileToCopyDest):
        """copy file"""
        try:
            shutil.copy(fileToCopySource, fileToCopyDest)
        except Exception, e:
            logMessage = u"copy_file Error: %s" % str(e)
            self.showDebugMessage(logMessage)
            self.textEdit.append(logMessage + fileToCopyDest)

    def copyIntro(self):
        """copy intro"""
        fileToCopySource = (self.app_bhzPathIntro + "/Intro_"
                            + self.comboBoxCopyBhz.currentText() + ".mp3")
        fileToCopyDest = (self.lineEditCopyDest.text() + "/02_"
                    + self.comboBoxCopyBhz.currentText() + "_-_"
                    + self.comboBoxCopyBhzAusg.currentText() + "_Intro.mp3")
        self.showDebugMessage(fileToCopySource)
        self.showDebugMessage(fileToCopyDest)

        fileExist = os.path.isfile(fileToCopySource)
        if fileExist is False:
            self.showDebugMessage("File not exists")
            self.textEdit.append(
                "<font color='red'>"
                + "Intro nicht vorhanden</font>: "
                + os.path.basename(str(fileToCopySource)))
            return

        self.copyFile(fileToCopySource, fileToCopyDest)
        self.checkCangeId3(fileToCopyDest)

    def copyIssueAnnouncement(self):
        """copy announcement of issue"""
        pfadAusgabe = (self.app_bhzPathIssueAnnouncement + "_"
                    + self.comboBoxCopyBhzAusg.currentText()[0:4] + "_"
                    + self.comboBoxCopyBhz.currentText())
        self.showDebugMessage(pfadAusgabe)
        fileToCopySource = (pfadAusgabe + "/"
                + self.comboBoxCopyBhz.currentText() + "_-_Ausgabe_"
                + self.comboBoxCopyBhzAusg.currentText() + ".mp3")
        fileToCopyDest = (self.lineEditCopyDest.text() + "/01_"
                + self.comboBoxCopyBhz.currentText() + "_-_Ausgabe_"
                + self.comboBoxCopyBhzAusg.currentText() + ".mp3")
        self.showDebugMessage(fileToCopySource)
        self.showDebugMessage(fileToCopyDest)

        fileExist = os.path.isfile(fileToCopySource)
        if fileExist is False:
            self.showDebugMessage("File not exists")
            self.textEdit.append(
                "<font color='red'>"
                + "Ausgabeansage nicht vorhanden</font>: "
                + os.path.basename(str(fileToCopySource)))
            return

        self.copyFile(fileToCopySource, fileToCopyDest)
        self.checkCangeId3(fileToCopyDest)

    def readCounterFile(self):
        """reaf counter file with filenames and counter """
        try:
            f = open(str(self.lineEditFileCounter.text()), 'r')
            lines = f.readlines()
            f.close()
            self.textEdit.append("<b>Counter-Datei gelesen: </b>"
                                    + self.lineEditFileCounter.text())
            #for item in lines:
            #    self.textEdit.append("line: " + item)
        except IOError as (errno, strerror):
            lines = None
            self.showDebugMessage(
                "read_from_file_lines_in_list: nicht lesbar: "
                + self.lineEditFileCounter.text())
        return lines

    def checkCangeId3(self, fileToCopyDest):
        """check id3 tags, remove it when present"""
        tag = None
        try:
            audio = ID3(fileToCopyDest)
            tag = "yes"
        except ID3NoHeaderError:
            self.showDebugMessage(u"No ID3 header found; skipping.")

        if tag is not None:
            if self.checkBoxCopyID3Change.isChecked():
                audio.delete()
                self.textEdit.append("<b>ID3 entfernt bei</b>: "
                                + ntpath.basename(str(fileToCopyDest)))
                self.showDebugMessage(u"ID3 entfernt bei " + fileToCopyDest)
            else:
                self.textEdit.append(
                    "<b>ID3 vorhanden, aber NICHT entfernt bei</b>: "
                    + fileToCopyDest)

    def checkChangeBitrateAndCopy(self, fileToCopySource, fileToCopyDest):
        """change bitrate and encode in destination"""
        isChangedAndCopy = None
        audioSource = MP3(fileToCopySource)
        if (audioSource.info.bitrate ==
            int(self.comboBoxPrefBitrate.currentText()) * 1000):
            return isChangedAndCopy

        isEncoded = None
        self.textEdit.append(u"Bitrate Vorgabe: "
                            + str(self.comboBoxPrefBitrate.currentText()))
        self.textEdit.append(
            u"<b>Bitrate folgender Datei entspricht nicht der Vorgabe:</b> "
            + str(audioSource.info.bitrate / 1000) + " " + fileToCopySource)

        if self.checkBoxCopyBitrateChange.isChecked():
            self.textEdit.append(u"<b>Bitrate aendern bei</b>: "
                                                        + fileToCopyDest)
            isEncoded = self.encodeFile(fileToCopySource, fileToCopyDest)
            if isEncoded is not None:
                self.textEdit.append(u"<b>Bitrate geaendert bei</b>: "
                                                        + fileToCopyDest)
                isChangedAndCopy = True
        else:
            self.textEdit.append(
                u"<b>Bitrate wurde NICHT geaendern bei</b>: "
                + fileToCopyDest)
        return isChangedAndCopy

    def compareFiles(self, dirsSource, counterLines):
        """compare files from source with conter file"""
        self.showDebugMessage(counterLines)
        self.textEdit.append("<b>Dateinamen pruefen...</b>")
        checkOK = self.checkFilenames(dirsSource)
        if checkOK is None:
            return

        self.textEdit.append("<b>Dateien pruefen...</b>")
        # loop trough files
        for item in dirsSource:
            if (item[len(item) - 4:len(item)] != ".MP3"
                            and item[len(item) - 4:len(item)] != ".mp3"):
                continue

            nFound = 0
            for line in counterLines:
                # remove line break
                nLenLine = len(line) - 1
                if line[3:nLenLine] == item:
                    nFound += 1

            if nFound == 0:
                self.textEdit.append(
                "<b><font color='red'>"
                + "Fuer diese mp3-Datei keine Entsprechung in "
                + "Counterdatei gefunden"
                + "</font></b>:")
                self.textEdit.append(item)
                self.textEdit.append("<b>Bearbeitung abgebrochen!</b>")
                return None

        # now let's take a look from the other site
        for line in counterLines:
            #self.textEdit.append("counter: " + line[3:nLenLine])
            # remove line break
            nLenLine = len(line) - 1
            nFound = 0
            for item in dirsSource:
                #self.textEdit.append(item)
                if (item[len(item) - 4:len(item)] != ".MP3"
                            and item[len(item) - 4:len(item)] != ".mp3"):
                    continue
                if line[3:nLenLine] == item:
                    nFound += 1
                    #self.textEdit.append("found: " +line[3:nLenLine])

            if nFound == 0:
                self.textEdit.append(
                "<b><font color='red'>"
                + "Keine mp3-Datei zu diesem Eintrag in "
                + "Counterdatei gefunden</font></b>:")
                self.textEdit.append(line[3:nLenLine])
                self.textEdit.append("<b>Bearbeitung abgebrochen!</b>")
                return None
        return "OK"

    def encodeFile(self, fileToCopySource, fileToCopyDest):
        """encode mp3-files"""
        self.showDebugMessage(u"encode_file")
        #damit die uebergabe der befehle richtig klappt,
        # muessen alle cmds im richtigen zeichensatz als strings encoded sein
        c_lame_encoder = "/usr/bin/lame"
        self.showDebugMessage(u"type c_lame_encoder")
        self.showDebugMessage(type(c_lame_encoder))
        self.showDebugMessage(u"fileToCopySource")
        self.showDebugMessage(type(fileToCopySource))
        self.showDebugMessage(fileToCopyDest)
        self.showDebugMessage(u"type(fileToCopyDest)")
        self.showDebugMessage(type(fileToCopyDest))

        p = subprocess.Popen([c_lame_encoder, "-b",
                self.comboBoxPrefBitrate.currentText(),
                "-m", "m", fileToCopySource, fileToCopyDest],
                stdout=subprocess.PIPE, stderr=subprocess.PIPE).communicate()

        self.showDebugMessage(u"returncode 0")
        self.showDebugMessage(p[0])
        self.showDebugMessage(u"returncode 1")
        self.showDebugMessage(p[1])

        # erfolgsmeldung suchen, wenn nicht gefunden: -1
        n_encode_percent = string.find(p[1], "(100%)")
        n_encode_percent_1 = string.find(p[1], "(99%)")
        self.showDebugMessage(n_encode_percent)
        c_complete = "no"

        # bei kurzen files kommt die 100% meldung nicht,
        # deshalb auch 99% durchgehen lassen
        if n_encode_percent == -1:
            # 100% nicht erreicht
            if n_encode_percent_1 != -1:
                # aber 99
                c_complete = "yes"
        else:
            c_complete = "yes"

        if c_complete == "yes":
            log_message = u"recoded_file: " + fileToCopySource
            self.showDebugMessage(log_message)
            return fileToCopyDest
        else:
            log_message = u"recode_file Error: " + fileToCopySource
            self.showDebugMessage(log_message)
            return None

    def metaLoadFile(self):
        """load meta file"""
        fileExist = os.path.isfile(self.lineEditMetaSource.text())
        if fileExist is False:
            self.showDebugMessage("File not exists")
            self.textEdit.append(
                "<font color='red'>"
                + "Meta-Datei konnte nicht geladen werden</font>: "
                + os.path.basename(str(self.lineEditMetaSource.text())))
            return

        config = ConfigParser.RawConfigParser()
        # Pfad von QTString in String  umwandeln
        config.read(str(self.lineEditMetaSource.text()))
        self.lineEditMetaProducer.setText(config.get('Daisy_Meta', 'Produzent'))
        self.lineEditMetaAutor.setText(config.get('Daisy_Meta', 'Autor'))
        self.lineEditMetaTitle.setText(config.get('Daisy_Meta', 'Titel'))
        self.lineEditMetaEdition.setText(config.get('Daisy_Meta', 'Edition'))
        self.lineEditMetaNarrator.setText(config.get('Daisy_Meta', 'Sprecher'))
        self.lineEditMetaKeywords.setText(config.get('Daisy_Meta',
                                                        'Stichworte'))
        self.lineEditMetaRefOrig.setText(config.get('Daisy_Meta',
                                            'ISBN/Ref-Nr.Original'))
        self.lineEditMetaPublisher.setText(config.get('Daisy_Meta', 'Verlag'))
        self.lineEditMetaYear.setText(config.get('Daisy_Meta', 'Jahr'))

    def actionRunDaisy(self):
        """create Daisy-Fileset"""
        if self.lineEditDaisySource.text() == "Quell-Ordner":
            errorMessage = u"Quell-Ordner wurde nicht ausgewaehlt.."
            self.showDialogCritical(errorMessage)
            return

        # Audios einlesen
        try:
            dirItems = os.listdir(self.lineEditDaisySource.text())
        except Exception, e:
            logMessage = u"read_files_from_dir Error: %s" % str(e)
            self.showDebugMessage(logMessage)

        self.progressBarDaisy.setValue(10)
        self.showDebugMessage(dirItems)
        self.textEditDaisy.append(u"<b>Folgende Audios werden bearbeitet:</b>")
        zMp3 = 0
        zList = len(dirItems)
        self.showDebugMessage(zList)
        dirAudios = []
        dirItems.sort()
        for item in dirItems:
            if (item[len(item) - 4:len(item)]
                == ".MP3" or item[len(item) - 4:len(item)] == ".mp3"):
                dirAudios.append(item)
                self.textEditDaisy.append(item)
                zMp3 += 1

        #totalAudioLength = self.calcAudioLengt(dirAudios)
        lTimes = self.calcAudioLengt(dirAudios)
        totalAudioLength = lTimes[0]
        lTotalElapsedTime = lTimes[1]
        lFileTime = lTimes[2]
        print totalAudioLength
        totalTime = timedelta(seconds=totalAudioLength)
        # umwandlung von timedelta in string:
        # minuten und sekunden musten immer zweistllig sein,
        # damit einstellige stunde eine null bekommt :zfill(8)
        lTotalTime = str(totalTime).split(".")
        cTotalTime = lTotalTime[0].zfill(8)
        #str(cTotalTime[0]).zfill(8)
        self.textEditDaisy.append(u"Gesamtlaenge: " + cTotalTime)
        self.writeNCC(cTotalTime, zMp3, dirAudios)
        self.progressBarDaisy.setValue(20)
        self.writeMasterSmil(cTotalTime, dirAudios)
        self.progressBarDaisy.setValue(50)
        self.writeSmil(lTotalElapsedTime, lFileTime, dirAudios)
        self.progressBarDaisy.setValue(100)

    def calcAudioLengt(self, dirAudios):
        """calc length of all audios"""
        totalAudioLength = 0
        lTotalElapsedTime = []
        lTotalElapsedTime.append(0)
        lFileTime = []
        for item in dirAudios:
            fileToCheck = os.path.join(str(
                            self.lineEditDaisySource.text()), item)
            audioSource = MP3(fileToCheck)
            self.showDebugMessage(item + " " + str(audioSource.info.length))
            totalAudioLength += audioSource.info.length
            lTotalElapsedTime.append(totalAudioLength)
            lFileTime.append(audioSource.info.length)
            lTimes = []
            lTimes.append(totalAudioLength)
            lTimes.append(lTotalElapsedTime)
            lTimes.append(lFileTime)
        return lTimes

    def writeNCC(self, cTotalTime, zMp3, dirAudios):
        """write NCC-Page"""
        try:
            fOutFile = open(os.path.join(
                str(self.lineEditDaisySource.text()), "ncc.html"), 'w')
        except IOError as (errno, strerror):
            self.showDebugMessage("I/O error({0}): {1}".format(errno, strerror))
            return
        self.textEditDaisy.append(u"<b>NCC-Datei schreiben</b>")
        fOutFile.write('<?xml version="1.0" encoding="utf-8"?>' + '\r\n')
        fOutFile.write('<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0'
            + ' Transitional//EN"'
            + ' "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">'
            + '\r\n')
        fOutFile.write('<html xmlns="http://www.w3.org/1999/xhtml">' + '\r\n')
        fOutFile.write('<head>' + '\r\n')
        fOutFile.write('<meta http-equiv="Content-type" '
            + 'content="text/html; charset=utf-8"/>' + '\r\n')
        fOutFile.write('<title>' + self.comboBoxCopyBhz.currentText()
                       + '</title>' + '\r\n')

        fOutFile.write('<meta name="ncc:generator" '
            + 'content="KOM-IN-DaisyCreator"/>' + '\r\n')
        fOutFile.write('<meta name="ncc:revision" content="1"/>' + '\r\n')

        today = datetime.date.today()
        fOutFile.write('<meta name="ncc:producedDate" content="'
            + today.strftime("%Y-%m-%d") + '"/>' + '\r\n')
        fOutFile.write('<meta name="ncc:revisionDate" content="'
            + today.strftime("%Y-%m-%d") + '"/>' + '\r\n')
        fOutFile.write('<meta name="ncc:tocItems" content="'
            + str(zMp3) + '"/>' + '\r\n')

        fOutFile.write('<meta name="ncc:totalTime" content="'
            + cTotalTime + '"/>' + '\r\n')
        fOutFile.write('<meta name="ncc:narrator" content="'
            + self.lineEditMetaNarrator.text() + '"/>' + '\r\n')
        fOutFile.write('<meta name="ncc:pageNormal" content="0"/>' + '\r\n')
        fOutFile.write('<meta name="ncc:pageFront" content="0"/>' + '\r\n')
        fOutFile.write('<meta name="ncc:pageSpecial" content="0"/>' + '\r\n')
        fOutFile.write('<meta name="ncc:sidebars" content="0"/>' + '\r\n')
        fOutFile.write('<meta name="ncc:prodNotes" content="0"/>' + '\r\n')
        fOutFile.write('<meta name="ncc:footnotes" content="0"/>' + '\r\n')
        fOutFile.write('<meta name="ncc:depth" content="'
            + str(self.spinBoxLevel.value()) + '"/>' + '\r\n')
        fOutFile.write('<meta name="ncc:maxPageNormal" content="'
            + str(self.spinBoxPages.value()) + '"/>' + '\r\n')
        fOutFile.write('<meta name="ncc:charset" content="utf-8"/>' + '\r\n')
        fOutFile.write('<meta name="ncc:multimediaType" content="audioNcc"/>'
            + '\r\n')
        #fOutFile.write( '<meta name="ncc:kByteSize" content=" "/>'+ '\r\n')
        fOutFile.write('<meta name="ncc:setInfo" content="1 of 1"/>' + '\r\n')

        fOutFile.write('<meta name="ncc:sourceDate" content="'
            + self.lineEditMetaYear.text() + '"/>' + '\r\n')
        fOutFile.write('<meta name="ncc:sourceEdition" content="'
            + self.lineEditMetaEdition.text() + '"/>' + '\r\n')
        fOutFile.write('<meta name="ncc:sourcePublisher" content="'
            + self.lineEditMetaPublisher.text() + '"/>' + '\r\n')

        #Anzahl files = Records 2x + ncc.html + master.smil
        fOutFile.write('<meta name="ncc:files" content="'
            + str(zMp3 + zMp3 + 2) + '"/>' + '\r\n')
        fOutFile.write('<meta name="ncc:producer" content="'
            + self.lineEditMetaProducer.text() + '"/>' + '\r\n')

        fOutFile.write('<meta name="dc:creator" content="'
            + self.lineEditMetaAutor.text() + '"/>' + '\r\n')
        fOutFile.write('<meta name="dc:date" content="'
            + today.strftime("%Y-%m-%d") + '"/>' + '\r\n')
        fOutFile.write('<meta name="dc:format" content="Daisy 2.02"/>' + '\r\n')
        fOutFile.write('<meta name="dc:identifier" content="'
            + self.lineEditMetaRefOrig.text() + '"/>' + '\r\n')
        fOutFile.write('<meta name="dc:language" content="de"'
                        + ' scheme="ISO 639"/>' + '\r\n')
        fOutFile.write('<meta name="dc:publisher" content="'
            + self.lineEditMetaPublisher.text() + '"/>' + '\r\n')
        fOutFile.write('<meta name="dc:source" content="'
            + self.lineEditMetaRefOrig.text() + '"/>' + '\r\n')
        fOutFile.write('<meta name="dc:subject" content="'
            + self.lineEditMetaKeywords.text() + '"/>' + '\r\n')
        fOutFile.write('<meta name="dc:title" content="'
            + self.lineEditMetaTitle.text() + '"/>' + '\r\n')
        # Medibus-OK items
        fOutFile.write('<meta name="prod:audioformat" content="wave 44 kHz"/>'
                       + '\r\n')
        fOutFile.write('<meta name="prod:compression" content="mp3 '
            + self.comboBoxPrefBitrate.currentText() + '/ kb/s"/>' + '\r\n')
        fOutFile.write('<meta name="prod:localID" content=" "/>' + '\r\n')
        fOutFile.write('</head>' + '\r\n')
        fOutFile.write('<body>' + '\r\n')
        z = 0
        for item in dirAudios:
            z += 1
            if z == 1:
                fOutFile.write('<h1 class="title" id="cnt_0001">'
                + '<a href="0001.smil#txt_0001">'
                + self.lineEditMetaAutor.text() + ": "
                + self.lineEditMetaTitle.text() + '</a></h1>' + '\r\n')
                continue
            # trennen
            itemSplit = self.splitFilename(item)
            cAuthor = self.extractAuthor(itemSplit)
            cTitle = self.extractTitle(itemSplit)
            fOutFile.write('<h' + str(self.spinBoxLevel.value())
                + ' id="cnt_' + str(z).zfill(4) + '"><a href="'
                + str(z).zfill(4) + '.smil#txt_' + str(z).zfill(4) + '">'
                + cAuthor + " - " + cTitle + '</a></h1>' + '\r\n')

        fOutFile.write("</body>" + '\r\n')
        fOutFile.write("</html>" + '\r\n')
        fOutFile.close

    def writeMasterSmil(self, cTotalTime, dirAudios):
        """write MasterSmil-Page """
        try:
            fOutFile = open(
                os.path.join(
                str(self.lineEditDaisySource.text()), "master.smil"), 'w')
        except IOError as (errno, strerror):
            self.showDebugMessage("I/O error({0}): {1}".format(errno, strerror))
            return
        self.textEditDaisy.append(u"<b>MasterSmil-Datei schreiben</b>")
        fOutFile.write('<?xml version="1.0" encoding="utf-8"?>' + '\r\n')
        fOutFile.write('<!DOCTYPE smil PUBLIC "-//W3C//DTD SMIL 1.0//EN"'
            + ' "http://www.w3.org/TR/REC-smil/SMIL10.dtd">' + '\r\n')
        fOutFile.write('<smil>' + '\r\n')
        fOutFile.write('<head>' + '\r\n')
        fOutFile.write('<meta name="dc:format" content="Daisy 2.02"/>' + '\r\n')
        fOutFile.write('<meta name="dc:identifier" content="'
            + self.lineEditMetaRefOrig.text() + '"/>' + '\r\n')
        fOutFile.write('<meta name="dc:title" content="'
            + self.lineEditMetaTitle.text() + '"/>' + '\r\n')
        fOutFile.write('<meta name="ncc:generator" '
            + 'content="KOM-IN-DaisyCreator"/>' + '\r\n')
        fOutFile.write('<meta name="ncc:format" content="Daisy 2.0"/>' + '\r\n')
        fOutFile.write('<meta name="ncc:timeInThisSmil" content="'
            + cTotalTime + '" />' + '\r\n')

        fOutFile.write('<layout>' + '\r\n')
        fOutFile.write('<region id="txt-view" />' + '\r\n')
        fOutFile.write('</layout>' + '\r\n')
        fOutFile.write('</head>' + '\r\n')
        fOutFile.write('<body>' + '\r\n')

        z = 0
        for item in dirAudios:
            z += 1
            # trennen
            itemSplit = self.splitFilename(item)
            cAuthor = self.extractAuthor(itemSplit)
            cTitle = self.extractTitle(itemSplit)
            fOutFile.write(
                '<ref src="' + str(z).zfill(4) + '.smil" title="'
                + cAuthor + " - " + cTitle + '" id="smil_'
                + str(z).zfill(4) + '"/>' + '\r\n')

        fOutFile.write('</body>' + '\r\n')
        fOutFile.write('</smil>' + '\r\n')
        fOutFile.close

    def writeSmil(self, lTotalElapsedTime, lFileTime, dirAudios):
        """write Smil-Pages"""
        z = 0
        for item in dirAudios:
            z += 1

            try:
                filename = str(z).zfill(4) + '.smil'
                fOutFile = open(
                    os.path.join(str(self.lineEditDaisySource.text()),
                    filename), 'w')
            except IOError as (errno, strerror):
                self.showDebugMessage(
                            "I/O error({0}): {1}".format(errno, strerror))
                return
            self.textEditDaisy.append(str(z).zfill(4)
                + u".smil - File schreiben")
            # split
            itemSplit = self.splitFilename(item)
            #cAuthor = self.extractAuthor(itemSplit)
            cTitle = self.extractTitle(itemSplit)

            fOutFile.write('<?xml version="1.0" encoding="utf-8"?>' + '\r\n')
            fOutFile.write('<!DOCTYPE smil PUBLIC "-//W3C//DTD SMIL 1.0//EN"'
                + ' "http://www.w3.org/TR/REC-smil/SMIL10.dtd">' + '\r\n')
            fOutFile.write('<smil>' + '\r\n')
            fOutFile.write('<head>' + '\r\n')
            fOutFile.write(
                '<meta name="ncc:generator" content="KOM-IN-DaisyCreator"/>'
                    + '\r\n')
            totalElapsedTime = timedelta(seconds=lTotalElapsedTime[z - 1])
            splittedTtotalElapsedTime = str(totalElapsedTime).split(".")
            print splittedTtotalElapsedTime
            totalElapsedTimehhmmss = splittedTtotalElapsedTime[0].zfill(8)
            if z == 1:
                # first entry has only one split
                totalElapsedTimeMilliMicro = "000"
            else:
                totalElapsedTimeMilliMicro = splittedTtotalElapsedTime[1][0:3]

            fOutFile.write(
                    '<meta name="ncc:totalElapsedTime" content="'
                    + totalElapsedTimehhmmss + "."
                    + totalElapsedTimeMilliMicro + '"/>' + '\r\n')

            fileTime = timedelta(seconds=lFileTime[z - 1])
            splittedFileTime = str(fileTime).split(".")
            FileTimehhmmss = splittedFileTime[0].zfill(8)
            # if no millissec, only one element is in the list
            if len(splittedFileTime) > 1:
                if len(splittedFileTime[1]) >= 3:
                    fileTimeMilliMicro = splittedFileTime[1][0:3]
                elif len(splittedFileTime[1]) == 2:
                        fileTimeMilliMicro = splittedFileTime[1][0:2]
            else:
                fileTimeMilliMicro = "000"

            fOutFile.write(
                '<meta name="ncc:timeInThisSmil" content="'
                    + FileTimehhmmss + "." + fileTimeMilliMicro
                    + '" />' + '\r\n')
            fOutFile.write(
                    '<meta name="dc:format" content="Daisy 2.02"/>' + '\r\n')
            fOutFile.write(
                    '<meta name="dc:identifier" content="'
                    + self.lineEditMetaRefOrig.text() + '"/>' + '\r\n')
            fOutFile.write(
                    '<meta name="dc:title" content="'
                    + cTitle + '"/>' + '\r\n')
            fOutFile.write('<layout>' + '\r\n')
            fOutFile.write('<region id="txt-view"/>' + '\r\n')
            fOutFile.write('</layout>' + '\r\n')
            fOutFile.write('</head>' + '\r\n')
            fOutFile.write('<body>' + '\r\n')
            lFileTimeSeconds = str(lFileTime[z - 1]).split(".")
            fOutFile.write(
                    '<seq dur="' + lFileTimeSeconds[0] + '.'
                    + fileTimeMilliMicro + 's">' + '\r\n')
            fOutFile.write('<par endsync="last">' + '\r\n')
            fOutFile.write(
                    '<text src="ncc.html#cnt_' + str(z).zfill(4)
                    + '" id="txt_' + str(z).zfill(4) + '" />' + '\r\n')
            fOutFile.write('<seq>' + '\r\n')
            if fileTime < timedelta(seconds=45):
                fOutFile.write(
                        '<audio src="' + item
                        + '" clip-begin="npt=0.000s" clip-end="npt='
                        + lFileTimeSeconds[0] + '.' + fileTimeMilliMicro
                        + 's" id="a_' + str(z).zfill(4) + '" />' + '\r\n')
            else:
                fOutFile.write(
                        '<audio src="' + item
                        + '" clip-begin="npt=0.000s" clip-end="npt='
                        + str(15) + '.' + fileTimeMilliMicro + 's" id="a_'
                        + str(z).zfill(4) + '" />' + '\r\n')
                zz = z + 1
                phraseSeconds = 15
                while phraseSeconds <= lFileTime[z - 1] - 15:
                    fOutFile.write(
                            '<audio src="' + item + '" clip-begin="npt='
                            + str(phraseSeconds) + '.' + fileTimeMilliMicro
                            + 's" clip-end="npt=' + str(phraseSeconds + 15)
                            + '.' + fileTimeMilliMicro + 's" id="a_'
                            + str(zz).zfill(4) + '" />' + '\r\n')
                    phraseSeconds += 15
                    zz += 1
                fOutFile.write(
                        '<audio src="' + item + '" clip-begin="npt='
                        + str(phraseSeconds) + '.' + fileTimeMilliMicro
                        + 's" clip-end="npt=' + lFileTimeSeconds[0] + '.'
                        + fileTimeMilliMicro + 's" id="a_' + str(zz).zfill(4)
                        + '" />' + '\r\n')

            fOutFile.write('</seq>' + '\r\n')
            fOutFile.write('</par>' + '\r\n')
            fOutFile.write('</seq>' + '\r\n')

            fOutFile.write('</body>' + '\r\n')
            fOutFile.write('</smil>' + '\r\n')
            fOutFile.close

    def splitFilename(self, item):
        """split filename into list"""
        if self.comboBoxDaisyTrenner.currentText() == "_-_":
            itemSplit = item.split(self.comboBoxDaisyTrenner.currentText())
            #itemSplit = item.split("_", 2)
        self.showDebugMessage("split filename into list")
        self.showDebugMessage(itemSplit)
        self.showDebugMessage(len(itemSplit))
        return itemSplit

    def extractAuthor(self, itemSplit):
        """extract author """
        return re.sub("_", " ", itemSplit[0])[2:]

    def extractTitle(self, itemSplit):
        """extract title """
        # letzter teil
        itemLeft = itemSplit[len(itemSplit) - 1]
        # davon file-ext abtrennen
        itemTitle = itemLeft.split(".mp3")
        cTitle = re.sub("_", " ", itemTitle[0])
        self.showDebugMessage("extract title")
        self.showDebugMessage(cTitle)
        return cTitle

    def showDialogCritical(self, errorMessage):
        QtGui.QMessageBox.critical(self, "Achtung", errorMessage)

    def showDebugMessage(self, debugMessage):
        if self.app_debugMod == "yes":
            print debugMessage

    def actionQuit(self):
        QtGui.qApp.quit()

    def main(self):
        self.showDebugMessage(u"let's rock")
        self.readConfig()
        self.checkPackages(self.app_lame)
        self.progressBarCopy.setValue(0)
        self.progressBarDaisy.setValue(0)
        # Bhz in Combo
        for item in self.app_bhzItems:
            self.comboBoxCopyBhz.addItem(item)
        # Issue nr in Comboncc:maxPageNormal
        prevYear = str(datetime.datetime.now().year - 1)
        currentYear = str(datetime.datetime.now().year)
        nextYear = str(datetime.datetime.now().year + 1)
        for item in self.app_prevAusgItems:
            self.comboBoxCopyBhzAusg.addItem(prevYear + "_" + item)
        for item in self.app_currentAusgItems:
            self.comboBoxCopyBhzAusg.addItem(currentYear + "_" + item)
        for item in self.app_nextAusgItems:
            self.comboBoxCopyBhzAusg.addItem(nextYear + "_" + item)
        # Separator in Combo
        self.comboBoxDaisyTrenner.addItem("_-_")
        self.comboBoxDaisyTrenner.addItem("Ausgabe-Nr.")
        self.comboBoxDaisyTrenner.addItem(prevYear)
        self.comboBoxDaisyTrenner.addItem(currentYear)
        self.comboBoxDaisyTrenner.addItem(nextYear)
        self.comboBoxDaisyTrenner.addItem("-")
        self.comboBoxDaisyTrenner.addItem("_")

        self.comboBoxPrefBitrate.addItem("64")
        self.comboBoxPrefBitrate.addItem("96")
        self.comboBoxPrefBitrate.addItem("128")
        self.comboBoxPrefBitrate.setCurrentIndex(1)
        # Conditions Checkboxes
        self.checkBoxCopyBhzIntro.setChecked(True)
        self.checkBoxCopyBhzIssueAnnouncement.setChecked(True)
        self.checkBoxCopyID3Change.setChecked(True)
        self.checkBoxCopyBitrateChange.setChecked(True)
        # Conditions spinboxes
        self.spinBoxLevel.setValue(1)
        self.spinBoxPages.setValue(0)
        # Help-Text
        self.readHelp()
        self.show()


if __name__ == '__main__':
    app = QtGui.QApplication(sys.argv)
    dyc = DaisyCopy()
    dyc.main()
    app.exec_()
    # This shows the interface we just created. No logic has been added, yet.
