#!/usr/bin/python
# -*- coding: utf-8 -*-

"""
Autor: Joerg Sorge
Distributed under the terms of GNU GPL version 2 or later
Copyright (C) Joerg Sorge joergsorge@gmail.com
2012-06-20

Dieses Programm 
- kopiert mp3-Files fuer die Verarbeitung zu Daisy-Buechern
- erzeugt die noetigen Dateien fuer eine Daisy-Struktur. 

Zusatz-Modul benoetigt: 
python-mutagen
sudo apt-get install python-mutagen

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
import types
import string
import re
from mutagen.mp3 import MP3
from mutagen.id3 import ID3
from mutagen.id3 import ID3NoHeaderError
import ConfigParser
import daisy_creator_mag_s_ui

#TODO:  Trenner in mag und abc ebenso anpassen (itemSecond zu left usw?
#TODO:  2.Progress  daisy einbauen 
#TODO:  config file read

class DaisyCopy(QtGui.QMainWindow, daisy_creator_mag_s_ui.Ui_DaisyMain):
    """ The second parent must be 'Ui_<obj. name of main widget class>'.
       
       """
 
    def __init__(self, parent=None):
        super(DaisyCopy, self).__init__(parent)
        # This is because Python does not automatically
        # call the parent's constructor.
        self.setupUi(self)
        # Pass this "self" for building widgets and
        # keeping a reference.
        self.app_debugMod = "yes"
        self.app_bhzItems = ["Zeitschrift"]        
        self.app_prevAusgItems = ["10",  "11",  "12",  "22",  "23",  "24",  "25"]
        self.app_currentAusgItems = ["01", "02", "03",  "04",  "05",  "06",  "07",  "08",  "09",  "10",  "11",  "12",  "13",  "14",  "15",  "16",  "17",  "18",  "19",  "20",  "21",  "22",  "23",  "24",  "25"]
        self.app_nextAusgItems = ["01",  "02",  "03",  "04"]
        self.app_bhzPfad = QtCore.QDir.homePath() 
        self.app_bhzPfadZiel = QtCore.QDir.homePath() 
        self.app_bhzPfadMeta = QtCore.QDir.homePath() 
        self.app_bhzPfadAusgabeansage = QtCore.QDir.homePath() 
        self.app_bhzPfadIntro = QtCore.QDir.homePath() 
        self.connectActions()
    
    def connectActions(self):
        """Actions definieren"""
        self.toolButtonCopySource.clicked.connect(self.actionOpenCopySource)
        self.toolButtonCopyDest.clicked.connect(self.actionOpenCopyDest)
        self.toolButtonCopyFile1.clicked.connect(self.actionOpenCounterFile)
        self.toolButtonMetaFile.clicked.connect(self.actionOpenMetaFile)
        self.commandLinkButton.clicked.connect(self.actionRunCopy)
        self.commandLinkButtonMeta.clicked.connect(self.metaLoadFile)
        self.commandLinkButtonDaisy.clicked.connect(self.actionRunDaisy)
        self.toolButtonDaisySource.clicked.connect(self.actionOpenDaisySource)
        self.pushButton.clicked.connect(self.actionQuit)

    def readConfig(self ):
        """read Config from file"""
        fileNotExist = None
        try:
            with open( "daisy_creator_mag_s.config" ) as f: pass
        except IOError as e:
            self.showDebugMessage(  u"File not exists" )
            self.textEdit.append("<b>Config-Datei konnte nicht geladen werden...</b>")   
            fileNotExist = "yes"
        
        if  fileNotExist is not None:
            return
        
        config = ConfigParser.RawConfigParser()
        config.read("daisy_creator_mag_s.config")
        self.app_bhzPfad = config.get('Ordner', 'BHZ')
        self.app_bhzPfadZiel = config.get('Ordner', 'BHZ-Ziel')
        self.app_bhzPfadMeta = config.get('Ordner', 'BHZ-Meta')
        self.app_bhzPfadAusgabeansage  = config.get('Ordner', 'BHZ-Ausgabeansage')
        self.app_bhzPfadIntro  = config.get('Ordner', 'BHZ-Intro')
        self.app_bhzItems  = config.get('Blindenhoerzeitschriften', 'BHZ').split(",")


    def actionOpenCopySource(self):
        """Quelle fuer copy"""
        # QtCore.QDir.homePath() 
        dirSource = QtGui.QFileDialog.getExistingDirectory(
                        self,
                        "Quell-Ordner",
                        self.app_bhzPfad
                    )
        # Don't attempt to open if open dialog
        # was cancelled away.
        if dirSource:
            self.lineEditCopySource.setText(dirSource)
            self.textEdit.append("Quelle:")
            self.textEdit.append(dirSource)
    
    def actionOpenDaisySource(self):
        """Quelle fuer daisy"""
        dirSource = QtGui.QFileDialog.getExistingDirectory(
                        self,
                        "Quell-Ordner",
                        self.app_bhzPfadZiel
                    )
        # Don't attempt to open if open dialog
        # was cancelled away.
        if dirSource:
            self.lineEditDaisySource.setText(dirSource)
            self.textEdit.append("Quelle:")
            self.textEdit.append(dirSource)
    
    def actionOpenCopyDest(self):
        """Ziel fuer Copy"""
        dirDest = QtGui.QFileDialog.getExistingDirectory(
                        self,
                        "Ziel-Ordner",
                        self.app_bhzPfadZiel
                    )
        # Don't attempt to open if open dialog
        # was cancelled away.
        if dirDest:
            self.lineEditCopyDest.setText(dirDest)
            self.textEdit.append("Ziel:")
            self.textEdit.append(dirDest)
    
    def actionOpenCounterFile(self):
        """Counterdatei zur Steuerung"""
        file1 = QtGui.QFileDialog.getOpenFileName (
                        self,
                        "Counter-Datei",
                        self.app_bhzPfad
                    )
        # Don't attempt to open if open dialog
        # was cancelled away.
        if file1:
            self.lineEditFileCounter.setText(file1)
            self.textEdit.append("Counter-Datei:")
            self.textEdit.append(file1)

    def actionOpenCopyFile2(self):
        """Zusatzdatei 2 zum kopieren"""
        file2 = QtGui.QFileDialog.getOpenFileName (
                        self,
                        "Datei 2",
                        QtCore.QDir.homePath()
                    )
        # Don't attempt to open if open dialog
        # was cancelled away.
        if file2:
            self.lineEdit_4.setText(file2)
            self.textEdit.append("Zusatz-Datei 2:")
            self.textEdit.append(file2)
    
    def actionOpenMetaFile(self):
        """Metadatei zum laden"""
        file = QtGui.QFileDialog.getOpenFileName (
                        self,
                        "Daisy_Meta",
                        self.app_bhzPfadMeta
                    )
        # Don't attempt to open if open dialog
        # was cancelled away.
        if file:
            self.lineEditMetaSource.setText(file)
    
    def actionRunCopy(self):
        """Hauptroutine zum Kopieren"""
        if self.lineEditCopySource.text()== "Quell-Ordner":
            errorMessage = u"Quell-Ordner wurde nicht ausgewaehlt.."
            self.showDialogCritical( errorMessage )
            return
            
        if self.lineEditCopyDest.text()== "Ziel-Ordner":
            errorMessage = u"Ziel-Ordner wurde nicht ausgewaehlt.."
            self.showDialogCritical( errorMessage )
            return
        
        if self.lineEditFileCounter.text()== "Counter-Datei waehlen":
            if self.comboBoxCopyBhz.currentText() == "ABC_Journal":
                errorMessage = u"Counter-Datei wurde nicht ausgewaehlt.."
                self.showDialogCritical( errorMessage )
                return
        else:
            counterLines = self.readCounterFile()
        
        self.showDebugMessage(  self.lineEditCopySource.text() )
        self.showDebugMessage(  self.lineEditCopyDest.text() )
        
        try:
            dirsSource = os.listdir( self.lineEditCopySource.text())
        except Exception, e:
            logMessage = u"read_files_from_dir Error: %s" % str(e)
            self.showDebugMessage( logMessage)
        
        self.showDebugMessage( dirsSource )
        # Anfangswert des Counters erhoehen wenn AusgAnsage und/oder Intro kopiert werden
        zCounter =0
        if self.checkBoxCopyBhzIntro.isChecked():
            zCounter +=1
        if self.checkBoxCopyBhzAusgAnsage.isChecked():
            zCounter +=1
        
        self.textEdit.append("<b>Kopieren:</b>")
        z = 0
        zList = len(dirsSource)
        self.showDebugMessage(  zList )
        dirsSource.sort()
        # files durchgehen
        for item in dirsSource:
            if item[ len(item)-4:len(item) ] == ".MP3" or item[ len(item)-4:len(item) ] == ".mp3":
                fileToCopySource = self.lineEditCopySource.text() + "/" + item 
                # pruefen ob file exists
                fileNotExist = None
                try:
                    with open( fileToCopySource ) as f: pass
                except IOError as e:
                    self.showDebugMessage(  u"File not exists" )
                    fileNotExist = "yes"
                    # max Anzahl korrigieren und Progress aktualisieren
                    zList = zList -1
                    pZ = z *100 / zList 
                    self.progressBar.setValue(pZ)
                
                self.showDebugMessage( fileToCopySource )
                # TODO:  Irgendwie die max Anzahl von files in Counterdatei ermitteln
                # und mit max Anzahl der audios vergleichen, wenn nicht gleich dann gar nicht Daisy zulassen
                if  fileNotExist is None:
                    # Dateiname in counterDatei suchen, Counter extrahieren und in Dateiname
                    zCounterFiles =0
                    for line in counterLines:
                        if line.find(item) != -1:
                            # Achtung: er findet hier natuerlich auch unvollstaendige dateinamen, 
                            # z.B. wenn vorn nur paar zeichen fehlen oder welche zuviel sind 
                            # wird der rest, der uebereinstimmt gefunden
                            zCounterFiles +=1
                            #self.textEdit.append(u"counterline"+ line)
                            self.showDebugMessage(  u"counterline"+ line)
                            self.showDebugMessage(  u"counter"+ line[9:11])
                            # Counter vorne dran
                            zFileCount = zCounter + int(line[9:11] )
                            #fileToCopyDest = self.lineEditCopyDest.text() + "/" + line[9:11] + "_" + item
                            fileToCopyDest = self.lineEditCopyDest.text() + "/" + str(zFileCount).zfill(2) + "_" + item
                    
                    # dateiname nicht gefunden
                    if zCounterFiles == 0:
                        self.textEdit.append("<b><font color='red'>Dateiname stimmt nicht ueberein</font></b>:")
                        self.textEdit.append( ntpath.basename(str(fileToCopySource)))
                        continue
                    
                    self.textEdit.append(ntpath.basename(str(fileToCopyDest)))
                    self.showDebugMessage( fileToCopySource )
                    self.showDebugMessage( fileToCopyDest )
                    
                    # Bitrate checken, eventuell aendern und gleich in Ziel neu encodieren
                    isChangedAndCopy = self.checkChangeBitrateAndCopy( fileToCopySource,  fileToCopyDest )
                    # nicht geaendert also kopieren
                    if  isChangedAndCopy is None: 
                        self.copyFile( fileToCopySource, fileToCopyDest)
                    
                    self.checkCangeId3( fileToCopyDest)
                    z +=1
                    self.showDebugMessage( z )
                    self.showDebugMessage( zList )
                    pZ = z *100 / zList 
                    self.showDebugMessage( pZ )
                    self.progressBar.setValue(pZ)
                else:
                    self.textEdit.append("<b>Uebersprungen</b>:")
                    self.textEdit.append(fileToCopySource)
        
        self.showDebugMessage( z )
        
        if self.checkBoxCopyBhzIntro.isChecked():
            self.copyIntro()
        
        if self.checkBoxCopyBhzAusgAnsage.isChecked():
            self.copyAusgabeAnsage()
        
        # TODO: Hier weiter
        # Metadaten laden
        self.lineEditMetaSource.setText(self.app_bhzPfadMeta + "/Daisy_Meta_" + self.comboBoxCopyBhz.currentText()  )
        self.metaLoadFile()
        # Zielpfad als Quellpfad fuer Daisy eintragen
        self.lineEditDaisySource.setText(self.lineEditCopyDest.text())

   
    def copyFile(self, fileToCopySource, fileToCopyDest):
       """Datei kopieren"""
       try:
            shutil.copy( fileToCopySource, fileToCopyDest )
       except Exception, e:
            logMessage = u"copy_file Error: %s" % str(e)
            self.showDebugMessage(logMessage)
            self.textEdit.append(logMessage + fileToCopyDest)
    
    def copyIntro(self):
        """Intro kopieren"""
        fileToCopySource = self.app_bhzPfadIntro + "/Intro_" +  self.comboBoxCopyBhz.currentText()  + ".mp3"
        fileToCopyDest = self.lineEditCopyDest.text() + "/02_" +  self.comboBoxCopyBhz.currentText() + "_-_" + self.comboBoxCopyBhzAusg.currentText() + "_Intro.mp3"
        self.showDebugMessage(fileToCopySource)
        self.showDebugMessage(fileToCopyDest)
        fileNotExist = None
        try:
            with open( fileToCopySource ) as f: pass
        except IOError as e:
            self.showDebugMessage(  u"File not exists" )
            self.textEdit.append("<b><font color='red'>Intro konnte nicht kopiert werden</font></b>:")
            self.textEdit.append(fileToCopySource)
            fileNotExist = "yes"
        
        if  fileNotExist is None:
            #shutil.copy( fileToCopySource, fileToCopyDest )
            self.copyFile( fileToCopySource, fileToCopyDest)
            self.checkCangeId3( fileToCopyDest)
    
    def copyAusgabeAnsage(self):
        """Intro kopieren"""
        pfadAusgabe = self.app_bhzPfadAusgabeansage + "_" + self.comboBoxCopyBhzAusg.currentText()[0:4] + "_" + self.comboBoxCopyBhz.currentText() 
        self.showDebugMessage(pfadAusgabe)
        fileToCopySource = pfadAusgabe + "/" +  self.comboBoxCopyBhz.currentText() + "_-_Ausgabe_" + self.comboBoxCopyBhzAusg.currentText() + ".mp3"
        fileToCopyDest = self.lineEditCopyDest.text() + "/01_" +  self.comboBoxCopyBhz.currentText() + "_-_Ausgabe_" + self.comboBoxCopyBhzAusg.currentText() + ".mp3"
        self.showDebugMessage(fileToCopySource)
        self.showDebugMessage(fileToCopyDest)
        fileNotExist = None
        try:
            with open( fileToCopySource ) as f: pass
        except IOError as e:
            self.showDebugMessage(  u"File not exists" )
            self.textEdit.append("<b><font color='red'>Ausgabeansage konnte nicht kopiert werden</font></b>:")
            self.textEdit.append(fileToCopySource)
            
            fileNotExist = "yes"
        
        if  fileNotExist is None:
            #shutil.copy( fileToCopySource, fileToCopyDest )
            self.copyFile( fileToCopySource, fileToCopyDest)
            self.checkCangeId3( fileToCopyDest)
    
    def readCounterFile(self):
        """Datei mit Dateinamen und Counter einlesen """
        try:
            f = open( str(self.lineEditFileCounter.text()), 'r')
            lines = f.readlines() 
            f.close()
            self.textEdit.append("<b>Counter-Datei gelesen: </b>" + self.lineEditFileCounter.text())
            #for item in lines:
            #    self.textEdit.append("line: " + item)
        except IOError as (errno, strerror):
            lines = None
            self.showDebugMessage("read_from_file_lines_in_list: nicht lesbar: " + filename)
        return lines
    
    def checkCangeId3(self, fileToCopyDest):
        """id3 Tags pruefen, ev. entfernen"""
        tag = None
        try:
            audio = ID3(fileToCopyDest)
            tag = "yes"
        except ID3NoHeaderError:
            self.showDebugMessage( u"No ID3 header found; skipping.")
        
        if tag is not None:
            if self.checkBoxCopyID3Change.isChecked():
                audio.delete()
                self.textEdit.append("<b>ID3 entfernt bei</b>: " + ntpath.basename( str(fileToCopyDest)))
                self.showDebugMessage( u"ID3 entfernt bei " + fileToCopyDest)
            else:
                self.textEdit.append("<b>ID3 vorhanden, aber NICHT entfernt bei</b>: " + fileToCopyDest)
    
    def checkChangeBitrateAndCopy(self,  fileToCopySource, fileToCopyDest):
        """Bitrate aendern und an Ziel encodieren"""
        isChangedAndCopy = None
        audioSource = MP3( fileToCopySource )
        if audioSource.info.bitrate != int(self.comboBoxPrefBitrate.currentText())*1000:
            isEncoded = None
            self.textEdit.append( u"Bitrate Vorgabe: " + str(self.comboBoxPrefBitrate.currentText()) )
            self.textEdit.append( u"<b>Bitrate folgender Datei entspricht nicht der Vorgabe:</b> " + str(audioSource.info.bitrate/1000) + " " + fileToCopySource)
            
            if self.checkBoxCopyBitrateChange.isChecked():
                self.textEdit.append(u"<b>Bitrate aendern bei</b>: " + fileToCopyDest)
                isEncoded = self.encodeFile( fileToCopySource, fileToCopyDest )
                if isEncoded is not None:
                    self.textEdit.append(u"<b>Bitrate geaendert bei</b>: " + fileToCopyDest)
                    isChangedAndCopy = True
            else:
                self.textEdit.append(u"<b>Bitrate wurde NICHT geaendern bei</b>: " + fileToCopyDest)
        return isChangedAndCopy
    
    def encodeFile(self, fileToCopySource, fileToCopyDest ):
        """mp3-files mit entspr Bitrate encoden"""
        self.showDebugMessage( u"encode_file" )
        #damit die uebergabe der befehle richtig klappt muessen alle cmds im richtigen zeichensatz als strings encoded sein
        c_lame_encoder = "/usr/bin/lame"
        self.showDebugMessage(  u"type c_lame_encoder" )
        self.showDebugMessage( type(c_lame_encoder) )
        self.showDebugMessage( u"fileToCopySource" )
        self.showDebugMessage( type(fileToCopySource) )
        self.showDebugMessage( fileToCopyDest )
        self.showDebugMessage( u"type(fileToCopyDest)" )
        self.showDebugMessage( type( fileToCopyDest ) )
        
        #p = subprocess.Popen([c_lame_encoder, "-b",  "64", "-m",  "m",  fileToCopySource, fileToCopyDest ],  stdout=subprocess.PIPE, stderr=subprocess.PIPE).communicate(  )
        p = subprocess.Popen([c_lame_encoder, "-b",  self.comboBoxPrefBitrate.currentText(), "-m",  "m",  fileToCopySource, fileToCopyDest ],  stdout=subprocess.PIPE, stderr=subprocess.PIPE).communicate(  )

        self.showDebugMessage( u"returncode 0" )
        self.showDebugMessage( p[0] )
        self.showDebugMessage( u"returncode 1" )
        self.showDebugMessage( p[1] )
    
        # erfolgsmeldung suchen, wenn nicht gefunden: -1
        n_encode_percent = string.find( p[1],  "(100%)" )
        n_encode_percent_1 = string.find( p[1],  "(99%)" )
        self.showDebugMessage( n_encode_percent )
        c_complete = "no"
    
        # bei kurzen files kommt die 100% meldung nicht, deshalb auch 99% durchgehen lassen
        if n_encode_percent == -1:
            # 100% nicht erreicht
            if n_encode_percent_1 != -1:
                # aber 99
                c_complete = "yes"
        else:
            c_complete = "yes"
    
        if c_complete == "yes" :
            log_message = u"recoded_file: " + fileToCopySource
            self.showDebugMessage( log_message )
            return fileToCopyDest
        else:
            log_message = u"recode_file Error: " + fileToCopySource
            self.showDebugMessage( log_message )
            return None
    
    def metaLoadFile(self ):
        config = ConfigParser.RawConfigParser()
        # Pfad von QTString in String  umwandeln
        config.read(str(self.lineEditMetaSource.text()))
        self.lineEditMetaProducer.setText(config.get('Daisy_Meta', 'Produzent'))
        self.lineEditMetaAutor.setText(config.get('Daisy_Meta', 'Autor'))
        self.lineEditMetaTitle.setText(config.get('Daisy_Meta', 'Titel'))
        self.lineEditMetaEdition.setText(config.get('Daisy_Meta', 'Edition'))
        self.lineEditMetaNarrator.setText(config.get('Daisy_Meta', 'Sprecher'))
        self.lineEditMetaKeywords.setText(config.get('Daisy_Meta', 'Stichworte'))
        self.lineEditMetaRefOrig.setText(config.get('Daisy_Meta', 'ISBN/Ref-Nr.Original'))
        self.lineEditMetaPublisher.setText(config.get('Daisy_Meta', 'Verlag'))
        self.lineEditMetaYear.setText(config.get('Daisy_Meta', 'Jahr'))
    
    def actionRunDaisy(self):
        """Daisy-Fileset erzeugen"""
        if self.lineEditDaisySource.text()== "Quell-Ordner":
            errorMessage = u"Quell-Ordner wurde nicht ausgewaehlt.."
            self.showDialogCritical( errorMessage )
            return
        
        # Audios einlesen
        try:
            dirItems = os.listdir( self.lineEditDaisySource.text())
        except Exception, e:
            logMessage = u"read_files_from_dir Error: %s" % str(e)
            self.showDebugMessage( logMessage)
        
        self.showDebugMessage( dirItems )
        self.textEditDaisy.append(u"<b>Folgende Audios werden bearbeitet:</b>")
        zMp3 = 0
        zList = len(dirItems)
        self.showDebugMessage(  zList )
        dirAudios = []
        dirItems.sort()
        for item in dirItems:
            if item[ len(item)-4:len(item) ] == ".MP3" or item[ len(item)-4:len(item) ] == ".mp3":
                dirAudios.append(item) 
                self.textEditDaisy.append(item)
                zMp3 +=1

        #totalAudioLength = self.calcAudioLengt(dirAudios)
        lTimes = self.calcAudioLengt(dirAudios)
        totalAudioLength = lTimes[0]
        lTotalElapsedTime = lTimes[1]
        lFileTime = lTimes[2]
        print totalAudioLength
        totalTime = timedelta(seconds = totalAudioLength)
        # umwandlung von timedelta in string: minuten und sekunden musten immer zweistllig sein, 
        # damit einstellige stunde eine null bekommt :zfill(8)
        lTotalTime = str(totalTime).split(".")
        cTotalTime = lTotalTime[0].zfill(8)
        #str(cTotalTime[0]).zfill(8) 
        self.textEditDaisy.append(u"Gesamtlaenge: " + cTotalTime)
        self.writeNCC( cTotalTime,  zMp3,  dirAudios)
        self.writeMasterSmil( cTotalTime,  dirAudios)
        self.writeSmil( lTotalElapsedTime,  lFileTime, dirAudios)
    
    def calcAudioLengt(self,  dirAudios):
        """Gesamtlange der Audios ermitteln"""
        totalAudioLength =0
        lTotalElapsedTime = []
        lTotalElapsedTime.append(0)
        lFileTime = []
        for item in dirAudios:
            fileToCheck = os.path.join( str(self.lineEditDaisySource.text()), item) 
            audioSource = MP3( fileToCheck )
            self.showDebugMessage(item + " "+ str(audioSource.info.length))
            totalAudioLength += audioSource.info.length
            lTotalElapsedTime.append(totalAudioLength)
            lFileTime.append(audioSource.info.length)
            lTimes = []
            lTimes.append(totalAudioLength)
            lTimes.append(lTotalElapsedTime)
            lTimes.append(lFileTime)
        return lTimes
    
    def writeNCC(self, cTotalTime,  zMp3,  dirAudios):
        """NCC-Page schreiben"""
        try:
            fOutFile = open( os.path.join( str(self.lineEditDaisySource.text()), "ncc.html")  , 'w')
        except IOError as (errno, strerror):
            self.showDebugMessage( "I/O error({0}): {1}".format(errno, strerror) )
        else:
            self.textEditDaisy.append(u"<b>NCC-Datei schreiben</b>")
            fOutFile.write( '<?xml version="1.0" encoding="utf-8"?>'+ '\r\n' )
            fOutFile.write( '<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">'+ '\r\n')
            fOutFile.write( '<html xmlns="http://www.w3.org/1999/xhtml">'+ '\r\n')
            fOutFile.write( '<head>'+ '\r\n')
            fOutFile.write( '<meta http-equiv="Content-type" content="text/html; charset=utf-8"/>'+ '\r\n')
            fOutFile.write( '<title>' + self.comboBoxCopyBhz.currentText() + '</title>'+ '\r\n')

            fOutFile.write( '<meta name="ncc:generator" content="KOM-IN-DaisyCreator"/>'+ '\r\n')
            fOutFile.write( '<meta name="ncc:revision" content="1"/>'+ '\r\n')
            today = datetime.date.today()
            fOutFile.write( '<meta name="ncc:producedDate" content="' + today.strftime("%Y-%m-%d") + '"/>'+ '\r\n')
            fOutFile.write( '<meta name="ncc:revisionDate" content="' + today.strftime("%Y-%m-%d") + '"/>'+ '\r\n')
            fOutFile.write( '<meta name="ncc:tocItems" content="' + str( zMp3 ) + '"/>'+ '\r\n')
            
            fOutFile.write( '<meta name="ncc:totalTime" content="' + cTotalTime+ '"/>'+ '\r\n')
            fOutFile.write( '<meta name="ncc:narrator" content="' + self.lineEditMetaNarrator.text() + '"/>'+ '\r\n')
            fOutFile.write( '<meta name="ncc:pageNormal" content="0"/>'+ '\r\n')
            fOutFile.write( '<meta name="ncc:pageFront" content="0"/>'+ '\r\n')
            fOutFile.write( '<meta name="ncc:pageSpecial" content="0"/>'+ '\r\n')
            fOutFile.write( '<meta name="ncc:sidebars" content="0"/>'+ '\r\n')
            fOutFile.write( '<meta name="ncc:prodnotes" content="0"/>'+ '\r\n')
            fOutFile.write( '<meta name="ncc:footnotes" content="0"/>'+ '\r\n')
            fOutFile.write( '<meta name="ncc:depth" content="' + str(self.spinBoxEbenen.value())+ '"/>'+ '\r\n')
            fOutFile.write( '<meta name="ncc:maxPageNormal" content="' +str(self.spinBoxPages.value()) +'"/>'+ '\r\n')
            fOutFile.write( '<meta name="ncc:charset" content="utf-8"/>'+ '\r\n')
            fOutFile.write( '<meta name="ncc:multimediaType" content="audioNcc"/>'+ '\r\n')
            #fOutFile.write( '<meta name="ncc:kByteSize" content=" "/>'+ '\r\n')
            fOutFile.write( '<meta name="ncc:setInfo" content="1 of 1"/>'+ '\r\n')
            
            fOutFile.write( '<meta name="ncc:sourceDate" content="' + self.lineEditMetaYear.text()+ '"/>'+ '\r\n')
            fOutFile.write( '<meta name="ncc:sourceEdition" content="' + self.lineEditMetaEdition.text() + '"/>'+ '\r\n')
            fOutFile.write( '<meta name="ncc:sourcePublisher" content="' + self.lineEditMetaPublisher.text()+ '"/>'+ '\r\n')

            #Anzahl files = Records 2x + ncc.html + master.smil
            fOutFile.write( '<meta name="ncc:files" content="' + str(zMp3 + zMp3 + 2) + '"/>'+ '\r\n')
            fOutFile.write( '<meta name="ncc:format" content="Daisy 2.0"/>'+ '\r\n')
            
            fOutFile.write( '<meta name="ncc:producer" content="' + self.lineEditMetaProducer.text()+ '"/>'+ '\r\n')
            fOutFile.write( '<meta name="ncc:Charset" content="ISO-8859-1"/>'+ '\r\n')

            fOutFile.write( '<meta name="dc:creator" content="' + self.lineEditMetaAutor.text()+ '"/>'+ '\r\n')
            fOutFile.write( '<meta name="dc:date" content="' + today.strftime("%Y-%m-%d")+ '"/>'+ '\r\n')
            fOutFile.write( '<meta name="dc:format" content="Daisy 2.02"/>'+ '\r\n')
            fOutFile.write( '<meta name="dc:identifier" content="' + self.lineEditMetaRefOrig.text()+ '"/>'+ '\r\n')
            fOutFile.write( '<meta name="dc:language" content="de" scheme="ISO 639"/>'+ '\r\n')
            fOutFile.write( '<meta name="dc:publisher" content="' + self.lineEditMetaPublisher.text()+ '"/>'+ '\r\n')

            fOutFile.write( '<meta name="dc:source" content="' +self.lineEditMetaRefOrig.text()+ '"/>'+ '\r\n')
            fOutFile.write( '<meta name="dc:sourceDate" content="' + self.lineEditMetaYear.text()+ '"/>'+ '\r\n')
            fOutFile.write( '<meta name="dc:sourceEdition" content="' +self.lineEditMetaEdition.text()+ '"/>'+ '\r\n')
            fOutFile.write( '<meta name="dc:sourcePublisher" content="' + self.lineEditMetaPublisher.text()+ '"/>'+ '\r\n')
            fOutFile.write( '<meta name="dc:subject" content="' + self.lineEditMetaKeywords.text()+ '"/>'+ '\r\n')
            fOutFile.write( '<meta name="dc:title" content="' +self.lineEditMetaTitle.text()+ '"/>'+ '\r\n')
            # Medibus-OK items
            fOutFile.write( '<meta name="prod:audioformat" content="wave 44 kHz"/>'+ '\r\n')
            fOutFile.write( '<meta name="prod:compression" content="mp3 ' + self.comboBoxPrefBitrate.currentText()  + '/ kb/s"/>'+ '\r\n')      
            fOutFile.write( '<meta name="prod:localID" content=" "/>'+ '\r\n')      
            fOutFile.write( '</head>'+ '\r\n')
            fOutFile.write( '<body>'+ '\r\n')
            #fOutFile.write('<h1 class="title" id="cnt_0001"><a href="0001.smil#txt_0001">' + self.lineEditMetaAutor.text()+ ": " + self.lineEditMetaTitle.text() + '</a></h1>'+ '\r\n')
            z = 0
            for item in dirAudios:
                z +=1
                if z == 1:
                    fOutFile.write('<h1 class="title" id="cnt_0001"><a href="0001.smil#txt_0001">' + self.lineEditMetaAutor.text()+ ": " + self.lineEditMetaTitle.text() + '</a></h1>'+ '\r\n')
                    continue
                # trennen
                if self.comboBoxDaisyTrenner.currentText()=="Ausgabe-Nr.":
                    itemSplit = item.split(self.comboBoxCopyBhzAusg.currentText()+"_")
                    print itemSplit
                    print len(itemSplit)
                else:
                    itemSplit = item.split(self.comboBoxDaisyTrenner.currentText())
                
                # erster Teil Autor
                # Unterstriche durch Leerzeichen ersetzen, zwei Stellen (Counterzahl weglassen ([2:])
                cAuthor = re.sub ("_", " ", itemSplit[0]  )[2:]
                # letzter teil Title
                itemSecond = itemSplit[len(itemSplit)-1]
                # davon file-ext abtrennen
                itemTitle = itemSecond.split(".mp3")
                print itemSplit
                print itemSplit[0]
                print itemSplit[1]
                # Unterrstriche durch Leerzeichen ersetzen
                cTitle = re.sub ("_", " ", itemTitle[0]  ) 
                
                fOutFile.write('<h'+ str(self.spinBoxEbenen.value())+' id="cnt_'+str(z).zfill(4)+'"><a href="'+str(z).zfill(4)+'.smil#txt_'+str(z).zfill(4)+'">'+ cAuthor +" - " + cTitle + '</a></h1>'+ '\r\n')
                           
            fOutFile.write( "</body>"+ '\r\n')
            fOutFile.write( "</html>"+ '\r\n')
            fOutFile.close

    def writeMasterSmil(self, cTotalTime,  dirAudios):
        """MasterSmil-Page schreiben"""
        try:
            fOutFile = open( os.path.join( str(self.lineEditDaisySource.text()), "master.smil")  , 'w')
        except IOError as (errno, strerror):
            self.showDebugMessage( "I/O error({0}): {1}".format(errno, strerror) )
        else:
            self.textEditDaisy.append(u"<b>MasterSmil-Datei schreiben</b>")
            fOutFile.write( '<?xml version="1.0" encoding="utf-8"?>'+ '\r\n' )
            fOutFile.write( '<!DOCTYPE smil PUBLIC "-//W3C//DTD SMIL 1.0//EN" "http://www.w3.org/TR/REC-smil/SMIL10.dtd">'+'\r\n')
            fOutFile.write( '<smil>'+'\r\n')
            fOutFile.write( '<head>'+'\r\n')
            fOutFile.write( '<meta name="dc:format" content="Daisy 2.02"/>'+'\r\n')
            fOutFile.write( '<meta name="dc:identifier" content="'+ self.lineEditMetaRefOrig.text()+ '"/>'+'\r\n')
            fOutFile.write( '<meta name="dc:title" content="' + self.lineEditMetaTitle.text() + '"/>'+'\r\n')
            fOutFile.write( '<meta name="ncc:generator" content="KOM-IN-DaisyCreator"/>'+'\r\n')
            fOutFile.write( '<meta name="ncc:format" content="Daisy 2.0"/>'+'\r\n')
            fOutFile.write( '<meta name="ncc:timeInThisSmil" content="' + cTotalTime +  '" />'+'\r\n')

            fOutFile.write( '<layout>'+'\r\n')
            fOutFile.write( '<region id="txt-view" />'+'\r\n')
            fOutFile.write( '</layout>'+'\r\n')
            fOutFile.write( '</head>'+'\r\n')
            fOutFile.write( '<body>'+'\r\n')
   
            #fOutFile.write( '<ref src="0001.smil" title="'+ self.lineEditMetaTitle.text() + '" id="smil_0001"/>'+'\r\n')            
            z = 0
            for item in dirAudios:
                z +=1
                # trennen
                if self.comboBoxDaisyTrenner.currentText()=="Ausgabe-Nr.":
                    itemSplit = item.split(self.comboBoxCopyBhzAusg.currentText()+"_")
                else:
                    itemSplit = item.split(self.comboBoxDaisyTrenner.currentText())
                # erster Teil Autor
                # Unterstriche durch Leerzeichen ersetzen, zwei Stellen (Counterzahl weglassen ([2:])
                cAuthor = re.sub ("_", " ", itemSplit[0]  )[2:]
                # letzter teil Title
                itemSecond = itemSplit[len(itemSplit)-1]
                # davon file-ext abtrennen
                itemTitle = itemSecond.split(".mp3")
                print itemSplit
                # Unterrstriche durch Leerzeichen ersetzen
                cTitle = re.sub ("_", " ", itemTitle[0]  ) 
                
                #fOutFile.write('<ref src="'+str(z).zfill(4)+'.smil" title="' + itemTitle[0] + '" id="smil_' + str(z).zfill(4) + '"/>'+'\r\n')
                fOutFile.write('<ref src="'+str(z).zfill(4)+'.smil" title="' +cAuthor + " - " +cTitle + '" id="smil_' + str(z).zfill(4) + '"/>'+'\r\n')
            
            fOutFile.write( '</body>'+'\r\n')
            fOutFile.write( '</smil>'+'\r\n')
            fOutFile.close

    def writeSmil(self, lTotalElapsedTime, lFileTime, dirAudios):
        """Smil-Pages schreiben"""
        z = 0
        for item in dirAudios:
            z +=1
            
            try:
                filename = str(z).zfill(4) +'.smil'
                fOutFile = open( os.path.join( str(self.lineEditDaisySource.text()), filename  ) , 'w')
            except IOError as (errno, strerror):
                self.showDebugMessage( "I/O error({0}): {1}".format(errno, strerror) )
            else:
                self.textEditDaisy.append( str(z).zfill(4) +u".smil - File schreiben")
                # trennen
                if self.comboBoxDaisyTrenner.currentText()=="Ausgabe-Nr.":
                    itemSplit = item.split(self.comboBoxCopyBhzAusg.currentText()+"_")
                else:
                    itemSplit = item.split(self.comboBoxDaisyTrenner.currentText())
                # erster Teil Autor
                # Unterstriche durch Leerzeichen ersetzen, zwei Stellen (Counterzahl weglassen ([2:])
                cAuthor = re.sub ("_", " ", itemSplit[0]  )[2:]
                # letzter teil Title
                itemSecond = itemSplit[len(itemSplit)-1]
                # davon file-ext abtrennen
                itemTitle = itemSecond.split(".mp3")
                print itemSplit
                # Unterrstriche durch Leerzeichen ersetzen
                
                fOutFile.write( '<?xml version="1.0" encoding="utf-8"?>'+ '\r\n' )
                fOutFile.write( '<!DOCTYPE smil PUBLIC "-//W3C//DTD SMIL 1.0//EN" "http://www.w3.org/TR/REC-smil/SMIL10.dtd">'+'\r\n')
                fOutFile.write( '<smil>'+'\r\n')
                fOutFile.write( '<head>'+'\r\n')
                fOutFile.write( '<meta name="ncc:generator" content="KOM-IN-DaisyCreator"/>'+'\r\n')
                fOutFile.write( '<meta name="ncc:format" content="Daisy 2.02"/>'+'\r\n')
                totalElapsedTime = timedelta(seconds = lTotalElapsedTime[z-1])
                splittedTtotalElapsedTime = str(totalElapsedTime).split(".")
                print splittedTtotalElapsedTime
                totalElapsedTimehhmmss = splittedTtotalElapsedTime[0].zfill(8)
                if z == 1:
                    # erster eintrag ergibt nur einen split
                    totalElapsedTimeMilliMicro = "000"
                else:
                    totalElapsedTimeMilliMicro = splittedTtotalElapsedTime[1][0:3] 
                #fOutFile.write( '<meta name="ncc:totalElapsedTime" content="' + str(lTotalElapsedTime[z-1] )+ '"/>')
                #fOutFile.write( '<meta name="ncc:totalElapsedTime" content="' + str( totalElapsedTime )+ '"/>'+'\r\n')
                fOutFile.write( '<meta name="ncc:totalElapsedTime" content="' +totalElapsedTimehhmmss + "." + totalElapsedTimeMilliMicro +'"/>'+'\r\n')
                
                fileTime = timedelta(seconds = lFileTime[z-1])
                splittedFileTime = str(fileTime).split(".")
                FileTimehhmmss = splittedFileTime[0].zfill(8)
                # wenn keine Millisicrosec gibts nur ein Element in der Liste
                if len(splittedFileTime) >1:
                    if len(splittedFileTime[1]) >= 3:
                        fileTimeMilliMicro = splittedFileTime[1][0:3] 
                    elif len(splittedFileTime[1]) == 2:
                        fileTimeMilliMicro = splittedFileTime[1][0:2] 
                else:
                    fileTimeMilliMicro = "000"
                
                #fOutFile.write( '<meta name="ncc:timeInThisSmil" content="' + str(lFileTime[z-1]) + '" />'+'\r\n')

                fOutFile.write( '<meta name="ncc:timeInThisSmil" content="' + FileTimehhmmss + "." + fileTimeMilliMicro +'" />'+'\r\n')
                fOutFile.write( '<meta name="dc:format" content="Daisy 2.02"/>'+'\r\n')
                fOutFile.write( '<meta name="dc:identifier" content="' + self.lineEditMetaRefOrig.text() + '"/>'+'\r\n')
                cTitle = re.sub ("_", " ", itemTitle[0]  ) 
                #fOutFile.write( '<meta name="dc:title" content="' +  itemTitle[0]  + '"/>'+'\r\n')
                fOutFile.write( '<meta name="dc:title" content="' +  cTitle  + '"/>'+'\r\n')
                fOutFile.write( '<layout>'+'\r\n')
                fOutFile.write( '<region id="txt-view"/>'+'\r\n')
                fOutFile.write( '</layout>'+'\r\n')
                fOutFile.write( '</head>'+'\r\n')
                fOutFile.write( '<body>'+'\r\n')
                #fOutFile.write( '<seq dur="' + FileTimehhmmss + '.' + fileTimeMilliMicro + 's">'+'\r\n')
                lFileTimeSeconds = str(lFileTime[z-1]).split(".")
                
                fOutFile.write( '<seq dur="' + lFileTimeSeconds[0] + '.' + fileTimeMilliMicro  +'s">'+'\r\n')
                fOutFile.write( '<par endsync="last">'+'\r\n')
                fOutFile.write( '<text src="ncc.html#cnt_' + str(z).zfill(4) + '" id="txt_' + str(z).zfill(4) + '" />'+'\r\n')
                fOutFile.write( '<seq>'+'\r\n')
                if fileTime < timedelta(seconds=45):
                    fOutFile.write( '<audio src="' + item + '" clip-begin="npt=0.000s" clip-end="npt=' + lFileTimeSeconds[0] + '.' + fileTimeMilliMicro + 's" id="a_' + str(z).zfill(4)  + '" />'+'\r\n')
                else:
                    fOutFile.write( '<audio src="' + item + '" clip-begin="npt=0.000s" clip-end="npt=' + str(15) + '.' + fileTimeMilliMicro + 's" id="a_' + str(z).zfill(4)  + '" />'+'\r\n')
                    zz=z+1
                    phraseSeconds = 15
                    while phraseSeconds <= lFileTime[z-1]-15:
                        fOutFile.write( '<audio src="' + item + '" clip-begin="npt='+ str(phraseSeconds)+ '.' + fileTimeMilliMicro + 's" clip-end="npt=' + str(phraseSeconds+15) + '.' + fileTimeMilliMicro + 's" id="a_' + str(zz).zfill(4)  + '" />'+'\r\n')
                        phraseSeconds +=15
                        zz +=1
                    fOutFile.write( '<audio src="' + item + '" clip-begin="npt='+ str(phraseSeconds)+ '.' + fileTimeMilliMicro + 's" clip-end="npt=' + lFileTimeSeconds[0] + '.' + fileTimeMilliMicro + 's" id="a_' + str(zz).zfill(4)  + '" />'+'\r\n')
                
                fOutFile.write( '</seq>'+'\r\n')
                fOutFile.write( '</par>'+'\r\n')
                fOutFile.write( '</seq>'+'\r\n')

                fOutFile.write( '</body>'+'\r\n')
                fOutFile.write('</smil>'+'\r\n')
                fOutFile.close


    def showDialogCritical(self,  errorMessage):
        QtGui.QMessageBox.critical(self, "Achtung", errorMessage)
    
    def showDebugMessage(self,  debugMessage):
        if self.app_debugMod == "yes":
            print debugMessage
    
    def actionQuit(self):
        QtGui.qApp.quit()
    
    def main(self):
        self.showDebugMessage( u"let's rock" )
        self.readConfig()
        self.progressBar.setValue(0)
        # Bhz in Combo
        for item in self.app_bhzItems:
            self.comboBoxCopyBhz.addItem(item)
        # Ausgabe in Comboncc:maxPageNormal
        prevYear = str( datetime.datetime.now().year -1 )
        currentYear = str( datetime.datetime.now().year )
        nextYear = str( datetime.datetime.now().year +1 )
        for item in self.app_prevAusgItems:
            self.comboBoxCopyBhzAusg.addItem(prevYear + "_" + item)
        for item in self.app_currentAusgItems:
            self.comboBoxCopyBhzAusg.addItem(currentYear + "_" + item)
        for item in self.app_nextAusgItems:
            self.comboBoxCopyBhzAusg.addItem(nextYear + "_" + item)
        # Trenner in Combo
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
        # Vorbelegung Checkboxen
        self.checkBoxCopyBhzIntro.setChecked(True)
        self.checkBoxCopyBhzAusgAnsage.setChecked(True)
        self.checkBoxCopyID3Change.setChecked(True)
        self.checkBoxCopyBitrateChange.setChecked(True)
        # Vorbelegung spinboxen
        self.spinBoxEbenen.setValue(1)
        self.spinBoxPages.setValue(0)
        self.show()
 
if __name__=='__main__':
    app = QtGui.QApplication(sys.argv)
    dyc = DaisyCopy()
    dyc.main()
    app.exec_()
    # This shows the interface we just created. No logic has been added, yet.
