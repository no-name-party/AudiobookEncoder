# -*- coding: utf8 -*-

__author__ = "Dennis Oesterle"
__copyright__ = "Copyright 2014, Dennis Oesterle"
__license__ = "CC BY-NC-SA - Attribution-NonCommercial-ShareAlike"
__appname__ = "Audiobook Encoder"
__version__ = "0.926b"
__email__ = "dennis@no-name-party.de"

#===============================================================================

# TODO
# Python 3.6
# Pyside 2 or PyQt5
# unconverted cover to trash
# delete audiobook -> move files to trash
# title and album name from meta data
# most rescent album names in list
# fix sorting files
# cleanup

#IMPMORTS
import sip
sip.setapi('QString', 2) # change pyqt4 api to version 2 to avoid qstrings
from PyQt4 import QtGui, QtCore
import sys
import xml.etree.ElementTree as ET
import AudioFileTools
import os
import shutil
import signal
import subprocess
import multiprocessing
from datetime import date
reload(sys)
sys.setdefaultencoding("UTF8")

#GUI============================================================================

class AudiobookEncoderMainWindow(QtGui.QMainWindow):
    """Create the main window"""

    def __init__(self):
        QtGui.QMainWindow.__init__(self)

        self.setWindowTitle("{0} - {1}" .format(__appname__, __version__))
        self.resize(1200, 720)
        self.setFixedSize(1200, 700)
        # move to center of screen 
        self.move(QtGui.QApplication.desktop().screen().rect().center() - self.rect().center())

        self.setWindowIcon(QtGui.QIcon(_script_dir + "/icons/icon32.png"))

        # fonts
        self.Header1Font = QtGui.QFont("Arial", 25, QtGui.QFont.Bold)
        self.Header2Font = QtGui.QFont("Arial", 15, QtGui.QFont.Bold)
        self.Header3Font = QtGui.QFont("Arial", 12, QtGui.QFont.Bold)
        self.Normal1Font = QtGui.QFont("Arial", 12, QtGui.QFont.Normal)

        # menubar
        menubar = self.menuBar()
        main_menu = menubar.addMenu("")

        about_m = main_menu.addAction("About")
        about_m.triggered.connect(lambda: AboutMenu(self))

        pref_m = main_menu.addAction("Preferences")
        pref_m.triggered.connect(lambda: OptionMenu(self, [self.Header2Font, self.Header3Font]))

        doc_m = main_menu.addAction("Website")
        doc_m.setMenuRole(QtGui.QAction.ApplicationSpecificRole) # this adds it to main app menu
        doc_m.triggered.connect(lambda: self.openBrowser("http://no-name-party.de/audiobook-encoder/"))

        doc_m = main_menu.addAction("Help")
        doc_m.setMenuRole(QtGui.QAction.ApplicationSpecificRole) # this adds it to main app menu
        doc_m.triggered.connect(lambda: self.openBrowser("http://no-name-party.de/audiobook-encoder/"))

        # tag fields
        self.book_name = TextBox(self, font=self.Header3Font, name="Title", position=[15, 15, 310, 60], lineEditSize=[300, 25])

        self.author = TextBox(self, font=self.Header3Font, name="Author", position=[15, 70, 270, 60], lineEditSize=[265, 25])
        #self.author.lineEdit.textEdited.connect(lambda: AudioFileTools.saveToXml(self.book_name.lineEdit.text(), xml_cache_root, _cache_dir, author = self.author.lineEdit.text()))
        self.author.lineEdit.textEdited.connect(lambda: AudioFileTools.saveToXml(self.book_name.lineEdit.text(), xml_cache_root, _cache_dir, author = [True, self.author.lineEdit.text()]))

        # author preset button
        self.apreset = QtGui.QPushButton("", self)
        self.apreset.setGeometry(290, 93, 25, 25)
        self.apreset.setIcon(QtGui.QIcon(_script_dir + "/icons/clip.png"))
        self.apreset.setToolTip("Save or set an author preset.")
        # preset pop up menu
        self.popMenu = QtGui.QMenu()
        self.apreset.setMenu(self.popMenu)
        self.popMenu.aboutToShow.connect(self.createMenuPreset)
        self.apreset.setStyleSheet("QPushButton::menu-indicator { image: none; }")

        self.comment = TextBox(self, font = self.Header3Font, name = "Comment", position = [15, 125, 310, 60], lineEditSize = [300, 25])
        self.comment.lineEdit.textEdited.connect(lambda: AudioFileTools.saveToXml(self.book_name.lineEdit.text(), xml_cache_root, _cache_dir, comment = [True, self.comment.lineEdit.text()]))

        self.genre = TextBox(self, font = self.Header3Font, name = "Genre", position = [15, 180, 310, 60], lineEditSize = [300, 25])
        self.genre.lineEdit.textEdited.connect(lambda: AudioFileTools.saveToXml(self.book_name.lineEdit.text(), xml_cache_root, _cache_dir, genre = [True, self.genre.lineEdit.text()]))

        # cover drag and drop zone
        self.cover_field = CoverWidget(self, self.Header2Font, self.book_name)

        # export field
        self.export_path = TextBox(self, font = self.Header3Font, name = "Export Destination", position = [15, 555, 270, 60], lineEditSize = [265, 25])

        self.add_path = QtGui.QPushButton("", self)
        self.add_path.setGeometry(290, 578, 25, 25)
        self.add_path.setIcon(QtGui.QIcon(_script_dir + "/icons/folder.png"))
        self.add_path.setToolTip("Change export path.")

        if not AudioFileTools.readOptionsXml(xml_options_root, "expdest"):
            AudioFileTools.saveOptionsXml(xml_options_root, _options_dir, _desktop_folder, "expdest" )

        self.export_path.lineEdit.setText(AudioFileTools.readOptionsXml(xml_options_root, "expdest"))
        self.add_path.clicked.connect(lambda: self.addPath(file_dialog = True))
        self.export_path.lineEdit.textEdited.connect(lambda: self.addPath(file_dialog = False))

        # build options quality
        self.build_options = QtGui.QComboBox(self)
        self.build_options.setGeometry(12, 615, 230, 25)
        self.build_options.addItems(qualityPresets)
        self.build_options.currentIndexChanged.connect(lambda: AudioFileTools.saveToXml(self.book_name.lineEdit.text(), xml_cache_root, _cache_dir, qualityPresets, quality = [True, self.build_options.currentIndex()]))
        self.build_options.setToolTip("Change quality of an audiobook.")

        # export yes or no button
        self.exportState = QtGui.QCheckBox("Activate", self)
        self.exportState.setGeometry(250, 615, 150, 25)
        self.exportState.setFont(self.Header3Font)
        self.exportState.setStyleSheet("QCheckBox {color: grey;}")
        self.exportState.setToolTip("Enable / disable exporting.")

        # book list
        widgets = [self.book_name, self.author, self.comment, self.genre, self.cover_field, self.export_path, self.build_options, self.exportState]
        self.book_list = TreeWidget(self, self.Header2Font, widgets)
        self.book_list.setFocus()

        # book rename event connections
        self.book_name.lineEdit.textEdited.connect(lambda: self.book_list.changeItemName(self.book_name.lineEdit.text()))
        self.exportState.stateChanged.connect(self.activateExport)

        #bt
        self.options = QtGui.QPushButton("Options", self)
        self.options.setGeometry(10, 645, 160, 35)
        self.options.setToolTip("...")
        self.options.clicked.connect(lambda: OptionMenu(self, [self.Header2Font, self.Header3Font]))

        self.export_audio = QtGui.QPushButton("Export", self)
        self.export_audio.setGeometry(165, 645, 160, 35)
        self.export_audio.setToolTip("Batch export all audiobooks.")
        self.export_audio.clicked.connect(self.onExport)

        # check for removed files
        self.removedFiles = AudioFileTools.checkFiles(xml_cache_root)
        if self.removedFiles:
            self.errorLog(self.removedFiles, mesg = False, error_log = False)

    def addPath(self, file_dialog = True):
        """let the user choose a path on os"""

        if file_dialog:
            if os.path.exists(self.export_path.lineEdit.text()):
                open_path = self.export_path.lineEdit.text()
            else:
                open_path = AudioFileTools.readOptionsXml(xml_options_root, "expdest")

            self.filepath = QtGui.QFileDialog.getExistingDirectory(self, 'Add Path', open_path)

            if self.filepath:
                self.export_path.lineEdit.setText(self.filepath + "/")
                AudioFileTools.saveToXml(self.book_name.lineEdit.text(), xml_cache_root, _cache_dir, destination = self.filepath + "/")
        else:
            AudioFileTools.saveToXml(self.book_name.lineEdit.text(), xml_cache_root, _cache_dir, destination = self.export_path.lineEdit.text())

    def createMenuPreset(self):
        """create preset item menu"""

        self.popMenu.clear()
        self.popMenu.addAction("Save / Change", self.savePreset)
        self.popMenu.addSeparator()
        for each_preset in sorted(AudioFileTools.readPresetAuthor(xml_options_root, _options_dir, menu = True)):
            self.popMenu.addAction(each_preset, lambda each_preset = each_preset: self.setPreset(each_preset))

    def savePreset(self):
        """save author preset"""

        author = AudioFileTools.readFromXml(xml_cache_root, self.book_name.lineEdit.text(), author = True)
        comment = AudioFileTools.readFromXml(xml_cache_root, self.book_name.lineEdit.text(), comments = True)
        genre = AudioFileTools.readFromXml(xml_cache_root, self.book_name.lineEdit.text(), genre = True)
        destination = AudioFileTools.readFromXml(xml_cache_root, self.book_name.lineEdit.text(),destination = True)
        quality = AudioFileTools.readFromXml(xml_cache_root, self.book_name.lineEdit.text(), quality = True)
        AudioFileTools.savePresetAuthor(xml_options_root, _options_dir, [author, comment, genre, destination, quality])

    def setPreset(self, preset_name):
        """set author preset and save to cache xml"""

        preset = AudioFileTools.readPresetAuthor(xml_options_root, _options_dir, menu = False, menu_item = preset_name)
        sel_books = self.book_list.selectedItems()

        for each in sel_books:
            if each.text(0).lower().endswith(".mp3"):
                selected_book =  each.parent().text(0)
            else:
                selected_book = each.text(0)

            self.author.lineEdit.setText(preset[0])
            AudioFileTools.saveToXml(selected_book, xml_cache_root, _cache_dir, author = [True, preset[0]])

            self.comment.lineEdit.setText(preset[1])
            AudioFileTools.saveToXml(selected_book, xml_cache_root, _cache_dir, comment = [True, preset[1]])

            self.genre.lineEdit.setText(preset[2])
            AudioFileTools.saveToXml(selected_book, xml_cache_root, _cache_dir, genre = [True, preset[2]])

            self.export_path.lineEdit.setText(preset[3])
            AudioFileTools.saveToXml(selected_book, xml_cache_root, _cache_dir, destination = preset[3])

            self.build_options.setCurrentIndex(int(preset[4].split(";")[0]))
            AudioFileTools.saveToXml(selected_book, xml_cache_root, _cache_dir, qualityPresets, quality = [True, int(preset[4].split(";")[0])])

    def errorLog(self, error_msg = False, mesg = False, error_log = True):
        """error log for the user"""

        self.mbox = QtGui.QMessageBox()

        if not error_log:
            if error_msg[0] == 1:
                mesg = " file is missing!                                                                       "
            else:
                mesg = " files are missing!                                                                     "

            self.mbox.setText(str(error_msg[0]) + mesg)
            self.mbox.setDetailedText(str(error_msg[1]))

        else:
            self.mbox.setText(mesg)

        self.mbox.open()

    def onExport(self):
        """export audiobooks"""

        books = AudioFileTools.readFromXml(xml_cache_root, all_audiobooks = True)
        somthing_to_export = [each_book for each_book in books  if int(AudioFileTools.readFromXml(xml_cache_root, each_book, export = True)[0])]

        if somthing_to_export:
            missingFiles = AudioFileTools.checkFiles(xml_cache_root)
            if missingFiles:
                self.errorLog(missingFiles, mesg = False, error_log = False)
            else:
                exportUi = LogUi(self, label = "Exporting...", msg = "Please do not move any files until this process is finished.", counter_msg = "0 / {0}".format(len(somthing_to_export)),fontSize = [20, 12], progress = True, close = False)
                AudioFileTools.exportAction(xml_cache_root, xml_options_root, _script_dir, exportUi)
        else:
            LogUi(self, label = "Exporting Error!", msg = "No audiobooks are activ.", fontSize = [20, 12])

    def activateExport(self):
        """active audiobooks for export"""

        if self.book_list.selectedItems():
            for each_item in self.book_list.selectedItems():
                selected_book = each_item.text(0)

                if selected_book.lower().endswith(".mp3"):
                    each_item = each_item.parent()
                    selected_book = each_item.text(0)

                AudioFileTools.saveToXml(selected_book, xml_cache_root, _cache_dir, export = [True, int(self.exportState.isChecked())])
                text_colors = AudioFileTools.readFromXml(xml_cache_root, selected_book, export = True)[1:]
                self.exportState.setStyleSheet("QCheckBox {color: " + text_colors[1] + ";}")

                # change text color of treewidget item
                each_item.setTextColor(0, QtGui.QColor(text_colors[0]))
                each_item.setTextColor(1, QtGui.QColor(text_colors[0]))

                # change highlight color in treewidget
                if text_colors[0] == "red":
                    highlight_color = QtCore.Qt.red
                else:
                    highlight_color = QtGui.QColor(56, 117, 215,)

                color_pallet = self.book_list.palette()
                color_pallet.setColor(QtGui.QPalette.Highlight, highlight_color)
                self.book_list.setPalette(color_pallet)

                self.book_list.countActivatedAudiobooks()

    def openBrowser(self, page):
            QtGui.QDesktopServices.openUrl(QtCore.QUrl(page))

class CoverWidget(QtGui.QLabel):
    """drag and drop zone for covers"""

    def __init__(self, parent, font, widget):
        super(CoverWidget, self).__init__(parent)

        self.setText("Drag and Drop <br> Cover Artwork")
        self.setToolTip("Double click to delete cover artwork.")
        self.setFont(font)
        self.setGeometry(15, 245, 300, 300)
        self.setAlignment(QtCore.Qt.AlignCenter)
        self.setStyleSheet("QLabel {border: 3px dashed grey; border-radius: 15px; font-size: 20px; color: grey;}")

        self.setAcceptDrops(True)

        self.cover = QtGui.QPixmap()

        self.titleWidget = widget

        self.warning = QtGui.QLabel("", parent)
        self.warning.setGeometry(300, 510, 50, 50)
        self.warning.setFont(font)
        self.warning.setStyleSheet("QLabel {color: crimson;}")
        self.warning.hide()

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            super(CoverWidget, self).dragEnterEvent(event)

    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                # everything will be converted to png and deleted
                if self.titleWidget.lineEdit.text():
                    audiobook_name = self.titleWidget.lineEdit.text()
                    cover_path = AudioFileTools.addCover(url.path(), xml_cache_root, _cache_dir, audiobook_name, save_xml = True)
                else:
                    cover_path = AudioFileTools.addCover(url.path(), save_xml = False)

                self.cover.load(cover_path)
                self.setPixmap(self.cover.scaled(280, 280, aspectRatioMode = QtCore.Qt.KeepAspectRatio))

                self.imageWarning(image_ratio=AudioFileTools.resizeCover(cover_path, info=True))

                event.acceptProposedAction()

        else:
            super(CoverWidget, self).dropEvent(event)

    def mouseDoubleClickEvent(self, event):
        self.deleteCoverImage()

    def contextMenuEvent(self, event):
        self.menu = QtGui.QMenu(self)

        action_resize = QtGui.QAction("Scale", self)
        self.menu.addAction(action_resize)

        if self.titleWidget.lineEdit.text():
            audiobook_name = self.titleWidget.lineEdit.text()
            cover_path = AudioFileTools.readFromXml(xml_cache_root, audiobook_name, cover = True)
            if cover_path:
                action_resize.triggered.connect(lambda: self.resizeReplaceCover(cover_path))

        action_del = QtGui.QAction("Delete", self)
        self.menu.addAction(action_del)
        action_del.triggered.connect(self.deleteCoverImage)

        self.menu.popup(QtGui.QCursor.pos())

    def resizeReplaceCover(self, cover_path):
        """resize and replace cover"""

        resize_completed = AudioFileTools.resizeCover(cover_path, scale=True)
        if resize_completed:
            self.cover.load(cover_path)
            self.setPixmap(self.cover.scaled(280, 280, aspectRatioMode = QtCore.Qt.KeepAspectRatio))
            self.imageWarning(image_ratio=AudioFileTools.resizeCover(cover_path, info=True))


    def imageWarning(self, image_ratio):
        """check if the image ratio is 1"""

        if not image_ratio[0] == 1.0:
            self.warning.setText("!")
            self.warning.setToolTip("Size: <br>" + str(image_ratio[1]) + " * " + str(image_ratio[2]))
            self.warning.show()
        else:
             self.warning.hide()

    def deleteCoverImage(self):
        audiobook_name = self.titleWidget.lineEdit.text()
        AudioFileTools.deleteCover(audiobook_name, xml_cache_root, _cache_dir)
        self.setText("Drag and Drop <br> Cover Artwork")
        self.warning.hide()


# QTreeWidget
class TreeWidget(QtGui.QTreeWidget):
    """list for all loaded audiobooks"""

    def __init__(self, parent, font, widgets):
        super(TreeWidget, self).__init__(parent)
        self.setAcceptDrops(True)
        self.setDragDropMode(QtGui.QAbstractItemView.InternalMove)
        self.setGeometry(360, 37, 825, 635)
        self.setHeaderLabels(["Audiobook", "Duration"])
        self.header().resizeSection(0, 720)
        self.header().setStretchLastSection(True)
        self.setHeaderHidden(False)
        self.setSelectionMode(QtGui.QAbstractItemView.ExtendedSelection)

        # this text is hidden, when the first book is dropped
        self.item_list_helper = QtGui.QLabel("Drag and Drop <br> Audiobooks", self)
        self.item_list_helper.setGeometry(350, 290, 150, 45)
        self.item_list_helper.setFont(font)
        self.item_list_helper.setAlignment(QtCore.Qt.AlignCenter)
        self.item_list_helper.setStyleSheet("QLabel {font-size: 20px; color: grey;}")

        # widgets
        self.titleWidget = widgets[0]
        self.authorWidget = widgets[1]
        self.commentWidget = widgets[2]
        self.genreWidget = widgets[3]
        self.coverWidget = widgets[4]
        self.export_pathWidget = widgets[5]
        self.buildOptionWidget = widgets[6]
        self.exportStateWidget = widgets[7]

        self.itemSelectionChanged.connect(self.changeWidgets)
        self.itemDoubleClicked.connect(self.openFolder)

        self.setToolTip("Press the Backspace key to delete audiobooks.")

        # read at startup
        self.readAtStartFromXml()

    def changeWidgets(self):
        if len(self.selectedItems()) == 1:
            selected_book = self.selectedItems()[0].text(0)

            if selected_book.lower().endswith(".mp3"):
                selected_book = self.selectedItems()[0].parent().text(0)

            if not selected_book.lower().endswith(".mp3"):
                self.titleWidget.lineEdit.setText(AudioFileTools.readFromXml(xml_cache_root, selected_book, title = True))
                self.authorWidget.lineEdit.setText(AudioFileTools.readFromXml(xml_cache_root, selected_book, author = True))
                self.commentWidget.lineEdit.setText(AudioFileTools.readFromXml(xml_cache_root, selected_book, comments = True))
                self.genreWidget.lineEdit.setText(AudioFileTools.readFromXml(xml_cache_root, selected_book, genre = True))
                self.export_pathWidget.lineEdit.setText(AudioFileTools.readFromXml(xml_cache_root, selected_book, destination = True))
                self.buildOptionWidget.setCurrentIndex(int(AudioFileTools.readFromXml(xml_cache_root, selected_book, quality = True)[0]))

                self.activateAudiobook(changeWidget=True)

                #cover
                if AudioFileTools.readFromXml(xml_cache_root, selected_book, cover = True):
                    cover_path = AudioFileTools.readFromXml(xml_cache_root, selected_book, cover = True)

                    if not AudioFileTools.checkCover(xml_cache_root, cover_path):
                        self.coverWidget.setText("Cover Artwork <br> not found")
                        self.coverWidget.warning.hide()
                    else:
                        self.coverWidget.cover.load(cover_path)
                        self.coverWidget.setPixmap(self.coverWidget.cover.scaled(280, 280, aspectRatioMode = QtCore.Qt.KeepAspectRatio))
                        self.coverWidget.imageWarning(image_ratio=AudioFileTools.resizeCover(cover_path, info=True))
                else:
                    self.coverWidget.setText("Drag and Drop <br> Cover Artwork")
                    self.coverWidget.warning.hide()

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            super(TreeWidget, self).dragEnterEvent(event)

    def dragMoveEvent(self, event):
        super(TreeWidget, self).dragMoveEvent(event)

    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            count = 1

            # check if it already exists in cache
            all_books = AudioFileTools.readFromXml(xml_cache_root, all_audiobooks = True)
            new_b = []

            if all_books:
                for each in all_books:
                    # create a list of all new books
                    if each.startswith("New Audiobook "):
                        new_b.append(each)

                # get the numbers
                b_count = [int(c.split("k ")[1]) for c in new_b]
                # continue numbering with highest number
                if b_count:
                    count = sorted(b_count)[-1] + 1

            audiobook_title = "New Audiobook " + str(count)

            # create a new xml cache 
            new_xml_book = AudioFileTools.createElementsXml(audiobook_title, xml_cache_root, xml_options_root, qualityPresets, _cache_dir)
            files_xml = new_xml_book.getchildren()[8]

            # write all files to xml, title, path, length            
            for each_path in event.mimeData().urls():
                path_to_file =  each_path.path()

                # get all files and checks if folder or files are dropped
                file_info = AudioFileTools.getFiles(path_to_file)

                # save all to xml
                if file_info:
                    AudioFileTools.saveDragXml(file_info, files_xml, xml_cache_root, _cache_dir)
                else:
                   pass

            # get the total time
            AudioFileTools.getTotalDuration(new_xml_book.getchildren()[5], files_xml, xml_cache_root, _cache_dir)
            self.addNewItems(new_xml_book)

        else:
            super(TreeWidget, self).dropEvent(event)

    def addNewItems(self, new_xml_book = False, draggedItems = True, audiobook_name = False):
        if draggedItems:
            # read xml
            item_titel = new_xml_book.getchildren()[0].text
        elif not draggedItems:
            item_titel = audiobook_name

        item_length = AudioFileTools.readFromXml(xml_cache_root, item_titel, length = True)

        self.item = QtGui.QTreeWidgetItem()
        self.item.setText(0, item_titel)
        self.item.setText(1, item_length)
        self.addTopLevelItem(self.item)
        self.item.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsDropEnabled)

        self.item.setToolTip(0, "Press the Backspace key to delete audiobooks.")

        files = AudioFileTools.readFromXml(xml_cache_root, item_titel, files = True)

        if files:
            for each_file in files:
                file_title = each_file[0] + "    {0}".format(each_file[2])
                file_length = each_file[1]

                sub_item = QtGui.QTreeWidgetItem(self.item, [file_title, file_length])
                # disable drag and drop etc
                sub_item.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsSelectable) # | QtCore.Qt.ItemIsDragEnabled)
                itemTip = "Double click to open directory.\nPress Space key to play / pause track.\n" + each_file[2]
                sub_item.setToolTip(0, itemTip)

            if draggedItems:
                self.item.setExpanded(int(AudioFileTools.readOptionsXml(xml_options_root, "expand")))

            #change text color
            text_color = AudioFileTools.readFromXml(xml_cache_root, item_titel, export = True)
            self.item.setTextColor(0, QtGui.QColor(text_color[1]))
            self.item.setTextColor(1, QtGui.QColor(text_color[1]))
            self.countActivatedAudiobooks()

            self.setCurrentItem(self.item)
            self.item_list_helper.hide()

        else:
            QtGui.QMessageBox.warning(self, "Message", "Please drop a valid file format! (mp3)", QtGui.QMessageBox.Ok)
            AudioFileTools.deleteAudiobook(item_titel, xml_cache_root, _cache_dir)
            root_item = self.invisibleRootItem()
            root_item.removeChild(self.item)

            self.countActivatedAudiobooks()

    def readAtStartFromXml(self):
        list_of_all_audiobooks = AudioFileTools.readFromXml(xml_cache_root, None, all_audiobooks = True)
        self.countActivatedAudiobooks(list_of_all_audiobooks)

        for each_book in list_of_all_audiobooks:
            self.addNewItems(draggedItems = False, audiobook_name = each_book)

    def changeItemName(self, new_audiobook_name):
        if self.selectedItems():
            if new_audiobook_name:
                if new_audiobook_name.lower().endswith(".mp3"):
                    LogUi(self, label = "Name Error!", msg = "You can't add \".mp3\" to audobook names.", fontSize = [20, 12])
                    new_audiobook_name = new_audiobook_name[:-4]

                selected_item = self.selectedItems()[0]
                selected_book = self.selectedItems()[0].text(0)

                if selected_book.lower().endswith(".mp3"):
                    selected_item = selected_item.parent()
                    selected_book =  selected_item.text(0)

                #write to xml
                changed_book_name = AudioFileTools.saveToXml(selected_book, xml_cache_root, _cache_dir, new_audiobook_name = new_audiobook_name)
                selected_item.setText(0, changed_book_name)

    def activateAudiobook(self, changeWidget=False, activate=False):
        """change action: activate audiobook for export"""

        if changeWidget:
            selected_book = self.selectedItems()[0].text(0)
            if selected_book.lower().endswith(".mp3"):
                selected_book =  self.selectedItems()[0].parent().text(0)

            currentExportState = AudioFileTools.readFromXml(xml_cache_root, selected_book, export = True)
            self.exportStateWidget.setChecked(int(currentExportState[0]))
            self.exportStateWidget.setStyleSheet("QCheckBox {color: " + currentExportState[2] + ";}")
        else:
            if activate:
                self.exportStateWidget.setChecked(0)
                self.exportStateWidget.setChecked(1)
            else:
                self.exportStateWidget.setChecked(1)
                self.exportStateWidget.setChecked(0)

    def countActivatedAudiobooks(self, audiobooks = False):
        """counts active audiobooks"""

        if not audiobooks:
            audiobooks = AudioFileTools.readFromXml(xml_cache_root, None, all_audiobooks = True)

        count_books = len(audiobooks)
        count_active_books = len([each_b for each_b in audiobooks if int(AudioFileTools.readFromXml(xml_cache_root, each_b, export = True)[0])])
        self.setHeaderLabels(["Audiobook ({0}/{1}) ".format(count_active_books, count_books), "Duration"])

    def keyPressEvent(self, event):

        if event.key() == QtCore.Qt.Key_Backspace:
            # delete files
            self.deleteFiles()

            if not self.invisibleRootItem().childCount():
                self.item_list_helper.show()

        if event.key() == QtCore.Qt.Key_Right:
            # open tree item
            if self.selectedItems():
                for each_item in self.selectedItems():
                    each_item.setExpanded(True)

        if event.key() == QtCore.Qt.Key_Left:
            # jump to root tree item and close item
            selected_item = self.selectedItems()
            self.clearSelection()
            if selected_item:
                itemsToExpand = []
                for each_item in selected_item:
                    selected_item_name = each_item.text(0)
                    self.setCurrentItem(each_item)

                    if selected_item_name.lower().endswith(".mp3"):
                        self.setItemSelected(each_item, False)
                        itemsToExpand.append(each_item.parent())
                    else:
                        itemsToExpand.append(each_item)
                        each_item.setExpanded(False)

                    for each_item in itemsToExpand:
                        self.setItemSelected(each_item, True)

        if event.key() == QtCore.Qt.Key_Up:
            if self.selectedItems():
                selected_item = self.selectedItems()[0]
                item_up = self.itemAbove(selected_item)
                self.setCurrentItem(item_up)
            else:
                item_count = self.invisibleRootItem().childCount()
                self.setCurrentItem(self.topLevelItem(item_count -1))

        if event.key() == QtCore.Qt.Key_Down:
            if self.selectedItems():
                selected_item = self.selectedItems()[0]
                item_down = self.itemBelow(selected_item)
                self.setCurrentItem(item_down)
            else:
                self.setCurrentItem(self.topLevelItem(0))

        if event.key() == QtCore.Qt.Key_Space:
            # play / pause file
            self.playFile()

        if event.modifiers() == QtCore.Qt.ControlModifier:
            # select all audiobooks cmd+a
            if event.key() == QtCore.Qt.Key_A:
                self.clearSelection()
                for each_book in range(self.invisibleRootItem().childCount()):
                    self.setItemSelected(self.invisibleRootItem().child(each_book), True)

    def contextMenuEvent(self, event):

        if self.selectedItems():
            item_name = self.selectedItems()[-1].text(0)
            self.menu = QtGui.QMenu(self)

            actions = [["Open in Finder", self.openFolder], ["Play File", self.playFile], ["Activate", lambda: self.activateAudiobook(activate=True)], ["Deactivate", self.activateAudiobook], ["Delete", self.deleteFiles]]

            if int(AudioFileTools.readOptionsXml(xml_options_root, "playfile")):
                actions[1][0] = "Stop File"
            if not item_name.lower().endswith(".mp3"):
                del actions[0:2]
            if item_name.lower().endswith(".mp3"):
                del actions[2:4]

            for each_action in actions:
                action = QtGui.QAction(each_action[0], self)
                self.menu.addAction(action)
                try:
                    action.triggered.connect(each_action[1])
                except:
                    pass

            self.menu.popup(QtGui.QCursor.pos())

    def openFolder(self):
        """double click and open a finder where the files are"""

        item_clicked =  self.selectedItems()[0].text(0).split("    ")

        if item_clicked[0].lower().endswith(".mp3"):
            selected_book =  self.selectedItems()[0].parent().text(0)
            files =  AudioFileTools.readFromXml(xml_cache_root, selected_book, files = True)

            # search for the clicked item in xml file and return path
            file_path = [os.path.dirname(each_file[2]) + "/" for each_file in files if each_file[2] == item_clicked[1]][0]

            if os.path.exists(file_path):
                subprocess.Popen(["open", file_path])
            else:
                folder_err = LogUi(self, label = "Folder Error!", msg = "Can't find folder:\n" + file_path)

    def deleteFiles(self):
        """delete files"""

        root_item = self.invisibleRootItem()

        for each_item in self.selectedItems():
            item_name = each_item.text(0)

            if not item_name.lower().endswith(".mp3"):
                # delete the hole audiobook
                # xml delete
                AudioFileTools.deleteAudiobook(item_name, xml_cache_root, _cache_dir)
                # item delete
                (each_item.parent() or root_item).removeChild(each_item)

                # clear all widgets
                self.titleWidget.lineEdit.setText("")
                self.authorWidget.lineEdit.setText("")
                self.commentWidget.lineEdit.setText("")
                self.genreWidget.lineEdit.setText("")
                self.coverWidget.setText("Drag and Drop <br> Cover Artwork")
                self.coverWidget.warning.hide()
                self.export_pathWidget.lineEdit.setText("")
                self.buildOptionWidget.setCurrentIndex(0)
                self.exportStateWidget.setChecked(1)
                self.exportStateWidget.setStyleSheet("QCheckBox {color: grey;}")
            else:
                # delete mp3 file
                self.setCurrentItem(each_item)
                audiobook_name = each_item.parent().text(0)
                item_name =  item_name.split("    ")[1]
                AudioFileTools.deleteFile(audiobook_name, item_name, xml_cache_root, _cache_dir)

                # get the new total duration of audiobook
                for each_book in xml_cache_root.getchildren():
                    if audiobook_name == each_book.getchildren()[0].text:
                        new_duration = AudioFileTools.getTotalDuration(each_book.getchildren()[5], each_book.getchildren()[8], xml_cache_root, _cache_dir)
                        each_item.parent().setText(1, new_duration)

                (each_item.parent() or root_item).removeChild(each_item)

        self.countActivatedAudiobooks()

    def playFile(self):
        """play/stop file"""

        if len(self.selectedItems()) == 1:
            item_sel =  self.selectedItems()[0].text(0).split("    ")
            if item_sel[0].lower().endswith(".mp3"):
                playfile = int(AudioFileTools.readOptionsXml(xml_options_root, "playfile"))

                selected_book = self.selectedItems()[0].parent().text(0)
                files =  AudioFileTools.readFromXml(xml_cache_root, selected_book, files = True)

                # search for the clicked item in xml file and return path
                file_path = [each_file[2] for each_file in files if each_file[2] == item_sel[1]][0]

                if os.path.exists(file_path):
                    afplayCMD = ["afplay", file_path]
                    if not playfile:
                        # play sound and set pid
                        process = subprocess.Popen(afplayCMD, stdout = subprocess.PIPE, stderr = subprocess.STDOUT, preexec_fn = os.setsid)
                        AudioFileTools.saveOptionsXml(xml_options_root, _options_dir, process.pid, "playfile")
                    else:
                        try:
                            os.killpg(playfile, signal.SIGTERM)
                            AudioFileTools.saveOptionsXml(xml_options_root, _options_dir, "0", "playfile")
                        except OSError:
                            process = subprocess.Popen(afplayCMD, stdout = subprocess.PIPE, stderr = subprocess.STDOUT, preexec_fn = os.setsid)
                            AudioFileTools.saveOptionsXml(xml_options_root, _options_dir, process.pid, "playfile")


class OptionMenu(QtGui.QDialog):
    """option menus"""

    def __init__(self, parent, font = None):
        super(OptionMenu, self).__init__(parent)
        self.resize(440, 620)

        self.optionHeader = QtGui.QLabel("Options", self)
        self.optionHeader.setGeometry(20, 20, 200, 20)
        self.optionHeader.setFont(font[0])
        self.optionHeader.setStyleSheet("QLabel {color: grey;}")

        self.itunesCheck = QtGui.QCheckBox("Add to iTunes Library", self)
        self.itunesCheck.move(20, 50)
        self.itunesCheck.setFont(font[1])
        self.itunesCheck.setStyleSheet("QCheckBox {color: grey;}")
        self.itunesCheck.setToolTip("This will add exported audiobooks to iTunes.")
        self.itunesCheck.setChecked(int(AudioFileTools.readOptionsXml(xml_options_root, "itunes")))
        self.itunesCheck.stateChanged.connect(lambda: AudioFileTools.saveOptionsXml(xml_options_root, _options_dir, int(self.itunesCheck.isChecked()), "itunes"))

        self.shotdownPc = QtGui.QCheckBox("Shutdown computer after export", self)
        self.shotdownPc.move(20, 80)
        self.shotdownPc.setFont(font[1])
        self.shotdownPc.setStyleSheet("QCheckBox {color: grey;}")
        self.shotdownPc.setChecked(int(AudioFileTools.readOptionsXml(xml_options_root, "shutdown")))
        self.shotdownPc.stateChanged.connect(lambda: AudioFileTools.saveOptionsXml(xml_options_root, _options_dir, int(self.shotdownPc.isChecked()), "shutdown" ))

        self.expandBook = QtGui.QCheckBox("Expand Audiobook", self)
        self.expandBook.move(20, 110)
        self.expandBook.setFont(font[1])
        self.expandBook.setStyleSheet("QCheckBox {color: grey;}")
        self.expandBook.setToolTip("Files are visible.")
        self.expandBook.setChecked(int(AudioFileTools.readOptionsXml(xml_options_root, "expand")))
        self.expandBook.stateChanged.connect(lambda: AudioFileTools.saveOptionsXml(xml_options_root, _options_dir, int(self.expandBook.isChecked()), "expand" ))

        self.taskbox = QtGui.QSpinBox(self)
        self.taskbox.setRange(1, 8)
        self.taskbox.move(20, 140)
        self.tbox_label = QtGui.QLabel("Task Size <small>(max. {0} recommended on your system)</small>".format(multiprocessing.cpu_count()), self)
        self.tbox_label.setGeometry(72, 142, 250, 20)
        self.tbox_label.setFont(font[1])
        self.tbox_label.setStyleSheet("QLabel {color: grey;}")
        self.taskbox.setValue(int(AudioFileTools.readOptionsXml(xml_options_root, "task")))
        self.taskbox.valueChanged.connect(lambda: AudioFileTools.saveOptionsXml(xml_options_root, _options_dir, self.taskbox.value(), "task" ))
        self.taskbox.setToolTip("Increase the task size to export more audiobooks parallel. This may slow down your system.")

        self.default_path = TextBox(self, font = font[1], name = "Default Export Destination", position = [20, 175, 270, 60], lineEditSize = [265, 25])
        self.add_path = QtGui.QPushButton("", self)
        self.add_path.setGeometry(295, 198, 25, 25)
        self.add_path.setIcon(QtGui.QIcon(_script_dir + "/icons/folder.png"))
        self.add_path.setToolTip("Change default export path.")
        self.default_path.lineEdit.setText(AudioFileTools.readOptionsXml(xml_options_root, "expdest"))
        self.add_path.clicked.connect(lambda: self.addPath(file_dialog = True))
        self.default_path.lineEdit.textEdited.connect(lambda: self.addPath(file_dialog = False))

        # author presets
        self.authorHeader = QtGui.QLabel("Delete Author Presets", self)
        self.authorHeader.setGeometry(20, 240, 200, 20)
        self.authorHeader.setFont(font[1])
        self.authorHeader.setStyleSheet("QLabel {color: grey;}")

        self.authorList = QtGui.QListWidget(self)
        self.authorList.setGeometry(21, 265, 400, 330)

        for each_preset in sorted(AudioFileTools.readPresetAuthor(xml_options_root, _options_dir, menu = True)):
            self.preset_item = QtGui.QListWidgetItem(each_preset)
            self.preset_item.setToolTip("Hit Backspace to delete a preset.")
            self.authorList.addItem(self.preset_item)

        self.open()

    def addPath(self, file_dialog = True):
        """let the user choose a path on os"""
        if file_dialog:
            if os.path.exists(AudioFileTools.readOptionsXml(xml_options_root, "expdest")):
                open_path = AudioFileTools.readOptionsXml(xml_options_root, "expdest")
            else:
                open_path = _desktop_folder

            self.filepath = QtGui.QFileDialog.getExistingDirectory(self, 'Add Path', open_path)

            if self.filepath:
                self.default_path.lineEdit.setText(self.filepath + "/")
                AudioFileTools.saveOptionsXml(xml_options_root, _options_dir, self.filepath + "/", "expdest" )
        else:
            AudioFileTools.saveOptionsXml(xml_options_root, _options_dir, self.default_path.lineEdit.text(), "expdest" )

    def keyPressEvent(self, event):
        """delete items with del key"""

        if event.key() == QtCore.Qt.Key_Backspace:
            delete_item = self.authorList.selectedItems()[0]
            AudioFileTools.deletePresetAuthor(xml_options_root, _options_dir, delete_item.text())
            self.authorList.takeItem(self.authorList.row(delete_item))

        elif event.key() == QtCore.Qt.Key_Escape:
            self.close()

class AboutMenu(QtGui.QDialog):
    """about menu"""

    def __init__(self, parent):
        super(AboutMenu, self).__init__(parent)
        self.resize(440, 400)

        self.logo = QtGui.QLabel(self)
        self.logo.setPixmap(QtGui.QPixmap(_script_dir + "/icons/icon324.png"))
        self.logo.setAlignment(QtCore.Qt.AlignCenter)

        self.AboutHeader = QtGui.QLabel("<center>Audiobook <br>Encoder</center>", self)
        self.AboutHeader.setFont(QtGui.QFont("Arial", 40, QtGui.QFont.Bold))
        self.AboutHeader.setStyleSheet("QLabel {color: grey;}")

        self.version = QtGui.QLabel("<center>Version {0}</center>".format(__version__), self)
        self.version.setFont(QtGui.QFont("Arial", 13, QtGui.QFont.Bold))
        self.version.setStyleSheet("QLabel {color: grey;}")

        self.copyright = QtGui.QLabel("<center>Copyright 2013-{0}<br><a href=\"mailto:{1}\"><span style=\"color:grey;\">{2}</span></a></center>".format(date.today().year, __email__, __author__), self)
        self.copyright.setFont(QtGui.QFont("Arial", 12, QtGui.QFont.Light))
        self.copyright.setStyleSheet("QLabel {color: grey;}")
        self.copyright.linkActivated.connect(lambda: QtGui.QDesktopServices.openUrl(QtCore.QUrl("mailto:" + __email__)))

        self.license = QtGui.QLabel("<center><b>License</b><br>{0}<br>{1}</center>".format(__license__.split(" - ")[0], __license__.split(" - ")[1]), self)
        self.license.setFont(QtGui.QFont("Arial", 12, QtGui.QFont.Light))
        self.license.setStyleSheet("QLabel {color: grey;}")

        # additional creitds
        self.addCreditsLabel = QtGui.QLabel("<center><b>Additional Credits</b>", self)
        self.addCreditsLabel.setFont(QtGui.QFont("Arial", 12, QtGui.QFont.Light))
        self.addCreditsLabel.setStyleSheet("QLabel {color: grey;}")

        self.sub_python = QtGui.QLabel("<center><a href=\"http://www.python.org/\"><span style=\"color:grey;\">Python</span></a></center>", self)
        self.sub_python.setFont(QtGui.QFont("Arial", 12, QtGui.QFont.Light))
        self.sub_python.linkActivated.connect(lambda: QtGui.QDesktopServices.openUrl(QtCore.QUrl("http://www.python.org/")))

        self.sub_pyside = QtGui.QLabel("<center><a href=\"http://www.riverbankcomputing.co.uk/software/pyqt/download\"><span style=\"color:grey;\">PyQt4</span></a></center>", self)
        self.sub_pyside.setFont(QtGui.QFont("Arial", 12, QtGui.QFont.Light))
        self.sub_pyside.linkActivated.connect(lambda: QtGui.QDesktopServices.openUrl(QtCore.QUrl("http://qt-project.org/wiki/PySide")))

        self.sub_mutagen = QtGui.QLabel("<center><a href=\"http://code.google.com/p/mutagen/\"><span style=\"color:grey;\">Mutagen</span></a></center>", self)
        self.sub_mutagen.setFont(QtGui.QFont("Arial", 12, QtGui.QFont.Light))
        self.sub_mutagen.linkActivated.connect(lambda: QtGui.QDesktopServices.openUrl(QtCore.QUrl("http://code.google.com/p/mutagen/")))

        self.sub_pil = QtGui.QLabel("<center><a href=\"http://www.pythonware.com/products/pil/\"><span style=\"color:grey;\">PIL</span></a></center>", self)
        self.sub_pil.setFont(QtGui.QFont("Arial", 12, QtGui.QFont.Light))
        self.sub_pil.linkActivated.connect(lambda: QtGui.QDesktopServices.openUrl(QtCore.QUrl("http://www.pythonware.com/products/pil/")))

        self.sub_qt_faststart = QtGui.QLabel("<center><a href=\"https://github.com/danielgtaylor/\"><span style=\"color:grey;\">qt-faststart</span></a></center>", self)
        self.sub_qt_faststart.setFont(QtGui.QFont("Arial", 12, QtGui.QFont.Light))
        self.sub_qt_faststart.linkActivated.connect(lambda: QtGui.QDesktopServices.openUrl(QtCore.QUrl("https://github.com/danielgtaylor/qtfaststart")))

        self.sub_abbinder = QtGui.QLabel("<center><a href=\"http://bluezbox.com/audiobookbinder/abbinder.html\"><span style=\"color:grey;\">abbinder</span></a></center>", self)
        self.sub_abbinder.setFont(QtGui.QFont("Arial", 12, QtGui.QFont.Light))
        self.sub_abbinder.linkActivated.connect(lambda: QtGui.QDesktopServices.openUrl(QtCore.QUrl("http://bluezbox.com/audiobookbinder/abbinder.html")))

        # layout
        self.overAllLayout = QtGui.QFormLayout(self)
        self.overAllLayout.setVerticalSpacing(10)

        self.additinalCreitdsLayout = QtGui.QFormLayout()
        self.additinalCreitdsLayout.setVerticalSpacing(0)

        # adding widgets to layout
        Qlables = [self.logo, self.AboutHeader, self.version, self.copyright, self.license, self.addCreditsLabel, self.additinalCreitdsLayout]
        sub_credit_widgets = [self.addCreditsLabel, self.sub_python, self.sub_pyside, self.sub_mutagen, self.sub_pil, self.sub_qt_faststart, self.sub_abbinder]
        map(lambda each : self.additinalCreitdsLayout.addRow(each), sub_credit_widgets)
        map(lambda each : self.overAllLayout.addRow(each), Qlables)

        self.setLayout(self.overAllLayout)

        self.open()

class LogUi(QtGui.QDialog):
    """Error logging window"""

    def __init__(self, parent, label = None, msg = None, counter_msg = False, fontSize = [15, 12], progress = False, close = True):
        super(LogUi, self).__init__(parent)
        self.resize(440, 200)

        self.optionHeader = QtGui.QLabel(label, self)
        self.optionHeader.setGeometry(20, 20, 200, 50)
        self.optionHeader.setFont(QtGui.QFont("Arial", fontSize[0], QtGui.QFont.Bold))
        self.optionHeader.setStyleSheet("QLabel {color: grey;}")

        self.msg = QtGui.QLabel(msg, self)
        self.msg.setGeometry(20, 50, 400, 50)
        self.msg.setFont(QtGui.QFont("Arial", fontSize[1], QtGui.QFont.Bold))
        self.msg.setStyleSheet("QLabel {color: grey;}")

        # disable user input to close window during export
        self.want_to_close = close

        if progress:
            self.progressbar = QtGui.QProgressBar(self)
            self.progressbar.setGeometry(20, 125, 400, 100)
            self.progressbar.setMinimum(1)
            self.progressbar.setMaximum(100)

        if counter_msg:
            # counter for export
            self.counter = QtGui.QLabel(counter_msg, self)
            self.counter.setGeometry(210, 110, 100, 50)
            self.counter.setFont(QtGui.QFont("Arial", fontSize[1], QtGui.QFont.Bold))
            self.counter.setStyleSheet("QLabel {color: grey;}")

        self.open()

    def keyPressEvent(self, event):
        if self.want_to_close:
            if event.key() == QtCore.Qt.Key_Escape:
                self.close()

class TextBox:
    """textbox with header"""

    def __init__(self, parent, font, name = "Name", position = [0, 0, 0, 0], lineEditSize = [0, 0], tip = ""):
        self.combobox = QtGui.QWidget(parent)
        self.combobox.setGeometry(position[0], position[1], position[2], position[3])
        self.label = QtGui.QLabel(name, self.combobox)
        self.label.setGeometry(1, 1, 200, 20)
        self.label.setFont(font)
        self.label.setStyleSheet("QLabel {color: grey;}")
        self.lineEdit = QtGui.QLineEdit(self.combobox)
        self.lineEdit.setGeometry(1, 23, lineEditSize[0], lineEditSize[1])
        if tip:
            self.lineEdit.setToolTip(tip)

#===============================================================================

def globalVars():
    global _home_dir, _desktop_folder, _script_dir, xml_cache_root, xml_options_root, _cache_dir, _options_dir, qualityPresets
    _home_dir = QtCore.QDir.homePath()
    _desktop_folder = _home_dir + "/Desktop/"
    _script_dir = QtCore.QDir.currentPath()
    _cache_dir =  _script_dir + "/cache.xml"
    _options_dir = _script_dir + "/options.xml"

    external_dir = _home_dir + "/.audiobookEncoder"

    if not os.path.exists(external_dir):
        os.makedirs(external_dir)

    if not os.path.exists(external_dir + "/cache.xml"):
        shutil.copy2(_cache_dir, external_dir)

    if not os.path.exists(external_dir + "/options.xml"):
        shutil.copy2(_options_dir, external_dir)

    _cache_dir =  external_dir + "/cache.xml"
    _options_dir = external_dir + "/options.xml"

    # xml table auslesen
    xml_cache_root = ET.parse(_cache_dir).getroot()
    xml_options_root = ET.parse(_options_dir).getroot()

    qualityPresets = ["96 Kbps, Stereo, 48 kHz", "128 Kbps, Stereo, 48 kHz", "256 Kbps, Stereo, 48 kHz", "320 Kbps, Stereo, 48 kHz"]

def main():
    globalVars()
    app = QtGui.QApplication(sys.argv)
    aEUi = AudiobookEncoderMainWindow()
    aEUi.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()