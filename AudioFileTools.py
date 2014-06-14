# -*- coding: utf8 -*- 

#IMPMORTS
import sip
sip.setapi('QString', 2)
from PyQt4 import QtCore
import os
import mutagen
from mutagen.mp4 import MP4
from mutagen.mp4 import MP4Cover
import datetime
import subprocess
import xml.etree.ElementTree as ET
import Image
import sys
import threading
import math
from multiprocessing.dummy import Pool
from Foundation import NSAutoreleasePool
reload(sys)
sys.setdefaultencoding("UTF8")

#file functions=================================================================

def getFiles(path_to_file):
    """get a list of all files and check if a folder or a file is dropped"""

    file_info = []
    # folder
    if os.path.isdir(path_to_file):
        for dirpath, dirnames, filenames in os.walk(path_to_file):
            if not dirpath.endswith("/"):
                dirpath = dirpath + "/"
            for filename in [f for f in filenames if f.lower().endswith("mp3")]:
                file_duration = getFileDuration(os.path.join(dirpath + filename))
                file_info.append(["FOLDER", filename, file_duration, (dirpath + filename)])

        return file_info
    
    # files
    if not os.path.isdir(path_to_file):
        if path_to_file.lower().endswith("mp3"):
            file_duration = getFileDuration(path_to_file)
            file_info = ["FILE", os.path.basename(path_to_file), file_duration, path_to_file]
    
            return file_info

def getFileDuration(path_to_audio):
    """get the file duration"""

    afinfo = ["afinfo", "-b", str(path_to_audio)]
    process = subprocess.Popen(afinfo, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
    stdout, stderr = process.communicate()

    # search by index for the seconds of file
    file_sec =  round(float(stdout[(stdout.index("----") + 4):(stdout.index(" sec"))]))
    duration = str(datetime.timedelta(seconds = file_sec))

    if len(duration.split(":")[0]) == 1:
        duration = "0" + duration

    return duration

def getTotalDuration(new_xml_book, audio_files_xml, xml_cache_root, _cache_dir):
    """get the total audiobook duration"""
    
    audiobook_duration = 0
    
    for each_file in audio_files_xml.getchildren():
        # read all files from xml
        time = each_file.getchildren()[1].text
        
        # split the string
        time_parts = time.split(":")
            
        h = int(time_parts[0])
        m = int(time_parts[1])
        s = int(time_parts[2])
     
        total_s = (h * 60 * 60) + (m * 60) + s
        
        # add all together    
        audiobook_duration += total_s
       
    audiobook_duration = str(datetime.timedelta(seconds = audiobook_duration))
     
    if len(audiobook_duration.split(":")[0]) == 1:
        audiobook_duration = "0" + audiobook_duration
    
    # delete day with a number  
    if "day" in audiobook_duration:
        days_to_hours = int(audiobook_duration.split("day")[0]) * 24
        old_hours = int(audiobook_duration.split(", ")[1].split(":")[0])
        new_hours = days_to_hours + old_hours
        m = audiobook_duration.split(", ")[1].split(":")[1]
        s = audiobook_duration.split(", ")[1].split(":")[2]
        audiobook_duration = "{0}:{1}:{2}" .format(new_hours, m, s)
    
    # save to xml
    new_xml_book.text = audiobook_duration
    ET.ElementTree(xml_cache_root).write(_cache_dir, encoding = "utf-8", xml_declaration = True, method = "xml")
    
    return audiobook_duration

def checkFiles(xml_cache_root):
    """check if all files are available"""

    all_books = readFromXml(xml_cache_root, all_audiobooks = True)
    all_files = []
    files_book = []
    removed_files = []

    for each_b in all_books:
        # get files from each book
        files_book = [each_file[2] for each_file in readFromXml(xml_cache_root, audiobook_name = each_b, files = True)]
        count = 0
        for each_file in files_book:
            if not os.path.exists(each_file):
                count += 1
        if count:
            removed_files.append([each_b, count])

    if removed_files:
        sum_files = sum([e[1] for e in removed_files])
        detail_err_string = "".join(["{0}: {1}\n" .format(e[0], e[1]) for e in removed_files])
        removed_files = [sum_files, detail_err_string]

    else:
        removed_files = False

    return removed_files

def checkCover(xml_cache_root, cover):
    """check if cover is available"""

    if not os.path.exists(cover):
        return False
    else:
        return True

#XML functions==================================================================

def saveDragXml(file_info, files_xml, xml_cache_root, _cache_dir):
    """on import save to xml and create some subelements"""
    
    if file_info[0][0] == "FOLDER":
        for each_file in file_info:
            _file = ET.SubElement(files_xml, "file")
            track_title = ET.SubElement(_file, "title")
            track_length = ET.SubElement(_file, "length")
            track_path = ET.SubElement(_file, "path")
            
            track_title.text = each_file[1]
            track_length.text = each_file[2]
            track_path.text = each_file[3]
            
    if file_info[0] == "FILE":
        _file = ET.SubElement(files_xml, "file")
        track_title = ET.SubElement(_file, "title")
        track_length = ET.SubElement(_file, "length")
        track_path = ET.SubElement(_file, "path")
        
        track_title.text = file_info[1]
        track_length.text = file_info[2]
        track_path.text = file_info[3]
    
    ET.ElementTree(xml_cache_root).write(_cache_dir, encoding = "utf-8", xml_declaration = True, method = "xml")

def createElementsXml(audiobook_title, xml_cache_root, xml_options_root, qualityPresets, _cache_dir):
    """create a new xml cache"""
    
    new_xml_audiobook = ET.SubElement(xml_cache_root, "audiobook")
    sub_tags = ["title", "author", "comments", "genre", "cover", "length", "destination", "quality", "files", "export"]
    
    for each_sub_tag in sub_tags:
        ET.SubElement(new_xml_audiobook, each_sub_tag)
    
    new_xml_audiobook.getchildren()[0].text = audiobook_title
    new_xml_audiobook.getchildren()[6].text = readOptionsXml(xml_options_root, "expdest")
    new_xml_audiobook.getchildren()[3].text = "Audiobook"
    new_xml_audiobook.getchildren()[7].text = str(0) + ";" + qualityPresets[0]
    new_xml_audiobook.getchildren()[9].text = "1,black,grey"

    ET.ElementTree(xml_cache_root).write(_cache_dir, encoding = "utf-8", xml_declaration = True, method = "xml")
    
    return new_xml_audiobook

def readFromXml(xml_cache_root, audiobook_name = None, all_audiobooks = None, title = None, author = None, comments = None, genre = None, cover = None, length = None, destination = None, quality = None, files = None, export = None):
    """read data from xml"""
    
    if all_audiobooks:
        list_audiobooks = []
        for each_book in xml_cache_root.getchildren():
            list_audiobooks.append(each_book.getchildren()[0].text)
        return list_audiobooks
        
    elif not all_audiobooks:
        for each_book in xml_cache_root.getchildren():               
            if audiobook_name == each_book.getchildren()[0].text:
                if title:
                    return each_book.getchildren()[0].text 
                 
                elif author:
                    return each_book.getchildren()[1].text
                
                elif comments:
                    return each_book.getchildren()[2].text
                
                elif genre:
                    return each_book.getchildren()[3].text
                
                elif cover:
                    return each_book.getchildren()[4].text
                
                elif length:
                    return each_book.getchildren()[5].text
                
                elif destination:
                    return each_book.getchildren()[6].text
                    
                elif quality:
                    return each_book.getchildren()[7].text.split(";")
                
                elif files:
                    files = each_book.getchildren()[8]
                    list_of_files = []
                    
                    for each_file in files:
                        list_of_files.append([unicode(each_file.getchildren()[0].text), unicode(each_file.getchildren()[1].text), unicode(each_file.getchildren()[2].text)])
                    
                    return list_of_files

                elif export:
                    return each_book.getchildren()[9].text.split(",")

def saveToXml(audiobook_name, xml_cache_root, _cache_dir, qualityPresets = False, new_audiobook_name = False, author = [False], comment = [False], genre = [False], destination = False, quality = [False], export = [False]):

    if new_audiobook_name:
        count_equal_name = 0
        for each_book in xml_cache_root.getchildren():
            if new_audiobook_name == each_book.getchildren()[0].text:
                count_equal_name += 1
                new_audiobook_name = new_audiobook_name + " " + str(count_equal_name)

        for each_book in xml_cache_root.getchildren():
            if audiobook_name == each_book.getchildren()[0].text:
                each_book.getchildren()[0].text = new_audiobook_name

        ET.ElementTree(xml_cache_root).write(_cache_dir, encoding = "utf-8", xml_declaration = True, method = "xml")

        return new_audiobook_name

    elif author[0]:
        for each_book in xml_cache_root.getchildren():
            if audiobook_name ==  each_book.getchildren()[0].text:
                each_book.getchildren()[1].text = author[1]

    elif comment[0]:
        for each_book in xml_cache_root.getchildren():
            if audiobook_name ==  each_book.getchildren()[0].text:
                each_book.getchildren()[2].text = comment[1]

    elif genre[0]:
        for each_book in xml_cache_root.getchildren():
            if audiobook_name ==  each_book.getchildren()[0].text:
                each_book.getchildren()[3].text = genre[1]

    elif destination:
        for each_book in xml_cache_root.getchildren():
            if audiobook_name ==  each_book.getchildren()[0].text:
                each_book.getchildren()[6].text = destination

    elif quality[0]:
        for each_book in xml_cache_root.getchildren():
            if audiobook_name ==  each_book.getchildren()[0].text:
                each_book.getchildren()[7].text = str(quality[1]) + ";" + qualityPresets[quality[1]]

    elif export[0]:
        for each_book in xml_cache_root.getchildren():
            if audiobook_name ==  each_book.getchildren()[0].text:
                if export[1]:
                    color = "black"
                    font_color = "grey"
                else:
                    color = "red"
                    font_color = "red"
                each_book.getchildren()[9].text = "{0},{1},{2}".format(str(export[1]), color, font_color)

    ET.ElementTree(xml_cache_root).write(_cache_dir, encoding = "utf-8", xml_declaration = True, method = "xml")

def deleteAudiobook(audiobook_name, xml_cache_root, _cache_dir):
    """delete the current selected audiobook"""
    
    for each_book in xml_cache_root.getchildren():
        if audiobook_name ==  each_book.getchildren()[0].text:
            xml_cache_root.remove(each_book)
            ET.ElementTree(xml_cache_root).write(_cache_dir, encoding = "utf-8", xml_declaration = True, method = "xml")

def deleteFile(audiobook_name, file_name, xml_cache_root, _cache_dir):
    """delete the current selected file"""

    for each_book in xml_cache_root.getchildren():
        if audiobook_name ==  each_book.getchildren()[0].text:
            for each_file in each_book.getchildren()[8]:
                if each_file[2].text == file_name:
                    each_book.getchildren()[8].remove(each_file)
                    ET.ElementTree(xml_cache_root).write(_cache_dir, encoding = "utf-8", xml_declaration = True, method = "xml")

def reorderFile(audiobook_name, xml_cache_root, _cache_dir):
    pass

def addCover(cover_url, xml_cache_root, _cache_dir, audiobook_name = None, save_xml = None):
    """converts everything in png, rgb and write to xml"""
    
    droped_cover = Image.open(cover_url)

    if not droped_cover.mode == "RGB":
        droped_cover.convert("RGB").save(cover_url)
        droped_cover = Image.open(cover_url)

    w, h = droped_cover.size
    size = [512, 512]
    
    # create png file / qt supports only png
    if not cover_url.endswith("png"):
        new_filename =  "/" + os.path.splitext(os.path.basename(os.path.abspath(cover_url)))[0] + ".png"
        new_cover_path = os.path.dirname(os.path.abspath(cover_url)) + new_filename 
        
        # change size
        if w or h > size:
            droped_cover.thumbnail(size, Image.ANTIALIAS) #@UndefinedVariable
        
        droped_cover.save(new_cover_path)
        cover_url = new_cover_path            
    
    # change size     
    if w or h > size:
        droped_cover.thumbnail(size, Image.ANTIALIAS) #@UndefinedVariable
        droped_cover.save(cover_url) 
    
    if audiobook_name:
        for each_book in xml_cache_root.getchildren():
            if audiobook_name ==  each_book.getchildren()[0].text:
                each_book.getchildren()[4].text = cover_url
    
    ET.ElementTree(xml_cache_root).write(_cache_dir, encoding = "utf-8", xml_declaration = True, method = "xml")
    
    return cover_url

def deleteCover(audiobook_name, xml_cache_root, _cache_dir):
    for each_book in xml_cache_root.getchildren():
            if audiobook_name ==  each_book.getchildren()[0].text:
                each_book.getchildren()[4].text = ""
    
    ET.ElementTree(xml_cache_root).write(_cache_dir, encoding = "utf-8", xml_declaration = True, method = "xml")

def saveOptionsXml(xml_options_root, _options_dir, widgetState = None, name = False):
    """write to option xml"""

    for each_opt in xml_options_root.getchildren():
        if each_opt.tag == name:
            each_opt.text = str(widgetState)

    ET.ElementTree(xml_options_root).write(_options_dir, encoding = "utf-8", xml_declaration = True, method = "xml")

def readOptionsXml(xml_options_root, name = False):
    """read from option xml"""

    for each_opt in xml_options_root.getchildren():
        if each_opt.tag == name:
            return unicode(each_opt.text)

def savePresetAuthor(xml_options_root, _options_dir, presets):
    """save preset author to xml"""

    option_xml = xml_options_root.getchildren()[4]

    new_preset = True

    # check if preset exists
    for each_preset in option_xml.getchildren():
        if each_preset.text == presets[0]:
            new_preset = False
            each_preset.text = presets[0] # author
            each_preset.getchildren()[0].text = presets[1] # comment
            each_preset.getchildren()[1].text = presets[2] # genre
            each_preset.getchildren()[2].text = presets[3] # destination
            each_preset.getchildren()[3].text = presets[4][0] + ";" + presets[4][1] # quality

    if new_preset:
        new_xml_preset = ET.SubElement(option_xml, "author")
        sub_tags = ["comments", "genre", "destination", "quality"]

        for each_sub_tag in sub_tags:
            ET.SubElement(new_xml_preset, each_sub_tag)

        new_xml_preset.text = presets[0]
        new_xml_preset.getchildren()[0].text = presets[1]
        new_xml_preset.getchildren()[1].text = presets[2]
        new_xml_preset.getchildren()[2].text = presets[3]
        new_xml_preset.getchildren()[3].text = presets[4][0] + ";" + presets[4][1]

    ET.ElementTree(xml_options_root).write(_options_dir, encoding = "utf-8", xml_declaration = True, method = "xml")

def deletePresetAuthor(xml_options_root, _options_dir, preset):
    """delete author preset from xml"""

    option_xml = xml_options_root.getchildren()[4]

    for each_preset in option_xml.getchildren():
        if each_preset.text == preset:
            option_xml.remove(each_preset)

    ET.ElementTree(xml_options_root).write(_options_dir, encoding = "utf-8", xml_declaration = True, method = "xml")

def readPresetAuthor(xml_options_root, _options_dir, menu = False, menu_item = False):
    """read preset informations from xml"""

    option_xml = xml_options_root.getchildren()[4]

    if menu:
        # get all preset items from xml
        menu_items = []
        for each_preset in option_xml.getchildren():
            menu_items.append(each_preset.text)

        return menu_items

    if not menu:
        # get the preset details
        for each_preset in option_xml.getchildren():
            if each_preset.text == menu_item:
                presets = [each_preset.text, each_preset.getchildren()[0].text, each_preset.getchildren()[1].text, each_preset.getchildren()[2].text, each_preset.getchildren()[3].text]

        return presets

#export functions===============================================================

def preExportActions(xml_cache_root):
    """return a list of exportable files"""

    def splitFiles(book_name, xml_cache_root, length = None):
        time_parts = length.split(":")    
        h = int(time_parts[0])
        
        files_book = [(each_file[1:]) for each_file in readFromXml(xml_cache_root, audiobook_name = book_name, files = True)]
        
        if h >= 13:
            paths = [[]]
            index = 0
            current_length = 0
            split_time = 45000 # 13 hours in seconds
            
            for each in files_book:
                if current_length <= split_time: 
                    # split the string
                    time_parts = each[0].split(":")
                    
                    h = int(time_parts[0])
                    m = int(time_parts[1])
                    s = int(time_parts[2])
                 
                    total_s = (h * 60 * 60) + (m * 60) + s
                    
                    new_length = current_length
                    new_length += total_s
                    
                    if new_length <= split_time:
                        # add all together    
                        current_length += total_s 
                        
                        paths[index].append(each[1])
                    
                    elif new_length >= split_time:
                        current_length = 0
                        index += 1
                        paths.append([each[1]])
                            
            else:
                return paths
       
        else:
            # all file paths for audiobooks that are shorter then 13h
            return [[each_path[1] for each_path in files_book]]

    books = readFromXml(xml_cache_root, all_audiobooks = True)
    exp_books = [[each_book, splitFiles(each_book, xml_cache_root, readFromXml(xml_cache_root, audiobook_name = each_book, length = True))] for each_book in books if int(readFromXml(xml_cache_root, each_book, export = True)[0])]

    return exp_books

def exportAction(xml_cache_root, xml_options_root, _script_dir, exportUi):
    """audiobook export"""

    # list of exportable books with files and lenght
    ls_books = preExportActions(xml_cache_root)

    def multiprocessExport(audiobook):
        #for audiobook in ls_books:
        NSAutoreleasePool.alloc().init() # we need this because this func is in another thread, qt can't take care of this

        # basic audiobook params
        a_name = audiobook[0]
        a_parts = len(audiobook[1])
        a_files = audiobook[1]
        ab_quality = str(readFromXml(xml_cache_root, a_name, quality = True)[1].split(", ")[0].split(" ")[0]) + "k"
        ar_quality = str(readFromXml(xml_cache_root, a_name, quality = True)[1].split(", ")[2].split(" ")[0]) + "000"
        a_dest = readFromXml(xml_cache_root, a_name, destination = True)

        # check if the export destination exists
        if not os.path.exists(a_dest):
            a_dest = QtCore.QDir.homePath() + "/Desktop/"

        # create parts
        for n, each in enumerate(a_files):
            # convert all files to one large string
            if len(each) >= 2:
                a_files_str = each
                parts = [1,1]

                if a_parts >= 2:
                    # adding Part to filename when exporting two files
                    # get tracknumbers
                    parts = [n + 1, a_parts]
                    final_name = a_dest + a_name + " Part " + str(parts[0]) + ".m4b"
                else:
                    final_name = a_dest + a_name + ".m4b"
            else:
                # for single files
                a_files_str = each
                parts = [1,1]
                final_name = a_dest + a_name + ".m4b"

            print parts, final_name, a_files_str

            # "-sv" = skip errors and go on with conversion, print some info on files being converted
            # "-b" = bitrate in KBps
            # "-r" = sample rate : 8000, 11025, 12000, 16000, 22050, 24000, 32000, (44100), 48000
            # "-c" = audio channels : 1/(2)
            abbinder = _script_dir + "/abbinder"
            abbinder_cmd = [abbinder, "-sv", "-b", ab_quality, "-r", ar_quality, "-c", "2", "-o", final_name] + a_files_str
            process = subprocess.Popen(abbinder_cmd, stdout = subprocess.PIPE, stderr = subprocess.STDOUT)
            stdout, stderr = process.communicate()

            postExportAction(xml_cache_root, xml_options_root, _script_dir, a_name, final_name, parts)

            if not parts == [1, 2]:
                # check if the book has more parts and finnish it before reenable the gui

                # progressbar
                cur_value = exportUi.progressbar.value()
                value_add_progressbar = int(math.ceil(float(100. / len(ls_books))))
                new_value = cur_value + value_add_progressbar

                if new_value >= 100:
                    new_value = 100

                exportUi.progressbar.setValue(new_value)
                counter_text = exportUi.counter.text()
                new_counter_text = "{0} / {1}".format((int(counter_text.split(" / ")[0]) + 1), counter_text.split(" / ")[1])
                exportUi.counter.setText(new_counter_text)

                # when done...
                if exportUi.progressbar.value() >= 100:
                    exportUi.want_to_close = True
                    exportUi.optionHeader.setText("Exporting... Done!")
                    exportUi.msg.setText("")

    if ls_books:
        def pool():
            # multiprocessing depending on cpu cores
            pool = Pool(int(readOptionsXml(xml_options_root, "task")))
            pool.map(multiprocessExport, ls_books)
            pool.close()
            pool.join()

        # no freezed ui with threads
        t = threading.Thread(target = pool)
        t.start()

def postExportAction(xml_cache_root, xml_options_root, _script_dir, audiobook_name, m4b_file, tracknumber):
    """adding meta tags, cover and itunes"""
    
    title = os.path.splitext(os.path.basename(m4b_file))[0]
    album = readFromXml(xml_cache_root, audiobook_name, title = True)
    artist = readFromXml(xml_cache_root, audiobook_name, author = True)
    comment = readFromXml(xml_cache_root, audiobook_name, comments = True)
    genre = readFromXml(xml_cache_root, audiobook_name, genre = True)
    tracknumber = u"{0}/{1}" .format(tracknumber[0], tracknumber[1])  
    coverImage = readFromXml(xml_cache_root, audiobook_name, cover = True)

    tags = mutagen.File(m4b_file, easy = True)
    # set custom tag content
    if title:
        tags["title"] = title
    if album:
        tags["album"] = album
    if artist:
        tags["artist"] = artist
    if comment:
        tags["comment"] = comment
    if genre:
        tags["genre"] = genre
    if tracknumber:
        tags["tracknumber"] = tracknumber
    
    tags.save()
    
    if coverImage:
        addCoverToM4B(m4b_file, coverImage)

    qtfaststart_dir = _script_dir + "/qtfaststart"
    qtfaststart_cmd = [qtfaststart_dir, m4b_file]

    process = subprocess.Popen(qtfaststart_cmd, stdout = subprocess.PIPE, stderr = subprocess.STDOUT)
    stdout, stderr = process.communicate()

    # adding files to itunes
    addToItunes(m4b_file, xml_options_root, _script_dir)
    shutdownAfterExport(xml_options_root, _script_dir)


def addCoverToM4B(filename, albumart):
    """add cover to m4a"""
    
    audio = MP4(filename)
    data = open(albumart, "rb").read()
        
    covr = []
    if albumart.endswith("png"):
        covr.append(MP4Cover(data, MP4Cover.FORMAT_PNG))
    else:
        covr.append(MP4Cover(data, MP4Cover.FORMAT_JPEG))

    audio.tags["covr"] = covr
    
    audio.save() 
    
    return True

def addToItunes(path, xml_options_root, _script_dir):
    """add file to itunes with applescript"""

    if int(readOptionsXml(xml_options_root, "itunes")):
        # check in options 
        appleScript = _script_dir + "/iTunesAddFile.scpt"
        appleScriptCMD = ["osascript", appleScript, path]
    
        process = subprocess.Popen(appleScriptCMD, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
        stdout, stderr = process.communicate()

def shutdownAfterExport(xml_options_root, _script_dir):
    """shutdown the computer after export"""
    
    if int(readOptionsXml(xml_options_root, "shutdown")):
        # check in options 
        osascript_cmd = ["osascript", "-e", "tell application \"System Events\" to shut down"]
    
        process = subprocess.Popen(osascript_cmd, stdout = subprocess.PIPE, stderr = subprocess.STDOUT)
        stdout, stderr = process.communicate()