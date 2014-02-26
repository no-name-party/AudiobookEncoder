# -*- coding: utf8 -*-

#IMPMORTS
import shutil
import plistlib
import AudiobookEncoder
from setuptools import setup

APP = ["AudiobookEncoder.py"]
DATA_FILES = []
OPTIONS = {"argv_emulation": True,
           "iconfile": "icons/icon324.icns",
           "resources": ["cache.xml", "options.xml", "icons", "iTunesAddFile.scpt", "AudioFileTools.py", "qtfaststart_lic.txt"],
           "includes": ["sip", "PyQt4", "AudioFileTools"],
           "excludes": ["PyQt4.QtDesigner", "PyQt4.QtNetwork", "PyQt4.QtOpenGL", "PyQt4.QtScript", "PyQt4.QtSql", "PyQt4.QtTest",
                        "PyQt4.QtWebKit", "PyQt4.QtXml", "PyQt4.phonon", "PyQt4.QtSvg", "PyQt4.QtXmlPatterns",
                        "PyQt4.QtDeclarative", "PyQt.QtHelp", "PyQt4.QtMultimedia", "PyQt4.Qt", "PyQt4.ScriptTools"]}

setup(
    app=APP,
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)

# copy some binaries and edits plist
shutil.copy2("/Users/dennisbeckstein/Dropbox/scripts/python/AudiobookEncoder/abbinder", "/Users/dennisbeckstein/Dropbox/scripts/python/AudiobookEncoder/dist/AudiobookEncoder.app/Contents/Resources/")
shutil.copy2("/Users/dennisbeckstein/Dropbox/scripts/python/AudiobookEncoder/qtfaststart", "/Users/dennisbeckstein/Dropbox/scripts/python/AudiobookEncoder/dist/AudiobookEncoder.app/Contents/Resources/")

shutil.move("/Users/dennisbeckstein/Dropbox/scripts/python/AudiobookEncoder/dist/AudiobookEncoder.app", "/Users/dennisbeckstein/Dropbox/scripts/python/AudiobookEncoder/")

plist = plistlib.readPlist("/Users/dennisbeckstein/Dropbox/scripts/python/AudiobookEncoder/AudiobookEncoder.app/Contents/Info.plist")
plist_changes = [("CFBundleName", "Audiobook Encoder"),
                ("CFBundleDisplayName", "Audiobook Encoder"),
                ("NSHumanReadableCopyright", "Dennis Oesterle \n" + AudiobookEncoder.__license__),
                ("CFBundleShortVersionString", AudiobookEncoder.__version__),
                ("CFBundleVersion", AudiobookEncoder.__version__)]

for each in plist_changes:
    plist[each[0]] = each[1]

plistlib.writePlist(plist, "/Users/dennisbeckstein/Dropbox/scripts/python/AudiobookEncoder/AudiobookEncoder.app/Contents/Info.plist")

shutil.rmtree("/Users/dennisbeckstein/Dropbox/scripts/python/AudiobookEncoder/build")
shutil.rmtree("/Users/dennisbeckstein/Dropbox/scripts/python/AudiobookEncoder/dist")
