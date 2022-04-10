import os
import aqt
from aqt import mw, gui_hooks
from aqt.utils import showInfo, qconnect
from anki.importing import TextImporter
from aqt.qt import *
from aqt.utils import showInfo, tooltip

## Import list from an old addon I wanted to make




def conditionalLoad():
    if checkFolder():
        #checks for folder

        importOnLoad()
        #if present the imports
    else:
        showInfo("There was nothing to import")
        #if not then shows text

def checkFolder():
    absolute_path = os.path.abspath(__file__)
    #gets the current file path

    Folder = "Library/Application Support/Anki2/addons21/PPT_Anki/__init__.py"
    #String to remove form file path

    Folder = absolute_path.replace(Folder, "")
    #Remove the string from the folder path to get the user directory

    projecttestfolder = os.path.join(Folder, ".anki-PPT")
    #create variable to test for .anki folder presence

    if not os.path.exists(projecttestfolder):
        #if not present

        os.mkdir(projecttestfolder)
        #creates folder if not present
        return False
        #returns False i.e. the folder does not exist

    else:
        return True
        #returns true i.e. it does


def importOnLoad():
    #Checks existance of .anki folder

    absolute_path = os.path.abspath(__file__)
    #gets file of __init__.py

    addonFile = "Library/Application Support/Anki2/addons21/PPT_Anki/__init__.py"
    #truncates file path using method form checkFolder()

    Folder = absolute_path.replace(addonFile, ".anki-PPT/")
    #trucates

    

    list = os.listdir(Folder) 
    #lists the files/folders within .anki
    
    number_files = len(list)
    #converts the list to a count
    if (number_files < 2): 
        # Checks there are any folders

        showInfo("There was nothing to import")
        #Shows there was nothing to import

        return 0
        # Leaves if there arent 
    
    card_type = ""
    target_deck = ""
    conf = open(Folder + "conf.csv", 'r')
    conf_text = conf.readline()
    conf.close()
    card_type = conf_text.split(',')[0]
    target_deck = conf_text.split(',')[1]
    deck_id = mw.col.decks.id(target_deck)
    #Imports cards into PowerPoint_Import deck - change this if you want
    
    file_1 =  Folder + "slides.csv"
    #creates new filepath
    file = file_1.encode(encoding = 'UTF-8', errors = 'strict')
    #encodes filepath in uniode for textImport

    mw.col.decks.select(deck_id)
    #selects deck
       
    notetype = mw.col.models.by_name(card_type)
    #selects note type
    deck = mw.col.decks.get(deck_id)
    deck['mid'] = notetype['id']
    mw.col.decks.save(deck)
    # and puts cards in the last deck used by the note type
    mw.col.set_aux_notetype_config(
        notetype["id"], "lastDeck", deck_id
    )
    # import into the collection
    ti = TextImporter(mw.col, file)
    ti.initMapping()
    ti.run()
    #Imports file

    os.remove(file_1)
    os.remove(Folder + 'conf.csv')
    #removes file

gui_hooks.profile_did_open.append(importOnLoad)

action = QAction(aqt.mw)
action.setText("Import from PowerPoint")
aqt.mw.form.menuTools.addAction(action)
action.triggered.connect(conditionalLoad)
