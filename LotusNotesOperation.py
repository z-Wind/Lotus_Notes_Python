#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""LotusNotes operation"""

__author__ = "子風"
__copyright__ = "Copyright 2016, Sun All rights reserved"
__version__ = "1.0.0"

import uuid
from win32com.client import DispatchEx
import pywintypes # for exception
import os

# 可用 檔案 -> 資料庫 -> 屬性 查到
def getDatabase(server, filePath, password):
    # Connect
    notesSession = DispatchEx('Lotus.NotesSession')
    try:
        notesSession.Initialize(password)
        notesDatabase = notesSession.GetDatabase(server, filePath)
        if not notesDatabase.IsOpen:
            try:
                notesDatabase.Open()
            except pywintypes.com_error:
                print( 'could not open database: {}'.format(db_name) )
        return notesDatabase
    except pywintypes.com_error:
        raise Exception('Cannot access database using %s on %s' % (filePath, server))

        
def roughlyShow(db):
    with open("roughlyshow.txt", "wb") as f:
        for view in db.Views:
            os.system("clear")
            print(view.Name)
            f.write("▇▇▇▇▇▇▇▇▇▇▇▇▇▇▇▇▇▇▇▇▇▇▇▇▇▇▇▇▇▇▇▇▇▇▇▇▇▇▇▇▇▇▇▇▇▇▇▇▇▇▇▇▇▇▇▇▇▇▇▇▇▇▇\n".encode())
            f.write("★view\n".encode())
            f.write((view.Name + "\n").encode())
            document = view.GetFirstDocument()
            f.write("\n◆Items in First Document\n".encode())
            if document:
                for item in document.Items:
                    f.write("▂▂▂▂▂▂▂▂▂▂▂▂▂▂▂▂▂▂▂▂▂▂▂▂▂▂▂▂▂▂▂▂▂▂▂▂▂▂▂▂▂▂▂▂▂▂▂▂▂▂▂▂▂▂▂▂▂▂▂▂▂▂▂\n".encode())
                    f.write(("◎ " + item.Name + "\n").encode())
                    f.write((item.Text + "\n").encode())
                    # print("---item---> " + item.Name + " == " + item.Text)
            else:
                f.write(("No Document in " + view.Name + "\n").encode())
            
        
def printAllViews(db):
    with open("allViews.txt", "w") as f:
        for view in db.Views:
            if view.IsFolder:
                f.write('{:>17}{}\n'.format("★Folder : ", view.Name))
            elif view.IsCategorized:
                f.write('{:>17}{}\n'.format("Categorized : ", view.Name))   
            elif view.IsHierarchical:
                f.write('{:>17}{}\n'.format("Hierarchical : ", view.Name))
            elif view.IsCalendar:
                f.write('{:>17}{}\n'.format("Calendar : ", view.Name))
            elif view.IsDefaultView:
                f.write('{:>17}{}\n'.format("DefaultView : ", view.Name))            
            elif view.IsModified:
                f.write('{:>17}{}\n'.format("Modified : ", view.Name))
            elif view.IsPrivate:
                f.write('{:>17}{}\n'.format("Private : ", view.Name))
            elif view.IsConflict:
                f.write('{:>17}{}\n'.format("Conflict : ", view.Name))
            else:
                f.write('{:>17}{}\n'.format("◆Others : ", view.Name))
           

def makeDocumentGenerator(db, view):
    # Get the first document
    document = view.GetFirstDocument()
    # If the document exists,
    while document:
        # Yield it
        yield document
        # Get the next document
        document = view.GetNextDocument(document)
        
        
def printAllDocuments(db, view):
    for document in makeDocumentGenerator(db, view):
        # Get fields
        printAllItemName(document, True)


def printAllItemName(document, showContent=False):
    for i in document.Items:
        if showContent:
            print(i.Name + " == " + i.Text)
        else:
            print(i.Name)

            
def createDocumentAndSave(db, **item):
    document = db.CreateDocument()
    for key in item:
        document.ReplaceItemValue(key, item[key])
    
    document.Save( False, False )
    return document
    
    
if __name__ == '__main__':
    # 可用 檔案 -> 資料庫 -> 屬性 查到
    server = ''
    filePath = '.nsf'
    
    password = ''
    
    db = getDatabase(server, filePath, password)
    # printAllViews(db)
    # input("pause press any key to continue")
    
    # Get view
    # viewName = '($All)'
    # view = db.GetView(viewName)
    # if not view:
        # raise Exception('Folder "%s" not found' % viewName)
        
    # printAllDocuments(db, view)
    # roughlyShow(db)
    item = {"Form": "Memo", "Subject": "test subject", "SendTo":['xxmail',], "UNIVERSALID":str(uuid.uuid4().hex),}
    document = createDocumentAndSave(db, **item)
    
    # send email
    document.SaveMessageOnSend = True
    document.Send(False)
    
    
    