#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""read Lotus Notes mail"""

__author__ = "子風"
__copyright__ = "Copyright 2016, Sun All rights reserved"
__version__ = "1.0.0"

from win32com.client import DispatchEx
import pywintypes # for exception
# Import system modules
import datetime
# Import system modules
import os
import tempfile

def getTemporaryPath():
    temporaryIndex, temporaryPath = tempfile.mkstemp()
    os.close(temporaryIndex)
    return temporaryPath

# Given a document, return a list of attachment filenames and their contents
def extractAttachments(document):
    # Prepare
    attachmentPacks = []
    # For each item,
    for whichItem in range(len(document.Items)):
        # Get item
        item = document.Items[whichItem]
        # If the item is an attachment,
        if item.Name == '$FILE':
            # Prepare
            attachmentPath = getTemporaryPath()
            # Get the attachment
            fileName = item.Values[0]
            fileBase, separator, fileExtension = fileName.rpartition('.')
            attachment = document.GetAttachment(fileName)
            attachment.ExtractFile(attachmentPath)
            attachmentContent = open(attachmentPath, 'rb').read()
            os.remove(attachmentPath)
            # Append
            attachmentPacks.append((fileBase, fileExtension, attachmentContent))
    # Return
    return attachmentPacks

def makeDocumentGenerator(folderName):
    # Get folder
    folder = notesDatabase.GetView(folderName)
    if not folder:
        raise Exception('Folder "%s" not found' % folderName)
    # Get the first document
    document = folder.GetFirstDocument()
    # If the document exists,
    while document:
        # Yield it
        yield document
        # Get the next document
        document = folder.GetNextDocument(document)

if __name__ == '__main__':
    # Get credentials
    mailServer =  'Your Server'
    mailPath = 'Your database'
    mailPassword = 'Your password'
    # Connect
    notesSession = DispatchEx('Lotus.NotesSession')
    try:
        notesSession.Initialize(mailPassword)
        notesDatabase = notesSession.GetDatabase(mailServer, mailPath)
    except pywintypes.com_error:
        raise Exception('Cannot access mail using %s on %s' % (mailPath, mailServer))

    # print('Title:' + notesDatabase.Title)

    # Get a list of folders
    # for view in notesDatabase.Views:
    #     if view.IsFolder:
    #         print('view : ' + view.Name)

    for document in makeDocumentGenerator('($Inbox)'):
        # Get fields
        subject = document.GetItemValue('Subject')[0].strip()
        date = datetime.datetime(
            year=document.GetItemValue('PostedDate')[0].year,
            month=document.GetItemValue('PostedDate')[0].month,
            day=document.GetItemValue('PostedDate')[0].day,
            hour=document.GetItemValue('PostedDate')[0].hour,
            minute=document.GetItemValue('PostedDate')[0].minute,
            second=document.GetItemValue('PostedDate')[0].second,
            microsecond=document.GetItemValue('PostedDate')[0].microsecond,
            tzinfo=document.GetItemValue('PostedDate')[0].tzinfo)
        fromWhom = document.GetItemValue('From')[0].strip()
        toWhoms = document.GetItemValue('SendTo')
        body = document.GetItemValue('Body')[0].strip()
        form = document.GetItemValue('Form')
        attachmentPacks = extractAttachments(document)

        print('Form: {4}\nSubject : {0}\nFrom : {1}\nTo : {2}\nAttach File : {3}'.format(subject, fromWhom, toWhoms, len(extractAttachments(document)), form))
        print('===========================================================================')

    print('Done')