#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""send Lotus Notes mail"""

__author__ = "子風"
__copyright__ = "Copyright 2016, Sun All rights reserved"
__version__ = "1.0.0"

import os, uuid
import itertools as it

from win32com.client import DispatchEx
import pywintypes # for exception

def send_mail(subject, body_text,sendto, copyto=None, blindcopyto=None, attach=None):
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

    document = notesDatabase.CreateDocument()
    document.ReplaceItemValue("Form","Memo")
    document.ReplaceItemValue("Subject", subject)

    # assign random uid because sometimes Lotus Notes tries to reuse the same one
    uid = str(uuid.uuid4().hex)
    document.ReplaceItemValue('UNIVERSALID', uid)

    # "SendTo" MUST be populated otherwise you get this error:
    # 'No recipient list for Send operation'
    document.ReplaceItemValue("SendTo", sendto)

    if copyto is not None:
        document.ReplaceItemValue("CopyTo", copyto)
    if blindcopyto is not None:
        document.ReplaceItemValue("BlindCopyTo", blindcopyto)

    # body
    body = document.CreateRichTextItem("Body")
    body.AppendText(body_text)

    # attachment
    if attach is not None:
        attachment = document.CreateRichTextItem("Attachment")
        for att in attach:
            attachment.EmbedObject(1454, "", att, "Attachment")

    # save in `Sent` view; default is False
    document.SaveMessageOnSend = True
    document.Send(False)


if __name__ == '__main__':
    subject = "test subject"
    body = "test body"
    sendto = ['to mail',]
    filesName = ['attachTest.txt', 'attachTest.txt']
    files = [os.path.abspath(file) for file in filesName]
    attachment = it.takewhile(lambda x: os.path.exists(x), files)

    send_mail(subject, body, sendto, attach=attachment)

    print('Done')