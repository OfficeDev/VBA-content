---
title: About the Order of Events
keywords: olfm10.chm3077138
f1_keywords:
- olfm10.chm3077138
ms.prod: outlook
ms.assetid: 04fc8833-90be-a7d1-d196-be10d1852aee
ms.date: 06/08/2017
---


# About the Order of Events

The following events occur in the order specified when a user completes an action.



|**Events**|**When**|
|:-----|:-----|
| **Open**|A form is opened to compose an item|
| **Send**,  **Write**,  **Close**|An item is sent|
| **BeforeAttachmentAdd**|Before an attachment is added to an item|
| **BeforeAttachmentPreview**|Before an attachment associated with an item is previewed|
| **AttachmentAdd**|An attachment has been added to an item|
| **BeforeAttachmentRead**|Before an attachment associated with an item is read from the file system, an attachment stream, or an  **[Attachment](attachment-object-outlook.md)** object|
| **AttachmentRead**|An attachment has been opened for reading|
| **BeforeAttachmentSave**|Before an attachment is saved into the Outlook item|
| **BeforeAttachmentWriteToTempFile**|Before an attachment associated with an item is written to a temporary file|
| **BeforeAutoSave**|Before the item is automatically saved by Outlook|
| **BeforeCheckNames**|Before Outlook starts resolving names in the recipients collection of the item against the address book, after the user explicitly uses the  **Check Names** command|
| **Write**,  **Close**|An item is posted|
| **Write**|An item is saved|
| **Read**,  **Open**|An item is opened in a folder|
| **Reply**|A user replies to an item's sender|
| **ReplyAll**|A user replies to an item's sender and all recipients|
| **Forward**|The newly-created item is passed to the procedure after the user selects the  **Forward** action for an item|
| **BeforeDelete**|Before Outlook deletes the item|
| **PropertyChange**|One of the item's standard properties is changed|
| **CustomPropertyChange**|One of the item's custom properties is changed|
| **CustomAction**|A user-defined action is initiated|
| **Unload**|Before an Outlook item is unloaded from memory, either programmatically or by user action|

The  **Click** event occurs only when you have defined it for a control in the Script Editor.


