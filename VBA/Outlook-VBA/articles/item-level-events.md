---
title: Item-Level Events
keywords: olfm10.chm3077126
f1_keywords:
- olfm10.chm3077126
ms.prod: outlook
ms.assetid: 2f90bac7-4256-4675-c5b0-15cc2bdf9046
ms.date: 06/08/2017
---


# Item-Level Events



Item-level events occur when something happens to an item displayed in a form, such as when it's saved or opened or when a user-defined action is started.
Most often, item-level events are handled by Microsoft Visual Basic Scripting Edition (VBScript) code within the form itself.
Some events can be cancelled. That is, your event handler can prevent Microsoft Outlook from performing the default action associated with the event. For example, you can write an event handler for the  **Forward** event to prevent an item from being sent to recipients who are not on a list of approved recipients. Learn about [canceling an event](canceling-an-event.md).
The following table lists the item-level events supported by Outlook.


|**Event**|**Cancelable?**|**Description**|
|:-----|:-----|:-----|
| **AttachmentAdd**|No|Occurs when an attachment has been added to the item|
| **AttachmentRead**|No|Occurs when an attachment has been opened for reading|
| **AttachmentRemove**|No|Occurs when an attachment has been removed from an item|
| **BeforeAttachmentAdd**|Yes|Occurs before an attachment is added to an item|
| **BeforeAttachmentPreview**|Yes|Occurs before an attachment associated with an item is previewed|
| **BeforeAttachmentRead**|Yes|Occurs before an attachment associated with an item is read from the file system, an attachment stream, or an  **[Attachment](attachment-object-outlook.md)** object|
| **BeforeAttachmentSave**|Yes|Occurs just before an attachment is saved into the Outlook item|
| **BeforeAttachmentWriteToTempFile**|Yes|Occurs before an attachment associated with an item is written to a temporary file|
| **BeforeAutoSave**|Yes|Occurs before the item is automatically saved by Outlook|
| **BeforeCheckNames**|Yes|Occurs before Outlook starts resolving names in the recipients collection of the item|
| **BeforeDelete**|Yes|Occurs before Outlook deletes an item that has been opened in an inspector|
| **Close**|Yes|Occurs before Outlook closes the inspector displaying the item|
| **CustomAction**|Yes|Occurs before Outlook executes a custom action of an item|
| **CustomPropertyChange**|No|Occurs when a custom item property has changed|
| **Forward**|Yes|Occurs before Outlook executes the  **Forward** action of an item|
| **Open**|Yes|Occurs before Outlook opens an inspector to display the item|
| **PropertyChange**|No|Occurs when an item property has changed|
| **Read**|No|Occurs when an item is opened for editing by a user|
| **Reply**|Yes|Occurs before Outlook executes the  **Reply** action of an item|
| **ReplyAll**|Yes|Occurs before Outlook executes the  **Reply to All** action of an item|
| **Send**|Yes|Occurs before Outlook sends the item|
| **Unload**|No|Occurs before an Outlook item is unloaded from memory, either programmatically or by user action|
| **Write**|Yes|Occurs before Outlook saves the item in a folder|

