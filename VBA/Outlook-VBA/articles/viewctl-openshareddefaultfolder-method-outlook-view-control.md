---
title: ViewCtl.OpenSharedDefaultFolder Method (Outlook View Control)
ms.prod: outlook
ms.assetid: 989d4a15-8aa6-4bc1-855f-1a4b2898ec35
ms.date: 06/08/2017
---


# ViewCtl.OpenSharedDefaultFolder Method (Outlook View Control)

Displays a specified user's default folder in the control.


## Version Information

 **Version Added:** Outlook 2010


## Syntax

 _expression_.  **OpenSharedDefaultFolder** **_(bstrRecipient, FolderType)_**

 _expression _ A variable that represents a **ViewCtl** object.


## Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _bstrRecipient_|Required| **String**|The owner of the folder. The string must contain a display name or alias that can be resolved to a valid recipient.|
| _FolderType_|Required| **OlxDefaultFolders**|The type of folder. Can be one of the following  **OlxDefaultFolders** constants: **olxFolderDeletedItems**(3),  **olxFolderOutbox**(4),  **olxFolderSentMail**(5),  **olxFolderInbox**(6),  **olxFolderCalendar**(9),  **olxFolderContacts**(10),  **olxFolderJournal**(11),  **olxFolderNotes**(12),  **olxFolderTasks**(13), or  **olxFolderDrafts**(16).|

## Remarks

An error occurs if the user running the control does not have permission to access the specified folder. 


