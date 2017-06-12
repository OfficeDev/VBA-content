---
title: Application.BeforeFolderSharingDialog Event (Outlook)
keywords: vbaol11.chm447
f1_keywords:
- vbaol11.chm447
ms.prod: outlook
api_name:
- Outlook.Application.BeforeFolderSharingDialog
ms.assetid: e06257eb-f2d9-63cf-1220-dda55ee0ea14
ms.date: 06/08/2017
---


# Application.BeforeFolderSharingDialog Event (Outlook)

Occurs before the  **Sharing** dialog box is displayed for a selected **[Folder](folder-object-outlook.md)** object.


## Syntax

 _expression_ . **BeforeFolderSharingDialog**( **_FolderToShare_** , **_Cancel_** )

 _expression_ An expression that returns a **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FolderToShare_|Required| **Folder**|The  **Folder** object to be shared.|
| _Cancel_|Required| **Boolean**|(Not used in VBScript).  **False** when the event occurs. If the event procedure sets this argument to **True** , the dialog box is not displayed.|

## Remarks

This event provides an add-in with the capability of replacing the sharing user interface supplied by Outlook with a custom user interface. This event does not occur if a sharing message is programmatically created and displayed.


## See also


#### Concepts


[Application Object](application-object-outlook.md)

