---
title: EncryptionProvider.ShowSettings Method (Office)
keywords: vbaof11.chm327009
f1_keywords:
- vbaof11.chm327009
ms.prod: office
api_name:
- Office.EncryptionProvider.ShowSettings
ms.assetid: 9e66ee97-54d5-9b09-ff22-810b92e63125
ms.date: 06/08/2017
---


# EncryptionProvider.ShowSettings Method (Office)

Used to display a dialog of the encryption settings for the current document.


## Syntax

 _expression_. **ShowSettings**( **_SessionHandle_**, **_ParentWindow_**, **_ReadOnly_**, **_Remove_** )

 _expression_ An expression that returns a **EncryptionProvider** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SessionHandle_|Required|**Long**|The ID of the current session.|
| _ParentWindow_|Required|**IUnknown**|Specifies the window that is called to display the encryption settings.|
| _ReadOnly_|Required|**Boolean**|Specifies whether you want the user to be able to change the encryption settings.|
| _Remove_|Required|**Boolean**|If True the encryption for a document will be removed during the next save operation.|

## Remarks

This method can only be called on an already encrypted document. You can use this method in your COM add-in to display whatever user experience you like based on the user'smissions. For example, in a pure encryption scenario, you can display a dialog box to change the document's password. In a rights management scenario, you can decide whether to show a dialog box for changing permissions or show the user's permissions.


## See also


#### Concepts


[EncryptionProvider Object](encryptionprovider-object-office.md)
#### Other resources


[EncryptionProvider Object Members](encryptionprovider-members-office.md)

