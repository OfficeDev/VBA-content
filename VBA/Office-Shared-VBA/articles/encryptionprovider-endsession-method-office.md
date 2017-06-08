---
title: EncryptionProvider.EndSession Method (Office)
keywords: vbaof11.chm327005
f1_keywords:
- vbaof11.chm327005
ms.prod: office
api_name:
- Office.EncryptionProvider.EndSession
ms.assetid: ce19f32e-a680-9d84-97d8-67d0f2d3b139
ms.date: 06/08/2017
---


# EncryptionProvider.EndSession Method (Office)

Ends the current encryption session.


## Syntax

 _expression_. **EndSession**( **_SessionHandle_** )

 _expression_ An expression that returns a **EncryptionProvider** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SessionHandle_|Required|**Long**|The ID of the current session.|

## Remarks

During a save operation, the  **CloneSession** method is called by your COM add-in to create a second, working copy of the **EncryptionProvider** object's encryption session for the file that is about to be saved. Next the **Save** method is called to get whatever custom information you would like to persist about your encryption settings. This information is available when this document is reopened later. Then the **EncryptStream** method is called which gives the provider the entire contents of the document. And finally, to complete the process, the **EndSession** method for the cloned session handle.


## See also


#### Concepts


[EncryptionProvider Object](encryptionprovider-object-office.md)
#### Other resources


[EncryptionProvider Object Members](encryptionprovider-members-office.md)

