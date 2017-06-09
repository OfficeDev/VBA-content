---
title: EncryptionProvider.CloneSession Method (Office)
keywords: vbaof11.chm327004
f1_keywords:
- vbaof11.chm327004
ms.prod: office
api_name:
- Office.EncryptionProvider.CloneSession
ms.assetid: d7548ad1-caec-27d8-db55-c4e6f747111e
ms.date: 06/08/2017
---


# EncryptionProvider.CloneSession Method (Office)

Creates a second, working copy of the  **EncryptionProvider** object's encryption session for a file that is about to be saved.


## Syntax

 _expression_. **CloneSession**( **_SessionHandle_** )

 _expression_ An expression that returns a **EncryptionProvider** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SessionHandle_|Required|**Long**|The ID of the cloned session.|

### Return Value

Long


## Remarks

The output of this method is another session handle with the same encryption settings. When an autoSave operation occurs, you are provided with this session handle.


## See also


#### Concepts


[EncryptionProvider Object](encryptionprovider-object-office.md)
#### Other resources


[EncryptionProvider Object Members](encryptionprovider-members-office.md)

