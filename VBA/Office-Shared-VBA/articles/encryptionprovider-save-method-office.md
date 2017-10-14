---
title: EncryptionProvider.Save Method (Office)
keywords: vbaof11.chm327006
f1_keywords:
- vbaof11.chm327006
ms.prod: office
api_name:
- Office.EncryptionProvider.Save
ms.assetid: 7dfb6cea-f97b-51c3-e6bb-a773eec3fa73
ms.date: 06/08/2017
---


# EncryptionProvider.Save Method (Office)

Saves an encrypted document.


## Syntax

 _expression_. **Save**( **_SessionHandle_**, **_EncryptionData_** )

 _expression_ An expression that returns a **EncryptionProvider** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SessionHandle_|Required|**Long**|The ID of the current session.|
| _EncryptionData_|Required|**IUnknown**|Contains the encryption information.|

### Return Value

Long


## Remarks

When you save a file to the Office Open XML File Format (which is the only format that supports custom file encryption), then the provider is called by your COM add-in to encrypt the document. If you attempt to save to a format that does not support custom file encryption and you have the appropriate rights to do so, then Microsoft Office will save the document without encryption. This allows documents to be exported to formats that do not support encryption or rights management.


## See also


#### Concepts


[EncryptionProvider Object](encryptionprovider-object-office.md)
#### Other resources


[EncryptionProvider Object Members](encryptionprovider-members-office.md)

