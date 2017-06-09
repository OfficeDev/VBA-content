---
title: EncryptionProvider.EncryptStream Method (Office)
keywords: vbaof11.chm327007
f1_keywords:
- vbaof11.chm327007
ms.prod: office
api_name:
- Office.EncryptionProvider.EncryptStream
ms.assetid: 58a379f4-fb74-4a2c-b0ed-ce3e3151c292
ms.date: 06/08/2017
---


# EncryptionProvider.EncryptStream Method (Office)

Encrypts and returns a stream of data for a document.


## Syntax

 _expression_. **EncryptStream**( **_SessionHandle_**, **_StreamName_**, **_UnencryptedStream_**, **_EncryptedStream_** )

 _expression_ An expression that returns a **EncryptionProvider** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SessionHandle_|Required|**Long**|The ID of the current session.|
| _StreamName_|Required|**String**|The name of the encrypted stream of document data.|
| _UnencryptedStream_|Required|**IUnknown**|The data stream before encryption.|
| _EncryptedStream_|Required|**IUnknown**|The data stream information after it has been encrypted.|

## Remarks

This method is typically called by your COM add-in during a save operation.


## See also


#### Concepts


[EncryptionProvider Object](encryptionprovider-object-office.md)
#### Other resources


[EncryptionProvider Object Members](encryptionprovider-members-office.md)

