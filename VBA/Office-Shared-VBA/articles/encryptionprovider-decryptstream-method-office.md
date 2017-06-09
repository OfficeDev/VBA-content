---
title: EncryptionProvider.DecryptStream Method (Office)
keywords: vbaof11.chm327008
f1_keywords:
- vbaof11.chm327008
ms.prod: office
api_name:
- Office.EncryptionProvider.DecryptStream
ms.assetid: da893485-b450-48aa-624d-e8bc2794c65a
ms.date: 06/08/2017
---


# EncryptionProvider.DecryptStream Method (Office)

Decrypts and returns a stream of encrypted data for a document.


## Syntax

 _expression_. **DecryptStream**( **_SessionHandle_**, **_StreamName_**, **_EncryptedStream_**, **_UnencryptedStream_** )

 _expression_ An expression that returns a **EncryptionProvider** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SessionHandle_|Required|**Long**|The ID of the current session.|
| _StreamName_|Required|**String**|The ID of the stream of data.|
| _EncryptedStream_|Required|**IUnknown**|The encrypted data stream.|
| _UnencryptedStream_|Required|**IUnknown**|The data stream before dencryption.|

## Remarks

This method is called by your COM add-in when the document is opened, and after your add-in has verified that the user opening the document is authenticated. This method is the inverse of EncryptStream method and converts encrypted data back into pure (un-encrypted) data. 


## See also


#### Concepts


[EncryptionProvider Object](encryptionprovider-object-office.md)
#### Other resources


[EncryptionProvider Object Members](encryptionprovider-members-office.md)

