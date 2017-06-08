---
title: EncryptionProvider.NewSession Method (Office)
keywords: vbaof11.chm327002
f1_keywords:
- vbaof11.chm327002
ms.prod: office
api_name:
- Office.EncryptionProvider.NewSession
ms.assetid: b90f842a-6eb3-3e95-7175-c3ca9c3ce138
ms.date: 06/08/2017
---


# EncryptionProvider.NewSession Method (Office)

Used by the  **EncryptionProvider** object to create a new encryption session. This session is used by the provider to cache document-specific information about the encryption, users, and rights while the document is in memory.


## Syntax

 _expression_. **NewSession**( **_ParentWindow_** )

 _expression_ An expression that returns a **EncryptionProvider** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ParentWindow_|Required|**IUnknown**|Specifies the window that is called to display the encryption settings.|

### Return Value

Long


## Remarks

This method is called by your COM add-in.


## See also


#### Concepts


[EncryptionProvider Object](encryptionprovider-object-office.md)
#### Other resources


[EncryptionProvider Object Members](encryptionprovider-members-office.md)

