---
title: Application.ActiveEncryptionSession Property (Word)
keywords: vbawd10.chm158335455
f1_keywords:
- vbawd10.chm158335455
ms.prod: word
api_name:
- Word.Application.ActiveEncryptionSession
ms.assetid: a9ea5257-4cda-e25d-8af5-e29c10b7faa2
ms.date: 06/08/2017
---


# Application.ActiveEncryptionSession Property (Word)

Returns a  **Long** that represents the encryption session associated with the active document. Read-only.


## Syntax

 _expression_ . **ActiveEncryptionSession**

 _expression_ An expression that returns an **Application** object.


## Remarks

The encryption provider mechanism manages each file on a separate process, so each file is associated with a separate encryption session.


 **Note**  This property applies only when a document implements custom encryption.


## See also


#### Concepts


[Application Object](application-object-word.md)

