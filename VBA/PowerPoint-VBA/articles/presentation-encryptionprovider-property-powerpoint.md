---
title: Presentation.EncryptionProvider Property (PowerPoint)
keywords: vbapp10.chm583109
f1_keywords:
- vbapp10.chm583109
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.EncryptionProvider
ms.assetid: 9b316f21-eeaf-4704-636f-ea68c7a36cfd
ms.date: 06/08/2017
---


# Presentation.EncryptionProvider Property (PowerPoint)

Returns a  **String** that specifies the name of the algorithm encryption provider that PowerPoint uses when encrypting documents. Read/write.


## Syntax

 _expression_. **EncryptionProvider**

 _expression_ An expression that returns a **Presentation** object.


### Return Value

String


## Remarks

You can implement an encryption provider by creating a custom COM add-in. Within your presentation, you can store information that the add-in can encrypt and decrypt, and to which it can apply rights. The add-in can also display permission, setup, or authentication user interfaces.


## See also


#### Concepts


[Presentation Object](presentation-object-powerpoint.md)

