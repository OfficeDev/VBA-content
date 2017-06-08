---
title: Application.ActiveEncryptionSession Property (PowerPoint)
keywords: vbapp10.chm502059
f1_keywords:
- vbapp10.chm502059
ms.prod: powerpoint
api_name:
- PowerPoint.Application.ActiveEncryptionSession
ms.assetid: 73a174d5-a088-97d0-5f71-931456493224
ms.date: 06/08/2017
---


# Application.ActiveEncryptionSession Property (PowerPoint)

Represents the encryption session associated with the active presentation. Read-only.


## Syntax

 _expression_. **ActiveEncryptionSession**

 _expression_ An expression that returns a **Application** object.


### Return Value

Long


## Remarks

The encryption provider mechanism manages each file on a separate process, so each file is associated with a separate encryption session.


 **Note**  This property applies only when a presentation implements custom encryption.


## See also


#### Concepts


[Application Object](application-object-powerpoint.md)

