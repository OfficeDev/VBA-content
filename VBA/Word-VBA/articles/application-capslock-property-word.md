---
title: Application.CapsLock Property (Word)
keywords: vbawd10.chm158335024
f1_keywords:
- vbawd10.chm158335024
ms.prod: word
api_name:
- Word.Application.CapsLock
ms.assetid: 73cc2530-5178-d348-739e-c4605b8f207d
ms.date: 06/08/2017
---


# Application.CapsLock Property (Word)

 **True** if the CAPS LOCK key is turned on. Read-only **Boolean** .


## Syntax

 _expression_ . **CapsLock**

 _expression_ A variable that represents an **[Application](application-object-word.md)** object.


## Example

This example retrieves the current state of the CAPS LOCK key.


```vb
Dim blnCapsLock As Boolean 
 
blnCapsLock 
= Application.CapsLock
```


## See also


#### Concepts


[Application Object](application-object-word.md)

