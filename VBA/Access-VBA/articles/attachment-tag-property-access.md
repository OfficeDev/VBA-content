---
title: Attachment.Tag Property (Access)
keywords: vbaac10.chm13973
f1_keywords:
- vbaac10.chm13973
ms.prod: access
api_name:
- Access.Attachment.Tag
ms.assetid: 6d3e9f35-1986-e6b4-5f35-2652123c007c
ms.date: 06/08/2017
---


# Attachment.Tag Property (Access)

Stores extra information about a control needed by a Microsoft Access application. Read/write  **String**.


## Syntax

 _expression_. **Tag**

 _expression_ A variable that represents an **Attachment** object.


## Remarks

You can enter a string expression up to 2048 characters long. The default setting is a zero-length string (" ").

Unlike other properties, the  **Tag** property setting doesn't affect any of an object's attributes.

You can use this property to assign an identification string to an object without affecting any of its other property settings or causing other side effects. The  **Tag** property is useful when you need to check the identity of a form, report, section, or control that is passed as a variable to a procedure.


## See also


#### Concepts


[Attachment Object](attachment-object-access.md)

