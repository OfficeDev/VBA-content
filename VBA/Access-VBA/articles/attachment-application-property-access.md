---
title: Attachment.Application Property (Access)
keywords: vbaac10.chm13903
f1_keywords:
- vbaac10.chm13903
ms.prod: access
api_name:
- Access.Attachment.Application
ms.assetid: db88250d-da59-300c-6f0c-3768c1bb8a7f
ms.date: 06/08/2017
---


# Attachment.Application Property (Access)

You can use the  **Application** property to access the active Microsoft Access **[Application](application-object-access.md)** object and its related properties. Read-only **Application** object.


## Syntax

 _expression_.

 _expression_ A variable that represents an **Attachment** object.


## Remarks

The  **Application** property is set by Microsoft Access and is read-only in all views.

Each Microsoft Access object has an  **Application** property that returns the current **Application** object. You can use this property to access any of the object's properties. For example, you could refer to the menu bar for the **Application** object from the current form by using the following syntax:




```vb
Me.Application.MenuBar 

```


## See also


#### Concepts


[Attachment Object](attachment-object-access.md)

