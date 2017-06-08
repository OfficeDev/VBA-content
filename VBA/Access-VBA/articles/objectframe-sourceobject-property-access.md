---
title: ObjectFrame.SourceObject Property (Access)
keywords: vbaac10.chm11573
f1_keywords:
- vbaac10.chm11573
ms.prod: access
api_name:
- Access.ObjectFrame.SourceObject
ms.assetid: 985c8b01-84d8-2da6-6cad-5de08d835434
ms.date: 06/08/2017
---


# ObjectFrame.SourceObject Property (Access)

You can use this property for linked unbound object frames to determine the complete path and file name of the file that contains the data linked to the object frame. Read-only  **String**.


## Syntax

 _expression_. **SourceObject**

 _expression_ A variable that represents an **ObjectFrame** object.


## Remarks

For unbound object frames, the  **SourceObject** property is set automatically when you use the **SourceObject** command on the **Insert** menu to insert a linked OLE object.

For linked unbound object frames, the  **SourceObject** property can't be set in any view.


## See also


#### Concepts


[ObjectFrame Object](objectframe-object-access.md)

