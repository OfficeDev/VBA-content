---
title: Application.VBE Property (Word)
keywords: vbawd10.chm158335037
f1_keywords:
- vbawd10.chm158335037
ms.prod: word
api_name:
- Word.Application.VBE
ms.assetid: 641109fd-7ece-9efd-65ba-56e223d8249c
ms.date: 06/08/2017
---


# Application.VBE Property (Word)

Returns a VBE object that represents the Visual Basic Editor.


## Syntax

 _expression_ . **VBE**

 _expression_ An expression that returns an **[Application](application-object-word.md)** object.


## Example

This example displays the number of references available for the active project.


```vb
MsgBox "References = " &; VBE.ActiveVBProject.References.Count
```


## See also


#### Concepts


[Application Object](application-object-word.md)

