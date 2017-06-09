---
title: Application.VBE Property (PowerPoint)
keywords: vbapp10.chm502020
f1_keywords:
- vbapp10.chm502020
ms.prod: powerpoint
api_name:
- PowerPoint.Application.VBE
ms.assetid: 33a3d113-31f6-3705-cdb9-d5e07fa82820
ms.date: 06/08/2017
---


# Application.VBE Property (PowerPoint)

Returns a  **VBE** object that represents the Visual Basic Editor. Read-only.


## Syntax

 _expression_. **VBE**

 _expression_ A variable that represents a **Application** object.


### Return Value

VBE


## Example

This example sets the name of the active project in the Visual Basic Editor.


```vb
Application.VBE.ActiveVBProject.Name = "TestProject"
```


## See also


#### Concepts


[Application Object](application-object-powerpoint.md)

