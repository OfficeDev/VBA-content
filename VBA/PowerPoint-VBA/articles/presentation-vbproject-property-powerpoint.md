---
title: Presentation.VBProject Property (PowerPoint)
keywords: vbapp10.chm583022
f1_keywords:
- vbapp10.chm583022
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.VBProject
ms.assetid: 76713c8c-2263-7a5a-8133-726cc94bd73a
ms.date: 06/08/2017
---


# Presentation.VBProject Property (PowerPoint)

Returns a  **VBProject** object that represents the individual Visual Basic project for the presentation. Read-only.


## Syntax

 _expression_. **VBProject**

 _expression_ A variable that represents a **Presentation** object.


### Return Value

VBProject


## Example

This example changes the name of the Visual Basic project for the active presentation.


```vb
ActivePresentation.VBProject.Name = "TestProject"
```


## See also


#### Concepts


[Presentation Object](presentation-object-powerpoint.md)

