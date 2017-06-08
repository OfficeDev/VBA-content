---
title: Presentation.Container Property (PowerPoint)
keywords: vbapp10.chm583041
f1_keywords:
- vbapp10.chm583041
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.Container
ms.assetid: cc0108b7-ce95-3a1b-a400-c49700a2362c
ms.date: 06/08/2017
---


# Presentation.Container Property (PowerPoint)

Returns the object that contains the specified embedded presentation. Read-only.


## Syntax

 _expression_. **Container**

 _expression_ A variable that represents a **Presentation** object.


### Return Value

Object


## Remarks

If the container doesn't support OLE Automation, or if the specified presentation isn't embedded in a Microsoft Binder file, this property fails.


## Example

This example hides the second section of the Microsoft Binder file that contains the embedded active presentation. The  **Container** property of the presentation returns a **Section** object, and the **Parent** property of the **Section** object returns a **Binder** object.


```vb
Application.ActivePresentation.Container.Parent.Sections(2) _
    .Visible = False
```


## See also


#### Concepts


[Presentation Object](presentation-object-powerpoint.md)

