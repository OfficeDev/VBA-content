---
title: TextRange2.Font Property (PowerPoint)
ms.assetid: 3d47ff57-6622-4eaa-b8ff-b395e9757096
ms.date: 06/08/2017
ms.prod: powerpoint
---


# TextRange2.Font Property (PowerPoint)

Returns a  **Font** object that represents character formatting for the **TextRange2** object. Read-only.


## Syntax

 _expression_. **Font**

 _expression_ An expression that returns a **TextRange2** object.


### Return Value

Font


## Example

This example sets the formatting for the text in shape one on slide one in the active PowerPoint presentation.


```vb
With ActivePresentation.Slides(1).Shapes(1) 
 With .TextFrame.TextRange2.Font 
 .Size = 48 
 .Name = "Palatino" 
 .Bold = True 
 .Color.RGB = RGB(255, 127, 255) 
 End With 
End With
```


## See also


#### Concepts


[TextRange2 Object (PowerPoint)](textrange2-object-powerpoint.md)


