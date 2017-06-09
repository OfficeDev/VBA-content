---
title: TextRange2.Font Property (Office)
ms.prod: office
api_name:
- Office.TextRange2.Font
ms.assetid: 005fa6bf-2dd5-32ec-18e8-30ff6260e55d
ms.date: 06/08/2017
---


# TextRange2.Font Property (Office)

Returns a  **Font** object that represents character formatting for the **TextRange2** object. Read-only.


## Syntax

 _expression_. **Font**

 _expression_ An expression that returns a **TextRange2** object.


### Return Value

Font


## Example

This example sets the formatting for the text in shape one on slide one in the active PowerPoint presentation.


```
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


[TextRange2 Object](textrange2-object-office.md)
#### Other resources


[TextRange2 Object Members](textrange2-members-office.md)

