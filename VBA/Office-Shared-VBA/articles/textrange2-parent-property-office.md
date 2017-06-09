---
title: TextRange2.Parent Property (Office)
ms.prod: office
api_name:
- Office.TextRange2.Parent
ms.assetid: 692dc869-1525-ffa5-023d-83cea9cec19e
ms.date: 06/08/2017
---


# TextRange2.Parent Property (Office)

Gets the  **Parent** object for the **TextRange2** object. Read-only.


## Syntax

 _expression_. **Parent**

 _expression_ An expression that returns a **TextRange2** object.


### Return Value

Object


## Example

This example adds an oval containing text to slide one in the active presentation and rotates the oval and the text 45 degrees. The parent object for the text frame is the  **Shape** object that contains the text.


```
Set myShapes = ActivePresentation.Slides(1).Shapes 
With myShapes.AddShape(Type:=msoShapeOval, Left:=50, _ 
 Top:=50, Width:=300, Height:=150).TextFrame 
 .TextRange2.Text = "Test text" 
 .Parent.Rotation = 45 
End With
```


## See also


#### Concepts


[TextRange2 Object](textrange2-object-office.md)
#### Other resources


[TextRange2 Object Members](textrange2-members-office.md)

