---
title: Shape.Text Property (Visio)
keywords: vis_sdr.chm11251175
f1_keywords:
- vis_sdr.chm11251175
ms.prod: visio
api_name:
- Visio.Shape.Text
ms.assetid: 5c002c5d-f5ce-7f89-d799-4fc6ccb1a1f7
ms.date: 06/08/2017
---


# Shape.Text Property (Visio)

Returns all of the shape's text. Read/write.


## Syntax

 _expression_ . **Text**

 _expression_ A variable that represents a **Shape** object.


### Return Value

String


## Remarks

In the text returned by the  **Text** property of a **Shape** object, fields are represented by an escape character (30 (&;H1E)) For example, if a **Shape** object's text contains a field that displays the file name of a drawing, the **Shape** object's **Text** property returns an escape character where that field is inserted into the text. If you want the text to contain the expanded field, get the shape's **Characters** property, and then get the **Text** property of the resulting **Characters** object.

If the shape is a group, the text returned is dependent on the value of the IsTextEditTarget cell.




- If IsTextEditTarget is TRUE, the  **Text** property of the **Shape** object returns the text of the group.
    
- If IsTextEditTarget is FALSE, the  **Text** property of the **Shape** object returns the text of the shape in the group at the top of the stacking order.
    


Objects from other applications and guides do not have a  **Text** property.

If your Visual Studio solution includes the  **Microsoft.Office.Interop.Visio** reference, this property maps to the following types:


-  **Microsoft.Office.Interop.Visio.IVShape.Text**
    

## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to get the  **Text** property of a shape.


```vb
 
Public Sub ShapeText_Example()  
 
    Dim vsoRectangle As Visio.Shape  
    Dim vsoOval As Visio.Shape  
    Dim vsoShapeFromCell As Visio.Shape  
    Dim vsoShapeFromCharacters As Visio.Shape  
    Dim vsoCell As Visio.Cell  
    Dim vsoCharacters As Visio.Characters  
 
    'Create 2 different shapes and add different text to each shape. 
    Set vsoRectangle = ActivePage.DrawRectangle(2, 3, 5, 4)  
    Set vsoOval = ActivePage.DrawOval(2, 5, 5, 7)  
    vsoRectangle.Text = "Rectangle Shape"  
    vsoOval.Text = "Oval Shape"  
 
    'Get a Cell object from the first shape. 
    Set vsoCell = vsoRectangle.Cells("Width")  
 
    'Get a Characters object from the second shape. 
    Set vsoCharacters = vsoOval.Characters  
 
    'Use the Shape property to get the Shape object. 
    Set vsoShapeFromCell = vsoCell.Shape  
    Set vsoShapeFromCharacters = vsoCharacters.Shape  
 
    'Use each shape's text to verify the proper Shape 
    'object was returned.  
    Debug.Print vsoShapeFromCell.Text  
    Debug.Print vsoShapeFromCharacters.Text  
 
End Sub
```


