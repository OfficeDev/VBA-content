---
title: Shapes.AddLabel Method (Word)
keywords: vbawd10.chm161415181
f1_keywords:
- vbawd10.chm161415181
ms.prod: word
api_name:
- Word.Shapes.AddLabel
ms.assetid: c5a5ea1f-5ca3-2c8d-40e2-bad94a220609
ms.date: 06/08/2017
---


# Shapes.AddLabel Method (Word)

Adds a text label to a drawing canvas.


## Syntax

 _expression_ . **AddLabel**( **_Orientation_** , **_Left_** , **_Top_** , **_Width_** , **_Height_** )

 _expression_ Required. A variable that represents a **[Shapes](shapes-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Orientation_|Required| **MsoTextOrientation**|The orientation of the text.|
| _Left_|Required| **Single**|The position, measured in points, of the left edge of the label relative to the left edge of the drawing canvas.|
| _Top_|Required| **Single**|The position, measured in points, of the top edge of the label relative to the top edge of the drawing canvas.|
| _Width_|Required| **Single**|The width of the label, in points.|
| _Height_|Required| **Single**|The height of the label, in points.|

### Return Value

 **Shapes**


## Example

This example adds a blue text label with the text "Hello World" to a new drawing canvas in the active document.


```vb
Sub NewCanvasTextLabel() 
 Dim shpCanvas As Shape 
 Dim shpLabel As Shape 
 
 'Add a drawing canvas to the active document 
 Set shpCanvas = ActiveDocument.Shapes.AddCanvas _ 
 (Left:=100, Top:=75, Width:=150, Height:=200) 
 
 'Add a label to the drawing canvas 
 Set
```


```vb
shpLabel = shpCanvas.CanvasItems.AddLabel _ 
 (Orientation:=msoTextOrientationHorizontal, _ 
 Left:=15, Top:=15, Width:=100, Height:=100) 
 
 'Fill the label textbox with a color, 
 'add text to the label and format it 
 With
```


```vb
shpLabel 
 With .Fill 
 .BackColor.RGB = RGB(Red:=0, Green:=0, Blue:=192) 
 'Make the fill visible 
 .Visible = msoTrue 
 End With 
 With .TextFrame.TextRange 
 .Text = "Hello World." 
 .Bold = True 
 .Font.Name = "Tahoma" 
 End With 
 End With 
End Sub
```


## See also


#### Concepts


[Shapes Collection Object](shapes-object-word.md)

