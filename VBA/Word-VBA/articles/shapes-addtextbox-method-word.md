---
title: Shapes.AddTextbox Method (Word)
keywords: vbawd10.chm161415187
f1_keywords:
- vbawd10.chm161415187
ms.prod: word
api_name:
- Word.Shapes.AddTextbox
ms.assetid: 7b5c766e-40b3-a390-561f-cd1a53eb93a7
ms.date: 06/08/2017
---


# Shapes.AddTextbox Method (Word)

Adds a text box to a drawing canvas.


## Syntax

 _expression_ . **AddTextbox**( **_Orientation_** , **_Left_** , **_Top_** , **_Width_** , **_Height_** )

 _expression_ Required. A variable that represents a **[Shapes](shapes-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Orientation_|Required| **MsoTextOrientation**|The orientation of the text. Some of these constants may not be available to you, depending on the language support (U.S. English, for example) that you?ve selected or installed.|
| _Left_|Required| **Single**|The position, measured in points, of the left edge of the text box.|
| _Top_|Required| **Single**|The position, measured in points, of the top edge of the text box.|
| _Width_|Required| **Single**|The width, measured in points, of the text box.|
| _Height_|Required| **Single**|The height, measured in points, of the text box.|

### Return Value

 **Shape**


## Example

This example add a textbox to a canvas in a new document.


```vb
Sub NewCanvasTextbox() 
 Dim docNew As Document 
 Dim shpCanvas As Shape 
 
 'Create a new document and add a drawing canvas 
 Set docNew = Documents.Add 
 Set shpCanvas = docNew.Shapes.AddCanvas _ 
 (Left:=100, Top:=75, Width:=150, Height:=200) 
 
 'Add a text box to the drawing canvas 
 shpCanvas.CanvasItems.AddTextbox _ 
 Orientation:=msoTextOrientationHorizontal, _ 
 Left:=1, Top:=1, Width:=100, Height:=100 
End Sub
```


## See also


#### Concepts


[Shapes Collection Object](shapes-object-word.md)

