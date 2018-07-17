---
title: CanvasShapes.AddTextbox Method (Word)
keywords: vbawd10.chm7536659
f1_keywords:
- vbawd10.chm7536659
ms.prod: word
api_name:
- Word.CanvasShapes.AddTextbox
ms.assetid: b8688f8f-db7e-7cf6-12ea-d0712b4ce26b
ms.date: 06/08/2017
---


# CanvasShapes.AddTextbox Method (Word)

Adds a text box to a drawing canvas. Returns a  **Shape** object that represents the text box.


## Syntax

 _expression_ . **AddTextbox**( **_Orientation_** , **_Left_** , **_Top_** , **_Width_** , **_Height_** )

 _expression_ Required. A variable that represents a **[CanvasShapes](canvasshapes-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Orientation_|Required| **MsoTextOrientation**|The orientation of the text. Some of the  **MsoTextOrientation** constants may not be available to you, depending on the language support (U.S. English, for example) that you?ve selected or installed.|
| _Left_|Required| **Single**|The position, measured in points, of the left edge of the text box.|
| _Top_|Required| **Single**|The position, measured in points, of the top edge of the text box.|
| _Width_|Required| **Single**|The width, measured in points, of the text box.|
| _Height_|Required| **Single**|The height, measured in points, of the text box.|

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


[CanvasShapes Collection](canvasshapes-object-word.md)

