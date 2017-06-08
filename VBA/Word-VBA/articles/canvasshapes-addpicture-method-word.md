---
title: CanvasShapes.AddPicture Method (Word)
keywords: vbawd10.chm7536655
f1_keywords:
- vbawd10.chm7536655
ms.prod: word
api_name:
- Word.CanvasShapes.AddPicture
ms.assetid: 445a19cf-1388-34d7-69d8-778d889d917f
ms.date: 06/08/2017
---


# CanvasShapes.AddPicture Method (Word)

Adds a picture to a drawing canvas. Returns a  **Shape** object that represents the picture and adds it to the **CanvasShapes** collection.


## Syntax

 _expression_ . **AddPicture**( **_FileName_** , **_LinkToFile_** , **_SaveWithDocument_** , **_Left_** , **_Top_** , **_Width_** , **_Height_** )

 _expression_ Required. A variable that represents a **[CanvasShapes](canvasshapes-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FileName_|Required| **String**|The path and file name of the picture.|
| _LinkToFile_|Optional| **Variant**| **True** to link the picture to the file from which it was created. **False** to make the picture an independent copy of the file. The default value is **False** .|
| _SaveWithDocument_|Optional| **Variant**| **True** to save the linked picture with the document. The default value is **False** .|
| _Left_|Optional| **Variant**|The position, measured in points, of the left edge of the new picture relative to the drawing canvas.|
| _Top_|Optional| **Variant**|The position, measured in points, of the top edge of the new picture relative to the drawing canvas.|
| _Width_|Optional| **Variant**|The width of the picture, in points.|
| _Height_|Optional| **Variant**|The height of the picture, in points.|

## Example

This example adds a picture to a newly created drawing canvas in the active document.


```vb
Sub NewCanvasPicture() 
 Dim shpCanvas As Shape 
 
 'Add a drawing canvas to the active document 
 Set shpCanvas = ActiveDocument.Shapes _ 
 .AddCanvas(Left:=100, Top:=75, _ 
 Width:=200, Height:=300) 
 
 'Add a graphic to the drawing canvas 
 shpCanvas.CanvasItems.AddPicture _ 
 FileName:="C:\Program Files\Microsoft Office\" &; _ 
 "Office\Bitmaps\Styles\stone.bmp", _ 
 LinkToFile:=False, SaveWithDocument:=True 
End Sub
```


## See also


#### Concepts


[CanvasShapes Collection](canvasshapes-object-word.md)

