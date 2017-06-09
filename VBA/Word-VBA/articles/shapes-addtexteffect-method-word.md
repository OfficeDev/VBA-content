---
title: Shapes.AddTextEffect Method (Word)
keywords: vbawd10.chm161415186
f1_keywords:
- vbawd10.chm161415186
ms.prod: word
api_name:
- Word.Shapes.AddTextEffect
ms.assetid: 1f1fca1b-f357-8526-75a4-b05a378736bc
ms.date: 06/08/2017
---


# Shapes.AddTextEffect Method (Word)

Adds a WordArt shape to a drawing canvas. Returns a  **Shape** object that represents the WordArt and adds it to the **CanvasShapes** collection.


## Syntax

 _expression_ . **AddTextEffect**( **_PresetTextEffect_** , **_Text_** , **_FontName_** , **_FontSize_** , **_FontBold_** , **_FontItalic_** , **_Left_** , **_Top_** )

 _expression_ Required. A variable that represents a **[Shapes](shapes-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _PresetTextEffect_|Required| **MsoPresetTextEffect**|A preset text effect. The values of the  **MsoPresetTextEffect** constants correspond to the formats listed in the **WordArt Gallery** dialog box (numbered from left to right and from top to bottom).|
| _Text_|Required| **String**|The text in the WordArt.|
| _FontName_|Required| **String**|The name of the font used in the WordArt.|
| _FontSize_|Required| **Single**|The size (in points) of the font used in the WordArt.|
| _FontBold_|Required| **MsoTriState**| **MsoTrue** to bold the WordArt font.|
| _FontItalic_|Required| **MsoTriState**| **MsoTrue** to italicize the WordArt font.|
| _Left_|Required| **Single**|The position, measured in points, of the left edge of the WordArt shape relative to the left edge of the drawing canvas.|
| _Top_|Required| **Single**|The position, measured in points, of the top edge of the WordArt shape relative to the top edge of the drawing canvas.|

## Remarks

When you add WordArt to a document, the height and width of the WordArt are automatically set based on the size and amount of text you specify.


## Example

This example adds a drawing canvas to a new document and inserts a WordArt shape inside the canvas that contains the text "Hello, World."


```vb
Sub NewCanvasTextEffect() 
 Dim docNew As Document 
 Dim shpCanvas As Shape 
 
 'Create a new document and add a drawing canvas 
 Set docNew = Documents.Add 
 Set shpCanvas = docNew.Shapes.AddCanvas( _ 
 Left:=100, Top:=100, Width:=150, _ 
 Height:=50) 
 
 'Add WordArt shape to the drawing canvas 
 shpCanvas.CanvasItems.AddTextEffect _ 
 PresetTextEffect:=msoTextEffect20, _ 
 Text:="Hello, World", FontName:="Tahoma", _ 
 FontSize:=15, FontBold:=msoTrue, _ 
 FontItalic:=msoFalse, _ 
 Left:=120, Top:=120 
End Sub
```


## See also


#### Concepts


[Shapes Collection Object](shapes-object-word.md)

