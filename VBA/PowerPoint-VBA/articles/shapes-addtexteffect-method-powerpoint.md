---
title: Shapes.AddTextEffect Method (PowerPoint)
keywords: vbapp10.chm543013
f1_keywords:
- vbapp10.chm543013
ms.prod: powerpoint
api_name:
- PowerPoint.Shapes.AddTextEffect
ms.assetid: 4428ac57-c704-475a-1640-78a556e9ac3d
ms.date: 06/08/2017
---


# Shapes.AddTextEffect Method (PowerPoint)

Creates a WordArt object. Returns a  **[Shape](shape-object-powerpoint.md)** object that represents the new WordArt object.


## Syntax

 _expression_. **AddTextEffect**( **_PresetTextEffect_**, **_Text_**, **_FontName_**, **_FontSize_**, **_FontBold_**, **_FontItalic_**, **_Left_**, **_Top_** )

 _expression_ A variable that represents a **Shapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _PresetTextEffect_|Required|**[MsoPresetTextEffect](http://msdn.microsoft.com/library/56a7008d-ce2c-f127-56de-851cb8fef44f%28Office.15%29.aspx)**|The preset text effect.|
| _Text_|Required|**String**|The text in the WordArt.|
| _FontName_|Required|**String**|The name of the font used in the WordArt.|
| _FontSize_|Required|**Single**|The size (in points) of the font used in the WordArt.|
| _FontBold_|Required|**[MsoTriState](http://msdn.microsoft.com/library/2036cfc9-be7d-e05c-bec7-af05e3c3c515%28Office.15%29.aspx)**|Determines whether the font used in the WordArt is set to bold.|
| _FontItalic_|Required|**[MsoTriState](http://msdn.microsoft.com/library/2036cfc9-be7d-e05c-bec7-af05e3c3c515%28Office.15%29.aspx)**|Determines whether the font used in the WordArt is set to italic.|
| _Left_|Required|**Single**|The position, measured in points, of the left edge of the WordArt's bounding box relative to the left edge of the slide.|
| _Top_|Required|**Single**|The position, measured in points, of the top edge of the WordArt's bounding box relative to the top edge of the slide.|

### Return Value

Shape


## Remarks

When you add WordArt to a document, the height and width of the WordArt are automatically set based on the size and amount of text you specify.


## Example

This example adds WordArt that contains the text "Test" to myDocument.


```vb
Set myDocument = ActivePresentation.Slides(1) 
Set newWordArt = myDocument.Shapes _ 
    .AddTextEffect(PresetTextEffect:=msoTextEffect1, _ 
    Text:="Test", FontName:="Arial Black", FontSize:=36, _ 
    FontBold:=msoFalse, FontItalic:=msoFalse, Left:=10, Top:=10)
```


## See also


#### Concepts


[Shapes Object](shapes-object-powerpoint.md)

