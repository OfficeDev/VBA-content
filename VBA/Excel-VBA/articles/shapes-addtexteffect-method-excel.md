---
title: Shapes.AddTextEffect Method (Excel)
keywords: vbaxl10.chm638085
f1_keywords:
- vbaxl10.chm638085
ms.prod: excel
api_name:
- Excel.Shapes.AddTextEffect
ms.assetid: ace2bd71-455d-d187-7fb0-77eed879ff95
ms.date: 06/08/2017
---


# Shapes.AddTextEffect Method (Excel)

Creates a WordArt object. Returns a  **[Shape](shape-object-excel.md)** object that represents the new WordArt object.


## Syntax

 _expression_ . **AddTextEffect**( **_PresetTextEffect_** , **_Text_** , **_FontName_** , **_FontSize_** , **_FontBold_** , **_FontItalic_** , **_Left_** , **_Top_** )

 _expression_ A variable that represents a **Shapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _PresetTextEffect_|Required| **[MsoPresetTextEffect](http://msdn.microsoft.com/library/56a7008d-ce2c-f127-56de-851cb8fef44f%28Office.15%29.aspx)**|The preset text effect.|
| _Text_|Required| **String**|The text in the WordArt.|
| _FontName_|Required| **String**|The name of the font used in the WordArt.|
| _FontSize_|Required| **Single**|The size (in points) of the font used in the WordArt.|
| _FontBold_|Required| **[MsoTriState](http://msdn.microsoft.com/library/2036cfc9-be7d-e05c-bec7-af05e3c3c515%28Office.15%29.aspx)**|The font used in the WordArt to bold.|
| _FontItalic_|Required| **[MsoTriState](http://msdn.microsoft.com/library/2036cfc9-be7d-e05c-bec7-af05e3c3c515%28Office.15%29.aspx)**|The font used in the WordArt to italic.|
| _Left_|Required| **Single**|The position (in points) of the upper-left corner of the WordArt's bounding box relative to the upper-left corner of the document.|
| _Top_|Required| **Single**|The position (in points) of the upper-left corner of the WordArt's bounding box relative to the top of the document.|

### Return Value

Shape


## Remarks

When you add WordArt to a document, the height and width of the WordArt are automatically set based on the size and amount of text you specify.


## Example

This example adds WordArt that contains the text "Test" to  `myDocument`.


```vb
Set myDocument = Worksheets(1) 
Set newWordArt = myDocument.Shapes.AddTextEffect( _ 
    PresetTextEffect:=msoTextEffect1, Text:="Test", _ 
    FontName:="Arial Black", FontSize:=36, _ 
    FontBold:=msoFalse, FontItalic:=msoFalse, Left:=10, _ 
    Top:=10)
```


## See also


#### Concepts


[Shapes Object](shapes-object-excel.md)

