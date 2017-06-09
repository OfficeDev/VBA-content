---
title: TextEffectFormat.Alignment Property (Excel)
keywords: vbaxl10.chm118002
f1_keywords:
- vbaxl10.chm118002
ms.prod: excel
api_name:
- Excel.TextEffectFormat.Alignment
ms.assetid: 0a86ac22-9496-d801-0cfb-a9fca5c30fec
ms.date: 06/08/2017
---


# TextEffectFormat.Alignment Property (Excel)

Returns or sets an  **[MsoTextEffectAlignment](http://msdn.microsoft.com/library/5a165109-c820-bbc2-235b-a24807abd0d0%28Office.15%29.aspx)** value that represents the alignment for WordArt.


## Syntax

 _expression_ . **Alignment**

 _expression_ A variable that represents a **TextEffectFormat** object.


## Example

This example adds a WordArt object to worksheet one and then right aligns the WordArt.


```vb
Set mySh = Worksheets(1).Shapes 
Set myTE = mySh.AddTextEffect(PresetTextEffect:=msoTextEffect1, _ 
    Text:="Test Text", FontName:="Palatino", FontSize:=54, _ 
    FontBold:=True, FontItalic:=False, Left:=100, Top:=50) 
myTE.TextEffect.Alignment = msoTextEffectAlignmentRight
```


## See also


#### Concepts


[TextEffectFormat Object](texteffectformat-object-excel.md)

