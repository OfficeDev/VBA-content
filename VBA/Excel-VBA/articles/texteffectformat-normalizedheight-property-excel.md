---
title: TextEffectFormat.NormalizedHeight Property (Excel)
keywords: vbaxl10.chm118008
f1_keywords:
- vbaxl10.chm118008
ms.prod: excel
api_name:
- Excel.TextEffectFormat.NormalizedHeight
ms.assetid: 25c9c1ed-971d-3a9f-bb3c-5059f2dd80db
ms.date: 06/08/2017
---


# TextEffectFormat.NormalizedHeight Property (Excel)

 **True** if all characters (both uppercase and lowercase) in the specified WordArt are the same height. Read/write **MsoTriState** .


## Syntax

 _expression_ . **NormalizedHeight**

 _expression_ A variable that represents a **TextEffectFormat** object.


## Remarks





| **MsoTriState** can be one of these **MsoTriState** constants.|
| **msoCTrue**|
| **msoFalse**|
| **msoTriStateMixed**|
| **msoTriStateToggle**|
| **msoTrue** All characters (both uppercase and lowercase) in the specified WordArt are the same height.|

## Example

This example adds WordArt that contains the text "Test Effect" to  `myDocument` and gives the new WordArt the name "texteff1." The code then makes all characters in the shape named "texteff1" the same height.


```vb
Set myDocument = Worksheets(1) 
myDocument.Shapes.AddTextEffect( _ 
 PresetTextEffect:=msoTextEffect1, _ 
 Text:="Test Effect", FontName:="Courier New", _ 
 FontSize:=44, FontBold:=True, _ 
 FontItalic:=False, Left:=10, Top:=10).Name = "texteff1" 
myDocument.Shapes("texteff1").TextEffect.NormalizedHeight = msoTrue
```


## See also


#### Concepts


[TextEffectFormat Object](texteffectformat-object-excel.md)

