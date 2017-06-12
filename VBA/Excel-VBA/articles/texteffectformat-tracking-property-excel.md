---
title: TextEffectFormat.Tracking Property (Excel)
keywords: vbaxl10.chm118013
f1_keywords:
- vbaxl10.chm118013
ms.prod: excel
api_name:
- Excel.TextEffectFormat.Tracking
ms.assetid: b5190203-66c4-238b-e5b4-b61a9c70d99c
ms.date: 06/08/2017
---


# TextEffectFormat.Tracking Property (Excel)

Returns or sets the ratio of the horizontal space allotted to each character in the specified WordArt to the width of the character. Can be a value from 0 (zero) through 5. (Large values for this property specify ample space between characters; values less than 1 can produce character overlap.) Read/write  **Single** .


## Syntax

 _expression_ . **Tracking**

 _expression_ A variable that represents a **TextEffectFormat** object.


## Remarks

The following table gives the values of the  **Tracking** property that correspond to the settings available in the user interface.



|**User interface setting**|**Equivalent Tracking property value**|
|:-----|:-----|
|Very Tight|0.8|
|Tight|0.9|
|Normal|1.0|
|Loose|1.2|
|Very Loose|1.5|

## Example

This example adds WordArt that contains the text "Test" to  `myDocument` and specifies that the characters be very tightly spaced.


```vb
Set myDocument = Worksheets(1) 
Set newWordArt = myDocument.Shapes.AddTextEffect( _ 
 PresetTextEffect:=msoTextEffect1, Text:="Test", _ 
 FontName:="Arial Black", FontSize:=36, _ 
 FontBold:=False, FontItalic:=False, Left:=100, _ 
 Top:=100) 
newWordArt.TextEffect.Tracking =0.8
```


## See also


#### Concepts


[TextEffectFormat Object](texteffectformat-object-excel.md)

