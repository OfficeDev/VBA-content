---
title: TextEffectFormat.PresetTextEffect Property (Excel)
keywords: vbaxl10.chm118010
f1_keywords:
- vbaxl10.chm118010
ms.prod: excel
api_name:
- Excel.TextEffectFormat.PresetTextEffect
ms.assetid: 13ff8b1a-d12e-47c1-6f82-0b3b9b5a7bf4
ms.date: 06/08/2017
---


# TextEffectFormat.PresetTextEffect Property (Excel)

Returns or sets the style of the specified WordArt. Read/write  **MsoPresetTextEffect** .


## Syntax

 _expression_ . **PresetTextEffect**

 _expression_ A variable that represents a **TextEffectFormat** object.


## Remarks

The values for this property correspond to the formats in the  **WordArt Gallery** dialog box (numbered from left to right, top to bottom).



| **MsoPresetTextEffect** can be one of these **MsoPresetTextEffect** constants.|
| **msoTextEffect1**|
| **msoTextEffect10**|
| **msoTextEffect11**|
| **msoTextEffect12**|
| **msoTextEffect13**|
| **msoTextEffect14**|
| **msoTextEffect15**|
| **msoTextEffect16**|
| **msoTextEffect17**|
| **msoTextEffect18**|
| **msoTextEffect19**|
| **msoTextEffect2**|
| **msoTextEffect20**|
| **msoTextEffect21**|
| **msoTextEffect22**|
| **msoTextEffect23**|
| **msoTextEffect24**|
| **msoTextEffect25**|
| **msoTextEffect26**|
| **msoTextEffect27**|
| **msoTextEffect28**|
| **msoTextEffect29**|
| **msoTextEffect3**|
| **msoTextEffect30**|
| **msoTextEffect4**|
| **msoTextEffect5**|
| **msoTextEffect6**|
| **msoTextEffect7**|
| **msoTextEffect8**|
| **msoTextEffect9**|
| **msoTextEffectMixed**|
Setting the  **PresetTextEffect** property automatically sets many other formatting properties of the specified shape.


## Example

This example sets the style for all WordArt on  `myDocument` to the first style listed in the **WordArt Gallery** dialog box.


```vb
Set myDocument = Worksheets(1) 
For Each s In myDocument.Shapes 
 If s.Type = msoTextEffect Then 
 s.TextEffect.PresetTextEffect = msoTextEffect1 
 End If 
Next
```


## See also


#### Concepts


[TextEffectFormat Object](texteffectformat-object-excel.md)

