---
title: TextEffectFormat.PresetTextEffect Property (PowerPoint)
keywords: vbapp10.chm556011
f1_keywords:
- vbapp10.chm556011
ms.prod: powerpoint
api_name:
- PowerPoint.TextEffectFormat.PresetTextEffect
ms.assetid: 629668e0-15c4-5867-acf9-6fc6ef8863ef
ms.date: 06/08/2017
---


# TextEffectFormat.PresetTextEffect Property (PowerPoint)

Returns or sets the style of the specified WordArt. Read/write.


## Syntax

 _expression_. **PresetTextEffect**

 _expression_ A variable that represents a **TextEffectFormat** object.


### Return Value

MsoPresetTextEffect


## Remarks

Setting the  **PresetTextEffect** property automatically sets many other formatting properties of the specified shape.

The value of the  **PresetTextEffect** property can be one of these **MsoPresetTextEffect** constants.


||
|:-----|
|**msoTextEffect1**|
|**msoTextEffect2**|
|**msoTextEffect3**|
|**msoTextEffect4**|
|**msoTextEffect5**|
|**msoTextEffect6**|
|**msoTextEffect7**|
|**msoTextEffect8**|
|**msoTextEffect9**|
|**msoTextEffect10**|
|**msoTextEffect11**|
|**msoTextEffect12**|
|**msoTextEffect13**|
|**msoTextEffect14**|
|**msoTextEffect15**|
|**msoTextEffect16**|
|**msoTextEffect17**|
|**msoTextEffect18**|
|**msoTextEffect19**|
|**msoTextEffect20**|
|**msoTextEffect21**|
|**msoTextEffect22**|
|**msoTextEffect23**|
|**msoTextEffect24**|
|**msoTextEffect25**|
|**msoTextEffect26**|
|**msoTextEffect27**|
|**msoTextEffect28**|
|**msoTextEffect29**|
|**msoTextEffect30**|
|**msoTextEffectMixed**|

## Example

This example sets the style for all WordArt on  `myDocument` to the first style listed in the **WordArt Quick Styles** tab.


```vb
Set myDocument = ActivePresentation.Slides(1)

For Each s In myDocument.Shapes

    If s.Type = msoTextEffect Then

        s.TextEffect.PresetTextEffect = msoTextEffect1

    End If

Next
```


## See also


#### Concepts


[TextEffectFormat Object](texteffectformat-object-powerpoint.md)

