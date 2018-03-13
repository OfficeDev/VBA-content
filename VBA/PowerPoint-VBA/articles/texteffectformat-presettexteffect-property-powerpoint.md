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
|<strong>msoTextEffect1</strong>|
|
<strong>msoTextEffect2</strong>|
|
<strong>msoTextEffect3</strong>|
|
<strong>msoTextEffect4</strong>|
|
<strong>msoTextEffect5</strong>|
|
<strong>msoTextEffect6</strong>|
|
<strong>msoTextEffect7</strong>|
|
<strong>msoTextEffect8</strong>|
|
<strong>msoTextEffect9</strong>|
|
<strong>msoTextEffect10</strong>|
|
<strong>msoTextEffect11</strong>|
|
<strong>msoTextEffect12</strong>|
|
<strong>msoTextEffect13</strong>|
|
<strong>msoTextEffect14</strong>|
|
<strong>msoTextEffect15</strong>|
|
<strong>msoTextEffect16</strong>|
|
<strong>msoTextEffect17</strong>|
|
<strong>msoTextEffect18</strong>|
|
<strong>msoTextEffect19</strong>|
|
<strong>msoTextEffect20</strong>|
|
<strong>msoTextEffect21</strong>|
|
<strong>msoTextEffect22</strong>|
|
<strong>msoTextEffect23</strong>|
|
<strong>msoTextEffect24</strong>|
|
<strong>msoTextEffect25</strong>|
|
<strong>msoTextEffect26</strong>|
|
<strong>msoTextEffect27</strong>|
|
<strong>msoTextEffect28</strong>|
|
<strong>msoTextEffect29</strong>|
|
<strong>msoTextEffect30</strong>|
|
<strong>msoTextEffectMixed</strong>|

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

