---
title: FillFormat.OneColorGradient Method (Word)
keywords: vbawd10.chm164102155
f1_keywords:
- vbawd10.chm164102155
ms.prod: word
api_name:
- Word.FillFormat.OneColorGradient
ms.assetid: 993ae539-0051-cbf1-390b-4852aa97f5fb
ms.date: 06/08/2017
---


# FillFormat.OneColorGradient Method (Word)

Sets the specified fill to a one-color gradient.


## Syntax

 _expression_ . **OneColorGradient**( **_Style_** , **_Variant_** , **_Degree_** )

 _expression_ Required. A variable that represents a **[FillFormat](fillformat-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Style_|Required| **MsoGradientStyle**|The gradient style. Can be any  **MsoGradientStyle** constant except **msoGradientFromTitle** which applies only to Microsoft PowerPoint.|
| _Variant_|Required| **Long**|The gradient variant. Can be a value from 1 to 4, corresponding to the four variants on the  **Gradient** tab in the **Fill Effects** dialog box. If Style is **msoGradientFromCenter** , this argument can be either 1 or 2.|
| _Degree_|Required| **Single**|The gradient degree. Can be a value from 0.0 (dark) to 1.0 (light).|

## Example

This example adds a rectangle with a one-color gradient fill to the active document.


```vb
With ActiveDocument
    
	.Shapes.AddShape(msoShapeRectangle, _ 
        90, 90, 90, 80).Fill 

    .ForeColor.RGB = RGB(0, 128, 128) 

    .OneColorGradient msoGradientHorizontal, 1, 1 

	End With
```


## See also


#### Concepts


[FillFormat Object](fillformat-object-word.md)

