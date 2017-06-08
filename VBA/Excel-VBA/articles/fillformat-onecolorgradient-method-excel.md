---
title: FillFormat.OneColorGradient Method (Excel)
keywords: vbaxl10.chm115003
f1_keywords:
- vbaxl10.chm115003
ms.prod: excel
api_name:
- Excel.FillFormat.OneColorGradient
ms.assetid: dc44ddab-7aee-acd9-1008-1a9bbae13829
ms.date: 06/08/2017
---


# FillFormat.OneColorGradient Method (Excel)

Sets the specified fill to a one-color gradient.


## Syntax

 _expression_ . **OneColorGradient**( **_Style_** , **_Variant_** , **_Degree_** )

 _expression_ A variable that represents a **FillFormat** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Style_|Required| **[MsoGradientStyle](http://msdn.microsoft.com/library/1f0e723f-293c-3646-fd77-da2c8842c71f%28Office.15%29.aspx)**|The gradient style.|
| _Variant_|Required| **Integer**|The gradient variant. Can be a value from 1 through 4, corresponding to one of the four variants on the  **Gradient** tab in the **Fill Effects** dialog box. If _GradientStyle_ is **msoGradientFromCenter** , the _Variant_ argument can only be 1 or 2.|
| _Degree_|Required| **Single**|The gradient degree. Can be a value from 0.0 (dark) through 1.0 (light).|

## See also


#### Concepts


[FillFormat Object](fillformat-object-excel.md)

