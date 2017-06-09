---
title: FillFormat.PresetGradient Method (Excel)
keywords: vbaxl10.chm115005
f1_keywords:
- vbaxl10.chm115005
ms.prod: excel
api_name:
- Excel.FillFormat.PresetGradient
ms.assetid: 0bcebb14-7f39-d20c-6701-76355c212f99
ms.date: 06/08/2017
---


# FillFormat.PresetGradient Method (Excel)

Sets the specified fill to a preset gradient.


## Syntax

 _expression_ . **PresetGradient**( **_Style_** , **_Variant_** , **_PresetGradientType_** )

 _expression_ A variable that represents a **FillFormat** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Style_|Required| **[MsoGradientStyle](http://msdn.microsoft.com/library/1f0e723f-293c-3646-fd77-da2c8842c71f%28Office.15%29.aspx)**|The gradient style.|
| _Variant_|Required| **Integer**|The gradient variant. Can be a value from 1 through 4, corresponding to one of the four variants on the  **Gradient** tab in the **Fill Effects** dialog box. If _GradientStyle_ is **msoGradientFromCenter** , the _Variant_ argument can only be 1 or 2.|
| _PresetGradientType_|Required| ** MsoPresetGradientType**|The preset gradient type.|

## See also


#### Concepts


[FillFormat Object](fillformat-object-excel.md)

