---
title: FillFormat.PresetGradientType Property (Excel)
keywords: vbaxl10.chm115018
f1_keywords:
- vbaxl10.chm115018
ms.prod: excel
api_name:
- Excel.FillFormat.PresetGradientType
ms.assetid: e9cb1ba6-9c40-3fef-7014-68069be4da1f
ms.date: 06/08/2017
---


# FillFormat.PresetGradientType Property (Excel)

Returns the preset gradient type for the specified fill. Read-only  **[MsoPresetGradientType](http://msdn.microsoft.com/library/d0ee19e7-bdd3-3102-61b4-dbb17d5c0363%28Office.15%29.aspx)** .


## Syntax

 _expression_ . **PresetGradientType**

 _expression_ A variable that represents a **FillFormat** object.


## Remarks

Use the  **[PresetGradient](fillformat-presetgradient-method-excel.md)** method to set the preset gradient type for the fill.


## Example

This example sets the fill format for chart two to the same style used for chart one.


```vb
Set c1f = Charts(1).ChartArea.Fill 
If c1f.Type = msoFillGradient Then 
    With Charts(2).ChartArea.Fill 
        .Visible = True 
        .PresetGradient c1f.GradientStyle, _ 
            c1f.GradientVariant, c1f.PresetGradientType 
    End With 
End If
```


## See also


#### Concepts


[FillFormat Object](fillformat-object-excel.md)

