---
title: Slicer.Name Property (Excel)
keywords: vbaxl10.chm905073
f1_keywords:
- vbaxl10.chm905073
ms.prod: excel
api_name:
- Excel.Slicer.Name
ms.assetid: cc8508d3-82fc-365b-c632-2565fd0071c5
ms.date: 06/08/2017
---


# Slicer.Name Property (Excel)

Returns or sets the name of the specified slicer. Read/write.


## Syntax

 _expression_ . **Name**

 _expression_ A variable that represents a **[Slicer](slicer-object-excel.md)** object.


### Return Value

String


## Remarks

The name must be unique across all slicers within a workbook. The default name uses the text the field name of the PivotField on which the slicer is based, and if necessary, appends a space and number to make the name unique.


## See also


#### Concepts


[Slicer Object](slicer-object-excel.md)

