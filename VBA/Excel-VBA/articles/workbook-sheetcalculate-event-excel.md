---
title: Workbook.SheetCalculate Event (Excel)
keywords: vbaxl10.chm503090
f1_keywords:
- vbaxl10.chm503090
ms.prod: excel
api_name:
- Excel.Workbook.SheetCalculate
ms.assetid: 0610bfa5-15dc-a57f-f362-cf897bd54b91
ms.date: 06/08/2017
---


# Workbook.SheetCalculate Event (Excel)

Occurs after any worksheet is recalculated or after any changed data is plotted on a chart.


## Syntax

 _expression_ . **SheetCalculate**( **_Sh_** , )

 _expression_ An expression that returns a **Workbook** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Sh_|Required| **Object**|Can be a  **[Chart](chart-object-excel.md)** or **[Worksheet](worksheet-object-excel.md)** object.|

## Example

This example sorts the range A1:A100 on worksheet one when any sheet in the workbook is calculated.


```vb
Private Sub Workbook_SheetCalculate(ByVal Sh As Object) 
 With Worksheets(1) 
 .Range("a1:a100").Sort Key1:=.Range("a1") 
 End With 
End Sub
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

