---
title: Chart.Select Event (Excel)
keywords: vbaxl10.chm500083
f1_keywords:
- vbaxl10.chm500083
ms.prod: excel
api_name:
- Excel.Chart.Select
ms.assetid: 00ea6501-e92e-5b95-f2b0-bb9b014bb5ec
ms.date: 06/08/2017
---


# Chart.Select Event (Excel)

Occurs when a chart element is selected.


## Syntax

 _expression_ . **Select**( **_ElementID_** , **_Arg1_** , **_Arg2_** )

 _expression_ An expression that returns a **Chart** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ElementID_|Required| **Long**|The selected chart element. For more information about these arguments, see the  **[BeforeDoubleClick](chart-beforedoubleclick-event-excel.md)** event.|
| _Arg1_|Required| **Long**|The selected chart element. For more information about these arguments, see the  **BeforeDoubleClick** event.|
| _Arg2_|Required| **Long**|The selected chart element. For more information about these arguments, see the  **BeforeDoubleClick** event.|

## Example

This example displays a message box if the user selects the chart title.


```vb
Private Sub Chart_Select(ByVal ElementID As Long, _ 
 ByVal Arg1 As Long, ByVal Arg2 As Long) 
 If ElementId = xlChartTitle Then 
 MsgBox "please don't change the chart title" 
 End If 
End Sub
```


## See also


#### Concepts


[Chart Object](chart-object-excel.md)

