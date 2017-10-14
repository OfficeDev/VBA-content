---
title: Chart.MouseDown Event (Excel)
keywords: vbaxl10.chm500076
f1_keywords:
- vbaxl10.chm500076
ms.prod: excel
api_name:
- Excel.Chart.MouseDown
ms.assetid: 6c4ef5ce-560e-a7d5-c602-99a999fb5535
ms.date: 06/08/2017
---


# Chart.MouseDown Event (Excel)

Occurs when a mouse button is pressed while the pointer is over a chart.


## Syntax

 _expression_ . **MouseDown**( **_Button_** , **_Shift_** , **_x_** , **_y_** )

 _expression_ A variable that represents a **Chart** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Button_|Required| **Long**|The mouse button that was released. Can be one of the following  **[XlMouseButton](xlmousebutton-enumeration-excel.md)** constants: **xlNoButton** , **xlPrimaryButton** , or **xlSecondaryButton** .|
| _Shift_|Required| **Long**|The state of the SHIFT, CTRL, and ALT keys when the event occurred. Can be one of or a sum of values.|
| _x_|Required| **Long**|The X coordinate of the mouse pointer in chart object client coordinates.|
| _y_|Required| **Long**|The Y coordinate of the mouse pointer in chart object client coordinates.|

### Return Value

Nothing


## Remarks

The following table specifies the values for the  _Shift_ parameter.



|**Value**|**Meaning**|
|:-----|:-----|
|0 (zero)|No keys|
|1|SHIFT key|
|2|CTRL key|
|4|ALT key|

## Example

This example runs when a mouse button is pressed while the pointer is over a chart.


```vb
Private Sub Chart_MouseDown(ByVal Button As Long, _ 
 ByVal Shift As Long, ByVal X As Long, ByVal Y As Long) 
 MsgBox "Button = " &; Button &; chr$(13) &; _ 
 "Shift = " &; Shift &; chr$(13) &; _ 
 "X = " &; X &; " Y = " &; Y 
End Sub
```


## See also


#### Concepts


[Chart Object](chart-object-excel.md)

