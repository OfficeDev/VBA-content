---
title: Application.WindowDeactivate Event (Excel)
keywords: vbaxl10.chm504092
f1_keywords:
- vbaxl10.chm504092
ms.prod: excel
api_name:
- Excel.Application.WindowDeactivate
ms.assetid: 6adcba54-3d4a-f780-915e-5798303faf60
ms.date: 06/08/2017
---


# Application.WindowDeactivate Event (Excel)

Occurs when any workbook window is deactivated.


## Syntax

 _expression_ . **WindowDeactivate**( **_Wb_** , **_Wn_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Wb_|Required| **Workbook**|The workbook displayed in the deactivated window.|
| _Wn_|Required| **Window**|The deactivated window.|

## Example

This example minimizes any workbook window when it's deactivated.


```vb
Private Sub Workbook_WindowDeactivate(ByVal Wn As Excel.Window) 
 Wn.WindowState = xlMinimized 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

