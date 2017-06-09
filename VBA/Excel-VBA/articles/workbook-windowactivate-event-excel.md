---
title: Workbook.WindowActivate Event (Excel)
keywords: vbaxl10.chm503083
f1_keywords:
- vbaxl10.chm503083
ms.prod: excel
api_name:
- Excel.Workbook.WindowActivate
ms.assetid: e99d955c-1975-44c3-05b3-3aa6e851083c
ms.date: 06/08/2017
---


# Workbook.WindowActivate Event (Excel)

Occurs when any workbook window is activated.


## Syntax

 _expression_ . **WindowActivate**( **_Wn_** )

 _expression_ A variable that represents a **Workbook** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Wn_|Required| **Window**| The activated window.|

## Example

This example maximizes any workbook window when it's activated.


```vb
Private Sub Workbook_WindowActivate(ByVal Wn As Excel.Window) 
 Wn.WindowState = xlMaximized 
End Sub
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

