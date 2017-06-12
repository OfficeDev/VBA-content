---
title: Workbook.WindowResize Event (Excel)
keywords: vbaxl10.chm503082
f1_keywords:
- vbaxl10.chm503082
ms.prod: excel
api_name:
- Excel.Workbook.WindowResize
ms.assetid: 6e473482-fe16-03a2-7a27-b0cd9535c3e6
ms.date: 06/08/2017
---


# Workbook.WindowResize Event (Excel)

Occurs when any workbook window is resized.


## Syntax

 _expression_ . **WindowResize**( **_Wn_** )

 _expression_ A variable that represents a **Workbook** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Wn_|Required| **Window**|The resized window.|

## Example

This example runs when any workbook window is resized.


```vb
Private Sub Workbook_WindowResize(ByVal Wn As Excel.Window) 
 Application.StatusBar = Wn.Caption &; " resized" 
End Sub
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

