---
title: Workbook.NewSheet Event (Excel)
keywords: vbaxl10.chm503079
f1_keywords:
- vbaxl10.chm503079
ms.prod: excel
api_name:
- Excel.Workbook.NewSheet
ms.assetid: 5abb254d-a2c3-7dac-e79f-0de74a081ecd
ms.date: 06/08/2017
---


# Workbook.NewSheet Event (Excel)

Occurs when a new sheet is created in the workbook.


## Syntax

 _expression_ . **NewSheet**( **_Sh_** )

 _expression_ A variable that represents a **Workbook** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Sh_|Required| **Object**|The new sheet. Can be a  **[Worksheet](worksheet-object-excel.md)** or **[Chart](chart-object-excel.md)** object.|

### Return Value

Nothing


## Example

This example moves new sheets to the end of the workbook.


```vb
Private Sub Workbook_NewSheet(ByVal Sh as Object) 
 Sh.Move After:= Sheets(Sheets.Count) 
End Sub
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

