---
title: Workbook.BeforePrint Event (Excel)
keywords: vbaxl10.chm503078
f1_keywords:
- vbaxl10.chm503078
ms.prod: excel
api_name:
- Excel.Workbook.BeforePrint
ms.assetid: 2c97cb32-2bb3-2848-b5ed-32d9129af080
ms.date: 06/08/2017
---


# Workbook.BeforePrint Event (Excel)

Occurs before the workbook (or anything in it) is printed.


## Syntax

 _expression_ . **BeforePrint**( **_Cancel_** )

 _expression_ A variable that represents a **Workbook** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True** , the workbook isn't printed when the procedure is finished.|

### Return Value

Nothing


## Example

This example recalculates all worksheets in the active workbook before printing anything.


```vb
Private Sub Workbook_BeforePrint(Cancel As Boolean) 
 For Each wk in Worksheets 
 wk.Calculate 
 Next 
End Sub
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

