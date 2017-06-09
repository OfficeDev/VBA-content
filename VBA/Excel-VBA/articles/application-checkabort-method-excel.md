---
title: Application.CheckAbort Method (Excel)
keywords: vbaxl10.chm133279
f1_keywords:
- vbaxl10.chm133279
ms.prod: excel
api_name:
- Excel.Application.CheckAbort
ms.assetid: e407aeff-b401-029a-9ada-8f11eef54fb0
ms.date: 06/08/2017
---


# Application.CheckAbort Method (Excel)

Stops recalculation in a Microsoft Excel application.


## Syntax

 _expression_ . **CheckAbort**( **_KeepAbort_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _KeepAbort_|Optional| **Boolean**|Allows recalculation to be performed for a Range.|

## Example

In this example, Excel stops recalculation in the application, except for cell A10. For you to be able to see the results of this example, other calculations should exist in the application that will allow you to see the differences between the cell designated to continue recalculating and other cells.


```vb
Sub UseCheckAbort() 
 
 Dim rngSubtotal As Variant 
 Set rngSubtotal = Application.Range("A10") 
 
 ' Stop recalculation except for designated cell. 
 Application.CheckAbort KeepAbort:=rngSubtotal 
 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

