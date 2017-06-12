---
title: Application.CalculationState Property (Excel)
keywords: vbaxl10.chm133265
f1_keywords:
- vbaxl10.chm133265
ms.prod: excel
api_name:
- Excel.Application.CalculationState
ms.assetid: 2f424286-7757-12e2-77c2-c26cf7c4bcf1
ms.date: 06/08/2017
---


# Application.CalculationState Property (Excel)

Returns an  **[XlCalculationState](xlcalculationstate-enumeration-excel.md)** constant that indicates the calculation state of the application, for any calculations that are being performed in Microsoft Excel. Read-only.


## Syntax

 _expression_ . **CalculationState**

 _expression_ A variable that represents an **Application** object.


## Example

In this example, Microsoft Excel checks to see if any calculations are being performed. If no calculations are being performed, a message displays the calculation state as "Done". Otherwise, a message displays the calculation state as "Not Done".


```vb
Sub StillCalculating() 
 
 If Application.CalculationState = xlDone Then 
 MsgBox "Done" 
 Else 
 MsgBox "Not Done" 
 End If 
 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

