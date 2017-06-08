---
title: Workbook.RunAutoMacros Method (Excel)
keywords: vbaxl10.chm199143
f1_keywords:
- vbaxl10.chm199143
ms.prod: excel
api_name:
- Excel.Workbook.RunAutoMacros
ms.assetid: 85dfdadf-75e6-437d-fb7a-e17681a69b35
ms.date: 06/08/2017
---


# Workbook.RunAutoMacros Method (Excel)

Runs the Auto_Open, Auto_Close, Auto_Activate, or Auto_Deactivate macro attached to the workbook. This method is included for backward compatibility. For new Visual Basic code, you should use the Open, Close, Activate and Deactivate events instead of these macros.


## Syntax

 _expression_ . **RunAutoMacros**( **_Which_** )

 _expression_ A variable that represents a **Workbook** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Which_|Required| **[XlRunAutoMacro](xlrunautomacro-enumeration-excel.md)**|Specifies the automatic macro to run.|

## Remarks





| **XlRunAutoMacro** can be one of these **XlRunAutoMacro** constants.|
| **xlAutoActivate** . Auto_Activate macros|
| **xlAutoClose** . Auto_Close macros|
| **xlAutoDeactivate** . Auto_Deactivate macros|
| **xlAutoOpen** . Auto_Open macros|

## Example

This example opens the workbook Analysis.xls and then runs its Auto_Open macro.


```vb
Workbooks.Open "ANALYSIS.XLS" 
ActiveWorkbook.RunAutoMacros xlAutoOpen
```

This example runs the Auto_Close macro for the active workbook and then closes the workbook.




```vb
With ActiveWorkbook 
 .RunAutoMacros xlAutoClose 
 .Close 
End With
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

