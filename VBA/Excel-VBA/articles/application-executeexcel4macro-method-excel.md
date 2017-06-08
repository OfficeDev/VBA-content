---
title: Application.ExecuteExcel4Macro Method (Excel)
keywords: vbaxl10.chm132097
f1_keywords:
- vbaxl10.chm132097
ms.prod: excel
api_name:
- Excel.Application.ExecuteExcel4Macro
ms.assetid: 0afa77ab-43e0-0120-4ffd-25e290c72f6c
ms.date: 06/08/2017
---


# Application.ExecuteExcel4Macro Method (Excel)

Runs a Microsoft Excel 4.0 macro function and then returns the result of the function. The return type depends on the function.


## Syntax

 _expression_ . **ExecuteExcel4Macro**( **_String_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _String_|Required| **String**|A Microsoft Excel 4.0 macro language function without the equal sign. All references must be given as R1C1 strings. If  _String_ contains embedded double quotation marks, you must double them. For example, to run the macro function =MID("sometext",1,4), _String_ would have to be "MID(""sometext"",1,4)".|

### Return Value

Variant


## Remarks

The Microsoft Excel 4.0 macro isn't evaluated in the context of the current workbook or sheet. This means that any references should be external and should specify an explicit workbook name. For example, to run the Microsoft Excel 4.0 macro "My_Macro" in Book1 you must use "Book1!My_Macro()". If you don't specify the workbook name, this method fails.


## Example

This example runs the  **GET.CELL(42)** macro function on cell C3 on Sheet1 and then displays the result in a message box. The **GET.CELL(42)** macro function returns the horizontal distance from the left edge of the active window to the left edge of the active cell. This macro function has no direct Visual Basic equivalent.


```vb
Worksheets("Sheet1").Activate 
Range("C3").Select 
MsgBox ExecuteExcel4Macro("GET.CELL(42)")
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

