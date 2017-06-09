---
title: Workbook.IsAddin Property (Excel)
keywords: vbaxl10.chm199106
f1_keywords:
- vbaxl10.chm199106
ms.prod: excel
api_name:
- Excel.Workbook.IsAddin
ms.assetid: b8c8b9f4-4be5-0260-957e-c6450f31a0c0
ms.date: 06/08/2017
---


# Workbook.IsAddin Property (Excel)

 **True** if the workbook is running as an add-in. Read/write **Boolean** .


## Syntax

 _expression_ . **IsAddin**

 _expression_ A variable that represents a **Workbook** object.


## Remarks

When you set this property to  **True** , the workbook has the following characteristics:


- You won't be prompted to save the workbook if changes are made while the workbook is open.
    
- The workbook window won't be visible.
    
- Any macros in the workbook won't be visible in the  **Macro** dialog box (displayed by pointing to **Macro** on the **Tools** menu and clicking **Macros** ).
    
- Macros in the workbook can still be run from the  **Macro** dialog box even though they're not visible. In addition, macro names don't need to be qualified with the workbook name.
    
- Holding down the SHIFT key when you open the workbook has no effect.
    

## Example

This example runs a section of code if the workbook is an add-in.


```vb
If ThisWorkbook.IsAddin Then 
 ' this code runs when the workbook is an add-in 
End If
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

