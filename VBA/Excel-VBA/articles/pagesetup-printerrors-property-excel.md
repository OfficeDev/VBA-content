---
title: PageSetup.PrintErrors Property (Excel)
keywords: vbaxl10.chm473105
f1_keywords:
- vbaxl10.chm473105
ms.prod: excel
api_name:
- Excel.PageSetup.PrintErrors
ms.assetid: 4a864a1e-cbdb-8ef7-536d-d2c5f518f9db
ms.date: 06/08/2017
---


# PageSetup.PrintErrors Property (Excel)

Sets or returns an  **[XlPrintErrors](xlprinterrors-enumeration-excel.md)** contstant specifying the type of print error displayed. This feature allows users to suppress the display of error values when printing a worksheet. Read/write .


## Syntax

 _expression_ . **PrintErrors**

 _expression_ A variable that represents a **PageSetup** object.


## Remarks





| **XlPrintErrors** can be one of these **XlPrintErrors** constants.|
| **xlPrintErrorsBlank**|
| **xlPrintErrorsDash**|
| **xlPrintErrorsDisplayed**|
| **xlPrintErrorsNA**|

## Example

In this example, Microsoft Excel uses a formula that returns an error in the active worksheet. The  **PrintErrors** property is set to display dashes. A Print Preview window displays the dashes for the print error. This example assumes a printer driver has been installed.


```vb
Sub UsePrintErrors() 
 
 Dim wksOne As Worksheet 
 
 Set wksOne = Application.ActiveSheet 
 
 ' Create a formula that returns an error value. 
 Range("A1").Value = 1 
 Range("A2").Value = 0 
 Range("A3").Formula = "=A1/A2" 
 
 ' Change print errors to display dashes. 
 wksOne.PageSetup.PrintErrors = xlPrintErrorsDash 
 
 ' Use the Print Preview window to see the dashes used for print errors. 
 ActiveWindow.SelectedSheets.PrintPreview 
 
End Sub
```


## See also


#### Concepts


[PageSetup Object](pagesetup-object-excel.md)

