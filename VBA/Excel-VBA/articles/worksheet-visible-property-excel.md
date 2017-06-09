---
title: Worksheet.Visible Property (Excel)
keywords: vbaxl10.chm174097
f1_keywords:
- vbaxl10.chm174097
ms.prod: excel
api_name:
- Excel.Worksheet.Visible
ms.assetid: 48860564-6079-932e-2cae-0802235be61e
ms.date: 06/08/2017
---


# Worksheet.Visible Property (Excel)

Returns or sets an  **[XlSheetVisibility](xlsheetvisibility-enumeration-excel.md)** value that determines whether the object is visible.


## Syntax

 _expression_ . **Visible**

 _expression_ A variable that represents a **Worksheet** object.


## Example

This example hides Sheet1.


```vb
Worksheets("Sheet1").Visible = False
```

This example makes Sheet1 visible.




```vb
Worksheets("Sheet1").Visible = True
```

This example makes every sheet in the active workbook visible.




```vb
For Each sh In Sheets 
 sh.Visible = True 
Next sh
```

This example creates a new worksheet and then sets its  **Visible** property to **xlVeryHidden** . To refer to the sheet, use its object variable, `newSheet`, as shown in the last line of the example. To use the  `newSheet` object variable in another procedure, you must declare it as a public variable ( `Public newSheet As Object`) in the first line of the module preceding any  **Sub** or **Function** procedure.




```vb
Set newSheet = Worksheets.Add 
newSheet.Visible = xlVeryHidden 
newSheet.Range("A1:D4").Formula = "=RAND()"
```


## See also


#### Concepts


[Worksheet Object](worksheet-object-excel.md)

