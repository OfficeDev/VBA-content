---
title: Application.Worksheets Property (Excel)
keywords: vbaxl10.chm132116
f1_keywords:
- vbaxl10.chm132116
ms.prod: excel
api_name:
- Excel.Application.Worksheets
ms.assetid: ee9350d3-f24e-ed40-b267-8101d3267b4d
ms.date: 06/08/2017
---


# Application.Worksheets Property (Excel)

For an  **Application** object, returns a **[Sheets](sheets-object-excel.md)** collection that represents all the worksheets in the active workbook. For a **Workbook** object, returns a **[Sheets](sheets-object-excel.md)** collection that represents all the worksheets in the specified workbook. Read-only **Sheets** object.


## Syntax

 _expression_ . **Worksheets**

 _expression_ A variable that represents an **Application** object.


## Remarks

Using this property without an object qualifier returns all the worksheets in the active workbook.

This property doesn't return macro sheets; use the  **[Excel4MacroSheets](application-excel4macrosheets-property-excel.md)** property or the **[Excel4IntlMacroSheets](application-excel4intlmacrosheets-property-excel.md)** property to return those sheets.


## Example

This example displays the value in cell A1 on Sheet1 in the active workbook.


```vb
MsgBox Worksheets("Sheet1").Range("A1").Value
```

This example displays the name of each worksheet in the active workbook.




```vb
For Each ws In Worksheets 
 MsgBox ws.Name 
Next ws
```

This example adds a new worksheet to the active workbook and then sets the name of the worksheet.




```vb
Set newSheet = Worksheets.Add 
newSheet.Name = "current Budget"
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

