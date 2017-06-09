---
title: Worksheet.ConsolidationOptions Property (Excel)
keywords: vbaxl10.chm175088
f1_keywords:
- vbaxl10.chm175088
ms.prod: excel
api_name:
- Excel.Worksheet.ConsolidationOptions
ms.assetid: 8c09aa4d-86fc-701f-3b89-f8e2be38b948
ms.date: 06/08/2017
---


# Worksheet.ConsolidationOptions Property (Excel)

Returns a three-element array of consolidation options, as shown in the following table. If the element is  **True** , that option is set. Read-only **Variant** .


## Syntax

 _expression_ . **ConsolidationOptions**

 _expression_ A variable that represents a **Worksheet** object.


## Remarks





|**Element**|**Meaning**|
|:-----|:-----|
|1|Use labels in top row|
|2|Use labels in left column|
|3|Create links to source data|

## Example

This example displays the consolidation options for Sheet1. The list appears on a new worksheet created by the example.


```vb
Set newSheet = Worksheets.Add 
aOptions = Worksheets("Sheet1").ConsolidationOptions 
newSheet.Range("A1").Value = "Use labels in top row" 
newSheet.Range("A2").Value = "Use labels in left column" 
newSheet.Range("A3").Value = "Create links to source data" 
For i = 1 To 3 
 If aOptions(i) = True Then 
 newSheet.Cells(i, 2).Value = "True" 
 Else 
 newSheet.Cells(i, 2).Value = "False" 
 End If 
Next i 
newSheet.Columns("A:B").AutoFit
```


## See also


#### Concepts


[Worksheet Object](worksheet-object-excel.md)

