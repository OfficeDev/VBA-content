---
title: PivotCell.PivotCellType Property (Excel)
keywords: vbaxl10.chm692073
f1_keywords:
- vbaxl10.chm692073
ms.prod: excel
api_name:
- Excel.PivotCell.PivotCellType
ms.assetid: f5462981-924c-4d6c-be99-5b7cea0222a4
ms.date: 06/08/2017
---


# PivotCell.PivotCellType Property (Excel)

Returns one of the  **[XlPivotCellType](xlpivotcelltype-enumeration-excel.md)** constants that identifies the PivotTable entity the cell corresponds to. Read-only.


## Syntax

 _expression_ . **PivotCellType**

 _expression_ A variable that represents a **PivotCell** object.


## Remarks





| **XlPivotCellType** can be one of these **XlPivotCellType** constants.|
| **xlPivotCellBlankCell** A structural blank cell in the PivotTable.|
| **xlPivotCellCustomSubtotal** A cell in the row or column area that is a custom subtotal.|
| **xlPivotCellDataField** A data field label (not the **Data** button).|
| **xlPivotCellDataPivotField** The **Data** button.|
| **xlPivotCellGrandTotal** A cell in a row or column area which is a grand total.|
| **xlPivotCellPageFieldItem** The cell that shows the selected item of a Page field.|
| **xlPivotCellPivotField** The button for a field (not the **Data** button).|
| **xlPivotCellPivotItem** A cell in the row or column area which is not a subtotal, grand total, custom subtotal, or blank line.|
| **xlPivotCellSubtotal** A cell in the row or column area which is a subtotal.|
| **xlPivotCellValue** Any cell in the data area (except a blank row).|

## Example

This example determines if cell A5 in the PivotTable is a data item and notifies the user. The example assumes a PivotTable exists on the active worksheet and cell A5 is contained in the PivotTable. If cell A5 is not in the PivotTable, the example handles the run-time error.


```vb
Sub CheckPivotCellType() 
 
 On Error GoTo Not_In_PivotTable 
 
 ' Determine if cell A5 is a data item in the PivotTable. 
 If Application.Range("A5").PivotCell.PivotCellType = xlPivotCellValue Then 
 MsgBox "The cell at A5 is a data item." 
 Else 
 MsgBox "The cell at A5 is not a data item." 
 End If 
 Exit Sub 
 
Not_In_PivotTable: 
 MsgBox "The chosen cell is not in a PivotTable." 
 
End Sub
```


## See also


#### Concepts


[PivotCell Object](pivotcell-object-excel.md)

