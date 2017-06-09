---
title: CellFormat Object (Excel)
keywords: vbaxl10.chm675072
f1_keywords:
- vbaxl10.chm675072
ms.prod: excel
api_name:
- Excel.CellFormat
ms.assetid: da4e50b9-6d5b-22e1-3113-0d1ea6686272
ms.date: 06/08/2017
---


# CellFormat Object (Excel)

Represents the search criteria for the cell format.


## Remarks

Use the  **[FindFormat](application-findformat-property-excel.md)** or **[ReplaceFormat](application-replaceformat-property-excel.md)** properties of the **[Application](application-object-excel.md)** object to return a **CellFormat** object.

With a  **CellFormat** object, you can use the **[Borders](cellformat-borders-property-excel.md)**, **[Font](cellformat-font-property-excel.md)**, or **[Interior](cellformat-interior-property-excel.md)** properties of the **CellFormat** object, to define the search criteria for the cell format.


## Example

The following example sets the search criteria for the interior of the cell format. 


```
Sub ChangeCellFormat() 
 
 ' Set the interior of cell A1 to yellow. 
 Range("A1").Select 
 Selection.Interior.ColorIndex = 36 
 MsgBox "The cell format for cell A1 is a yellow interior." 
 
 ' Set the CellFormat object to replace yellow with green. 
 With Application 
 .FindFormat.Interior.ColorIndex = 36 
 .ReplaceFormat.Interior.ColorIndex = 35 
 End With 
 
 ' Find and replace cell A1's yellow interior with green. 
 ActiveCell.Replace What:="", Replacement:="", LookAt:=xlPart, _ 
 SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=True, _ 
 ReplaceFormat:=True 
 MsgBox "The cell format for cell A1 is replaced with a green interior." 
 
End Sub
```


## Methods



|**Name**|
|:-----|
|[Clear](cellformat-clear-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[AddIndent](cellformat-addindent-property-excel.md)|
|[Application](cellformat-application-property-excel.md)|
|[Borders](cellformat-borders-property-excel.md)|
|[Creator](cellformat-creator-property-excel.md)|
|[Font](cellformat-font-property-excel.md)|
|[FormulaHidden](cellformat-formulahidden-property-excel.md)|
|[HorizontalAlignment](cellformat-horizontalalignment-property-excel.md)|
|[IndentLevel](cellformat-indentlevel-property-excel.md)|
|[Interior](cellformat-interior-property-excel.md)|
|[Locked](cellformat-locked-property-excel.md)|
|[MergeCells](cellformat-mergecells-property-excel.md)|
|[NumberFormat](cellformat-numberformat-property-excel.md)|
|[NumberFormatLocal](cellformat-numberformatlocal-property-excel.md)|
|[Orientation](cellformat-orientation-property-excel.md)|
|[Parent](cellformat-parent-property-excel.md)|
|[ShrinkToFit](cellformat-shrinktofit-property-excel.md)|
|[VerticalAlignment](cellformat-verticalalignment-property-excel.md)|
|[WrapText](cellformat-wraptext-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
