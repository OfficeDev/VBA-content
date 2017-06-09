---
title: Hyperlink.Range Property (Excel)
keywords: vbaxl10.chm536075
f1_keywords:
- vbaxl10.chm536075
ms.prod: excel
api_name:
- Excel.Hyperlink.Range
ms.assetid: 0fdc49ba-fd3f-1125-fe3c-481828b7319e
ms.date: 06/08/2017
---


# Hyperlink.Range Property (Excel)

Returns a  **[Range](range-object-excel.md)** object that represents the range the specified hyperlink is attached to.


## Syntax

 _expression_ . **Range**

 _expression_ A variable that represents a **Hyperlink** object.


## Example

The following example stores in a variable the address for the AutoFilter applied to the Crew worksheet.


```
rAddress = Worksheets("Crew").AutoFilter.Range.Address
```

This example scrolls through the workbook window until the hyperlink range is in the upper-left corner of the active window.




```vb
Workbooks(1).Activate 
Set hr = ActiveSheet.Hyperlinks(1).Range 
ActiveWindow.ScrollRow = hr.Row 
ActiveWindow.ScrollColumn = hr.Column
```


## See also


#### Concepts


[Hyperlink Object](hyperlink-object-excel.md)

