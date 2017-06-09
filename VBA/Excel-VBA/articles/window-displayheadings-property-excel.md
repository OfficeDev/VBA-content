---
title: Window.DisplayHeadings Property (Excel)
keywords: vbaxl10.chm356084
f1_keywords:
- vbaxl10.chm356084
ms.prod: excel
api_name:
- Excel.Window.DisplayHeadings
ms.assetid: 7105f3a4-2322-c796-5ca6-59ea46d2e248
ms.date: 06/08/2017
---


# Window.DisplayHeadings Property (Excel)

 **True** if both row and column headings are displayed; **False** if no headings are displayed. Read/write **Boolean** .


## Syntax

 _expression_ . **DisplayHeadings**

 _expression_ A variable that represents a **Window** object.


## Remarks

This property applies only to worksheets and macro sheets.

This property affects only displayed headings. Use the  **[PrintHeadings](pagesetup-printheadings-property-excel.md)** property to control the printing of headings.


## Example

This example turns off the display of row and column headings in the active window in Book1.xls.


```vb
Workbooks("BOOK1.XLS").Worksheets("Sheet1").Activate 
ActiveWindow.DisplayHeadings = False 

```


## See also


#### Concepts


[Window Object](window-object-excel.md)

