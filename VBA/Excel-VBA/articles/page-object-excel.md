---
title: Page Object (Excel)
keywords: vbaxl10.chm830072
f1_keywords:
- vbaxl10.chm830072
ms.prod: excel
api_name:
- Excel.Page
ms.assetid: debd4537-af71-8699-b714-6854c3cf0fad
ms.date: 06/08/2017
---


# Page Object (Excel)

Represents a page in a workbook. Use the  **PageSetup** object and the related methods and properties for programmatically defining page layout in a workbook.


## Remarks

Use the  **Item** method to access a specific page in a workbook. The following example accesses the first page in the active workbook.


```vb
Dim objPage As Page 
 
Set objPage = ActiveWorkbook.ActiveWindow _ 
 .Panes(1).Pages.Item(1)
```


## See also


#### Other resources



[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)

