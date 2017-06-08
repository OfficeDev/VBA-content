---
title: DialogSheetView Object (Excel)
keywords: vbaxl10.chm786072
f1_keywords:
- vbaxl10.chm786072
ms.prod: excel
api_name:
- Excel.DialogSheetView
ms.assetid: d468b3e8-c73e-d94a-0902-193f6983d893
ms.date: 06/08/2017
---


# DialogSheetView Object (Excel)

Represents the current  **Dialog** sheet view in a workbook.


## Remarks

To access this object, you must have a dialog sheet that was developed in the active workbook. Without the dialog sheet, the view properties for the object return an empty string value.


## Example

The following example turns on the dialog sheet view for the active workbook.


```vb
Worksheets("Sheet1").DialogSheetView.Visible = True
```


## See also


#### Other resources



[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)

