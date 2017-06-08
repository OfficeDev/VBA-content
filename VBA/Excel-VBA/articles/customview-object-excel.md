---
title: CustomView Object (Excel)
keywords: vbaxl10.chm507072
f1_keywords:
- vbaxl10.chm507072
ms.prod: excel
api_name:
- Excel.CustomView
ms.assetid: e16b1920-faeb-62d4-4d27-59745c4f5355
ms.date: 06/08/2017
---


# CustomView Object (Excel)

Represents a custom workbook view.


## Remarks

 The **CustomView** object is a member of the **[CustomViews](customviews-object-excel.md)** collection.


## Example

Use  **CustomViews** ( _index_ ), where _index_ is the name or index number of the custom view, to return a **CustomView** object. The following example shows the custom view named "Current Inventory."


```vb
ThisWorkbook.CustomViews("Current Inventory").Show
```


## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)


