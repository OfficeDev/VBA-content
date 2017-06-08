---
title: OLEFormat Object (Excel)
keywords: vbaxl10.chm631072
f1_keywords:
- vbaxl10.chm631072
ms.prod: excel
api_name:
- Excel.OLEFormat
ms.assetid: 96ee06d8-e922-c48c-4406-bb2f5cbaa02a
ms.date: 06/08/2017
---


# OLEFormat Object (Excel)

Contains OLE object properties.


## Remarks

If the  **[Shape](shape-object-excel.md)** object doesn't represent a linked or embedded object, the **[OLEFormat](shape-oleformat-property-excel.md)** property fails.


## Example

Use the  **OLEFormat** property to return the **OLEFormat** object. The following example activates an OLE object in the **[Shapes](shapes-object-excel.md)** collection.


```vb
Worksheets(1).Shapes(1).OLEFormat.Activate
```


## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)


