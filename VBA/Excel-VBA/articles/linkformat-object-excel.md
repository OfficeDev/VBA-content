---
title: LinkFormat Object (Excel)
keywords: vbaxl10.chm633072
f1_keywords:
- vbaxl10.chm633072
ms.prod: excel
api_name:
- Excel.LinkFormat
ms.assetid: 3d8085bf-c113-7cbe-871b-01f3b6017824
ms.date: 06/08/2017
---


# LinkFormat Object (Excel)

Contains linked OLE object properties.


## Remarks

If the  **[Shape](shape-object-excel.md)** object doesn't represent a linked object, the **[LinkFormat](shape-linkformat-property-excel.md)** property fails.


## Example

Use the  **LinkFormat** property to return the **LinkFormat** object. The following example updates an OLE object in the **[Shapes](shapes-object-excel.md)** collection.


```vb
Worksheets(1).Shapes(1).LinkFormat.Update
```


## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)


