---
title: Research Object (Excel)
keywords: vbaxl10.chm848072
f1_keywords:
- vbaxl10.chm848072
ms.prod: excel
api_name:
- Excel.Research
ms.assetid: de9d8a1d-4942-88f4-ba8c-30bd06e1f24b
ms.date: 06/08/2017
---


# Research Object (Excel)

Represents the controls of a  **Research** query.


## Remarks

When working with  **Research** queries, you must have an existing GUID that corresponds to a live data source. If the data source is unavailable or does not exist, a run-time error occurs.


## Example

The following example returns data from an existing data source and translates the information into working content.


```vb
Worksheets("Sheet1").Research.Translate = True
```


## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)


