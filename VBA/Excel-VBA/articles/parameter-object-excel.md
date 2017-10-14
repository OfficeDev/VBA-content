---
title: Parameter Object (Excel)
keywords: vbaxl10.chm522072
f1_keywords:
- vbaxl10.chm522072
ms.prod: excel
api_name:
- Excel.Parameter
ms.assetid: 2a30f4ef-2cae-c96d-4480-3ba55fa871e8
ms.date: 06/08/2017
---


# Parameter Object (Excel)

Represents a single parameter used in a parameter query.


## Remarks

 The **Parameter** object is a member of the **[Parameters](parameters-object-excel.md)** collection.


## Example

Use  **[Parameters](querytable-parameters-property-excel.md)** ( _index_ ), where _index_ is the index number of the parameter, to return a single **Parameter** object. The following example modifies the prompt string for parameter one.


```vb
With Worksheets(1).QueryTables(1).Parameters(1) 
 .SetParam xlPrompt, "Please " &; .PromptString 
End With
```


## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)


