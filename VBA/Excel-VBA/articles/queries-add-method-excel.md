---
title: Queries.Add Method (Excel)
keywords: vbaxl10.chm976074
f1_keywords:
- vbaxl10.chm976074
ms.assetid: 184711c0-2ce4-ba6e-df56-1f7fdd60ab2c
ms.date: 06/08/2017
ms.prod: excel
---


# Queries.Add Method (Excel)

Adds a new [WorkbookQuery](workbookquery-object-excel.md) object to the **Queries** collection.


## Syntax

 _expression_ . **Add**( _Name_,  _Formula_,  _Description_)

 _expression_ A variable that represents a **Queries** object.


### Parameters



| _Name_|Required|STRING|The name of the query.|
| _Formula_|Required|STRING|The Power Query M formula for the new query.|
| _Description_|Optional|VARIANT|The description of the query.|

### Return Value

[WorkbookQuery](workbookquery-object-excel.md)


## Example

The following example shows how to add a query to a workbook from an existing CSV file.


```vb
Dim myConnection As WorkbookConnection
Dim mFormula As String
mFormula = _
"let Source = Csv.Document(File.Contents(""C:\data.txt""),null,""#(tab)"",null,1252) in Source"
query1 = ActiveWorkbook.Queries.Add(?query1?, mFormula)

```


## See also


#### Other resources


[Queries Object](queries-object-excel.md)


