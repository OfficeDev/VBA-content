---
title: MailMergeDataSource.Filters Property (Publisher)
keywords: vbapb10.chm6291463
f1_keywords:
- vbapb10.chm6291463
ms.prod: publisher
api_name:
- Publisher.MailMergeDataSource.Filters
ms.assetid: 7b8fa974-08e5-9691-c69d-314eb6a5c651
ms.date: 06/08/2017
---


# MailMergeDataSource.Filters Property (Publisher)

Returns a  **[MailMergeFilters](mailmergefilters-object-publisher.md)** object that represents filters applied to the mail merge or catalog merge data source.


## Syntax

 _expression_. **Filters**

 _expression_A variable that represents a  **MailMergeDataSource** object.


### Return Value

MailMergeFilters


## Example

This example adds a new filter that removes all records with a blank Region field and then applies the filter to the active publication. This example assumes that a mail merge data source is attached to the active publication.


```vb
Sub FilterDataSource() 
 With ActiveDocument.MailMerge.DataSource 
 .Filters.Add Column:="Region", _ 
 Comparison:=msoFilterComparisonIsBlank, _ 
 Conjunction:=msoFilterConjunctionAnd 
 .ApplyFilter 
 End With 
End Sub
```


