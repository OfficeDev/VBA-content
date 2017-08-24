---
title: MailMergeFilterCriterion.Column Property (Publisher)
keywords: vbapb10.chm6815747
f1_keywords:
- vbapb10.chm6815747
ms.prod: publisher
api_name:
- Publisher.MailMergeFilterCriterion.Column
ms.assetid: 000b4b4c-73a1-ea9f-6f44-bc6eac15cb4b
ms.date: 06/08/2017
---


# MailMergeFilterCriterion.Column Property (Publisher)

Returns a  **String** that represents the name of the field in the mail merge data source to use in the filter. Read/write.


## Syntax

 _expression_. **Column**

 _expression_A variable that represents a  **MailMergeFilterCriterion** object.


## Example

The following example changes an existing filter to remove from the mail merge all records that do not have a Region field equal to "WA".


```vb
Sub SetQueryCriterion() 
 Dim intItem As Integer 
 With ActiveDocument.MailMerge.DataSource.Filters 
 For intItem = 1 To .Count 
 With .Item(intItem) 
 If .Column = "Region" Then 
 .Comparison = msoFilterComparisonNotEqual 
 .CompareTo = "WA" 
 If .Conjunction = "Or" Then .Conjunction = "And" 
 End If 
 End With 
 Next intItem 
 End With 
End Sub
```


