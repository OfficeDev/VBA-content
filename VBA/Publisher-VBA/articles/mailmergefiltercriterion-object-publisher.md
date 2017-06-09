---
title: MailMergeFilterCriterion Object (Publisher)
keywords: vbapb10.chm6881279
f1_keywords:
- vbapb10.chm6881279
ms.prod: publisher
api_name:
- Publisher.MailMergeFilterCriterion
ms.assetid: 2814890f-009b-b277-3ea4-c1f167a5e1c9
ms.date: 06/08/2017
---


# MailMergeFilterCriterion Object (Publisher)

Represents a filter to be applied to an attached mail merge or catalog merge data source. The  **MailMergeFilterCriterion** object is a member of the **MailMergeFilters** object.
 


## Example

Each filter is a line in a query string. Use the  **[Column](mailmergefiltercriterion-column-property-publisher.md)**, **[Comparison](mailmergefiltercriterion-comparison-property-publisher.md)**, **[CompareTo](mailmergefiltercriterion-compareto-property-publisher.md)**, and **[Conjunction](mailmergefiltercriterion-conjunction-property-publisher.md)** properties to return or set the data source query criterion. The following example changes an existing filter to remove from the mail merge all records that do not have a Region field equal to "WA". This example assumes that a data source is attached to the active publication.
 

 

```
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
 Next 
 End With 
End Sub
```

Use the  **[Add](mailmergefilters-add-method-publisher.md)** method of the **MailMergeFilters** object to add a new filter criterion to the query. This example adds a new line to the query string and then applies the combined filter to the data source. This example assumes that a data source is attached to the active publication.
 

 



```
Sub FilterDataSource() 
 With ActiveDocument.MailMerge.DataSource 
 .Filters.Add Column:="Region", _ 
 Comparison:=msoFilterComparisonIsBlank, _ 
 Conjunction:=msoFilterConjunctionAnd 
 .ApplyFilter 
 End With 
End Sub
```


## Properties



|**Name**|
|:-----|
|[Application](mailmergefiltercriterion-application-property-publisher.md)|
|[Column](mailmergefiltercriterion-column-property-publisher.md)|
|[CompareTo](mailmergefiltercriterion-compareto-property-publisher.md)|
|[Comparison](mailmergefiltercriterion-comparison-property-publisher.md)|
|[Conjunction](mailmergefiltercriterion-conjunction-property-publisher.md)|
|[Creator](mailmergefiltercriterion-creator-property-publisher.md)|
|[Index](mailmergefiltercriterion-index-property-publisher.md)|
|[Parent](mailmergefiltercriterion-parent-property-publisher.md)|

