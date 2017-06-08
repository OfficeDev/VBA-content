---
title: MailMergeFields.AddSkipIf Method (Word)
keywords: vbawd10.chm153026670
f1_keywords:
- vbawd10.chm153026670
ms.prod: word
api_name:
- Word.MailMergeFields.AddSkipIf
ms.assetid: feaa8b59-292c-0e6f-661a-af501b395cf9
ms.date: 06/08/2017
---


# MailMergeFields.AddSkipIf Method (Word)

Adds a SKIPIF field to a mail merge main document. Returns a  **MailMergeField** object. .


## Syntax

 _expression_ . **AddSkipIf**( **_Range_** , **_MergeField_** , **_Comparison_** , **_CompareTo_** )

 _expression_ Required. A variable that represents a **[MailMergeFields](mailmergefields-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Range_|Required| **Range object**|The location for the SKIPIF field.|
| _MergeField_|Required| **String**|The merge field name.|
| _Comparison_|Required| **WdMailMergeComparison**|The operator used in the comparison.|
| _CompareTo_|Optional| **Variant**|The text to compare with the contents of MergeField.|

### Return Value

MailMergeField


## Remarks

A SKIPIF field compares two expressions, and if the comparison is true, SKIPIF moves to the next record in the data source and starts a new merge document.


## Example

This example adds a SKIPIF field before the first MERGEFIELD field in Main.doc. If the next postal code equals 98040, the next record is skipped.


```
Documents("Main.doc").MailMerge.Fields(1).Select 
Selection.Collapse Direction:=wdCollapseStart 
Documents("Main.doc").MailMerge.Fields.AddSkipIf _ 
 Range:=Selection.Range, MergeField:="PostalCode", _ 
 Comparison:=wdMergeIfEqual, CompareTo:="98040"
```


## See also


#### Concepts


[MailMergeFields Collection Object](mailmergefields-object-word.md)

