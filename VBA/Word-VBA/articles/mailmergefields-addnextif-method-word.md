---
title: MailMergeFields.AddNextIf Method (Word)
keywords: vbawd10.chm153026668
f1_keywords:
- vbawd10.chm153026668
ms.prod: word
api_name:
- Word.MailMergeFields.AddNextIf
ms.assetid: ac89e9c2-48b5-243b-65f4-4904fb18d043
ms.date: 06/08/2017
---


# MailMergeFields.AddNextIf Method (Word)

Adds a NEXTIF field to a mail merge main document. Returns a  **MailMergeField** object.


## Syntax

 _expression_ . **AddNextIf**( **_Range_** , **_MergeField_** , **_Comparison_** , **_CompareTo_** )

 _expression_ Required. A variable that represents a **[MailMergeFields](mailmergefields-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Range_|Required| **Range object**|The location for the NEXTIF field.|
| _MergeField_|Required| **String**|The merge field name.|
| _Comparison_|Required| **WdMailMergeComparison**|The operator used in the comparison.|
| _CompareTo_|Required| **String**|The text to compare with the contents of MergeField.|

### Return Value

MailMergeField


## Remarks

A NEXTIF field compares two expressions, and if the comparison is true, the next record is merged into the current merge document.


## Example

This example adds a NEXTIF field before the first MERGEFIELD field in Main.doc. If the next postal code equals 98004, the next record is merged into the current merge document.


```
Documents("Main.doc").MailMerge.Fields(1).Select 
Selection.Collapse Direction:=wdCollapseStart 
Documents("Main.doc").MailMerge.Fields.AddNextIf _ 
 Range:=Selection.Range, MergeField:="PostalCode", _ 
 Comparison:=wdMergeIfEqual, CompareTo:="98004"
```


## See also


#### Concepts


[MailMergeFields Collection Object](mailmergefields-object-word.md)

