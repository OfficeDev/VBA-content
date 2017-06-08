---
title: MailMergeFields.Add Method (Word)
keywords: vbawd10.chm153026661
f1_keywords:
- vbawd10.chm153026661
ms.prod: word
api_name:
- Word.MailMergeFields.Add
ms.assetid: a90cca41-15d7-92e0-2f60-9268d1579271
ms.date: 06/08/2017
---


# MailMergeFields.Add Method (Word)

Returns a  **MailMergeField** object that represents a mail merge field added to the data source document.


## Syntax

 _expression_ . **Add**( **_Range_** , **_Name_** )

 _expression_ Required. A variable that represents a **[MailMergeFields](mailmergefields-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Range_|Required| **Range**|The range where you want the field to appear. This field replaces the range, if the range isn't collapsed.|
| _Name_|Required| **String**|The name of the field.|

### Return Value

MailMergeField


## Example

This example replaces the selection with a mail merge field named MiddleInitial.


```vb
ActiveDocument.MailMerge.Fields.Add Range:=Selection.Range, _ 
 Name:="MiddleInitial"
```


## See also


#### Concepts


[MailMergeFields Collection Object](mailmergefields-object-word.md)

