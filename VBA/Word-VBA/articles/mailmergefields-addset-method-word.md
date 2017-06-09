---
title: MailMergeFields.AddSet Method (Word)
keywords: vbawd10.chm153026669
f1_keywords:
- vbawd10.chm153026669
ms.prod: word
api_name:
- Word.MailMergeFields.AddSet
ms.assetid: 6b35e6ab-e918-26bd-6cdd-547882d2bef5
ms.date: 06/08/2017
---


# MailMergeFields.AddSet Method (Word)

Adds a SET field to a mail merge main document. Returns a  **MailMergeField** object.


## Syntax

 _expression_ . **AddSet**( **_Range_** , **_Name_** , **_ValueText_** , **_ValueAutoText_** )

 _expression_ Required. A variable that represents a **[MailMergeFields](mailmergefields-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Range_|Required| **Range object**|The location for the SET field.|
| _Name_|Required| **String**|The bookmark name that ValueText is assigned to.|
| _ValueText_|Optional| **Variant**|The text associated with the bookmark specified by the Name argument.|
| _ValueAutoText_|Optional| **Variant**|The AutoText entry that includes text associated with the bookmark specified by the Name argument. If this argument is specified, ValueText is ignored.|

### Return Value

MailMergeField


## Remarks

A SET field defines the text of the specified bookmark.


## Example

This example adds a SET field at the beginning of the active document and then adds a REF field to display the text after the selection.


```vb
Dim rngTemp as Range 
 
Set rngTemp = ActiveDocument.Range(Start:=0, End:=0) 
 
ActiveDocument.MailMerge.Fields.AddSet Range:=rngTemp, _ 
 Name:="Name", ValueText:="Joe Smith" 
Selection.Collapse Direction:=wdCollapseEnd 
ActiveDocument.Fields.Add Range:=Selection.Range, _ 
 Type:=wdFieldRef, Text:="Name"
```


## See also


#### Concepts


[MailMergeFields Collection Object](mailmergefields-object-word.md)

