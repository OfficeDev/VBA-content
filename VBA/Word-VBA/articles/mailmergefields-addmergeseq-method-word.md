---
title: MailMergeFields.AddMergeSeq Method (Word)
keywords: vbawd10.chm153026666
f1_keywords:
- vbawd10.chm153026666
ms.prod: word
api_name:
- Word.MailMergeFields.AddMergeSeq
ms.assetid: e437677d-2b2b-e921-d5e2-817a67624b66
ms.date: 06/08/2017
---


# MailMergeFields.AddMergeSeq Method (Word)

Adds a MERGESEQ field to a mail merge main document. Returns a  **MailMergeField** object.


## Syntax

 _expression_ . **AddMergeSeq**( **_Range_** )

 _expression_ Required. A variable that represents a **[MailMergeFields](mailmergefields-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Range_|Required| **Range object**|The location for the MERGESEQ field.|

### Return Value

MailMergeField


## Remarks

A MERGESEQ field inserts a number based on the sequence in which records are merged (for example, when record 50 of records 50 to 100 is merged, MERGESEQ inserts the number 1).


## Example

This example inserts text and a MERGESEQ field at the end of the active document.


```vb
Dim rngTemp As Range 
 
Set rngTemp = ActiveDocument.Content 
 
rngTemp.Collapse Direction:=wdCollapseEnd 
ActiveDocument.MailMerge.Fields.AddMergeSeq Range:=rngTemp 
rngTemp.InsertAfter "Sequence Number: "
```


## See also


#### Concepts


[MailMergeFields Collection Object](mailmergefields-object-word.md)

