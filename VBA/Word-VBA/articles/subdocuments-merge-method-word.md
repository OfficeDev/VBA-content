---
title: Subdocuments.Merge Method (Word)
keywords: vbawd10.chm159907942
f1_keywords:
- vbawd10.chm159907942
ms.prod: word
api_name:
- Word.Subdocuments.Merge
ms.assetid: 486b0b4e-1bc7-4ba3-15f0-466aede8c172
ms.date: 06/08/2017
---


# Subdocuments.Merge Method (Word)

Merges the specified subdocuments of a master document into a single subdocument.


## Syntax

 _expression_ . **Merge**( **_FirstSubdocument_** , **_LastSubdocument_** )

 _expression_ Required. A variable that represents a **[Subdocuments](subdocuments-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FirstSubdocument_|Optional| **Variant**|The path and file name of the original document you want to merge revisions with.|
| _LastSubdocument_|Optional| **Variant**|The last subdocument in a range of subdocuments to be merged.|

## Example

This example merges the first and second subdocuments in the active document into one subdocument.


```vb
If ActiveDocument.Subdocuments.Count >= 2 Then 
 Set aDoc = ActiveDocument 
 aDoc.Subdocuments.Merge _ 
 FirstSubdocument:=aDoc.Subdocuments(1), _ 
 LastSubdocument:=aDoc.Subdocuments(2) 
End If
```


## See also


#### Concepts


[Subdocuments Collection Object](subdocuments-object-word.md)

