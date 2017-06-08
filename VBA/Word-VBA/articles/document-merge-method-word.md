---
title: Document.Merge Method (Word)
keywords: vbawd10.chm158007658
f1_keywords:
- vbawd10.chm158007658
ms.prod: word
api_name:
- Word.Document.Merge
ms.assetid: e7ab537d-dfd3-177b-722a-6fe693c158d8
ms.date: 06/08/2017
---


# Document.Merge Method (Word)

Merges the changes marked with revision marks from one document to another.


## Syntax

 _expression_ . **Merge**( **_Name_** , **_MergeTarget_** , **_DetectFormatChanges_** , **_UseFormattingFrom_** , **_AddToRecentFiles_** )

 _expression_ Required. A variable that represents a **[Document](document-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The path and file name of the document with which to merge.|
| _MergeTarget_|Optional| **WdMergeTarget**|Specifies where to place the final merged content.|
| _DetectFormatChanges_|Optional| **Boolean**|Specifies whether or not to mark formatting differences.|
| _UseFormattingFrom_|Optional| **WdUseFormattingFrom**|Specifies which document to use for formatting in the merged document.|
| _AddToRecentFiles_|Optional| **Boolean**|Specifies whether to add the document in the Name parameter to the list of recent files.|

## Example

This example merges changes from Sales1.doc into Sales2.doc (the active document).


```vb
If InStr(1, ActiveDocument.Name, "sales2.doc", 1) Then _ 
 ActiveDocument.Merge Name:="C:\Docs\Sales1.doc"
```


## See also


#### Concepts


[Document Object](document-object-word.md)

