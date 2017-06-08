---
title: Range.InsertFile Method (Word)
keywords: vbawd10.chm157155451
f1_keywords:
- vbawd10.chm157155451
ms.prod: word
api_name:
- Word.Range.InsertFile
ms.assetid: 9f35bacd-1cf3-42a4-c8ab-8c1cf183d2ab
ms.date: 06/08/2017
---


# Range.InsertFile Method (Word)

Inserts all or part of the specified file.


## Syntax

 _expression_ . **InsertFile**( **_FileName_** , **_Range_** , **_ConfirmConversions_** , **_Link_** , **_Attachment_** )

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FileName_|Required| **String**|The path and file name of the file to be inserted. If you don't specify a path, Word assumes the file is in the current folder.|
| _Range_|Optional| **Variant**|If the specified file is a Word document, this parameter refers to a bookmark. If the file is another type (for example, a Microsoft Excel worksheet), this parameter refers to a named range or a cell range (for example, R1C1:R3C4).|
| _ConfirmConversions_|Optional| **Variant**| **True** to have Word prompt you to confirm conversion when inserting files in formats other than the Word Document format.|
| _Link_|Optional| **Variant**| **True** to insert the file by using an INCLUDETEXT field.|
| _Attachment_|Optional| **Variant**| **True** to insert the file as an attachment to an e-mail message.|

## Example

This example uses an INCLUDETEXT field to insert the TEST.DOC file at the end of the current document.


```vb
ActiveDocument.Range.Collapse Direction:=wdCollapseEnd 
ActiveDocument.Range.InsertFile FileName:="C:\TEST.DOC", Link:=True
```


## See also


#### Concepts


[Range Object](range-object-word.md)

