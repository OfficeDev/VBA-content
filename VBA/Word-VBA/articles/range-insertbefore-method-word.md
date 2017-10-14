---
title: Range.InsertBefore Method (Word)
keywords: vbawd10.chm157155430
f1_keywords:
- vbawd10.chm157155430
ms.prod: word
api_name:
- Word.Range.InsertBefore
ms.assetid: ac77dcf7-ffcd-b109-8e17-ea6db169e85a
ms.date: 06/08/2017
---


# Range.InsertBefore Method (Word)

Inserts the specified text before the specified range.


## Syntax

 _expression_ . **InsertBefore**( **_Text_** )

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Text_|Required| **String**|The text to be inserted.|

## Remarks

After the text is inserted, the range is expanded to include the new text. If the range is a bookmark, the bookmark is also expanded to include the next text.

You can insert characters such as quotation marks, tab characters, and nonbreaking hyphens by using the Visual Basic  **Chr** function with the **InsertBefore** method. You can also use the following Visual Basic constants: **vbCr** , **vbLf** , **vbCrLf** and **vbTab** .


## Example

This example inserts the text "Introduction" as a separate paragraph at the beginning of the active document.


```vb
With ActiveDocument.Content 
 .InsertParagraphBefore 
 .InsertBefore "Introduction" 
End With
```


## See also


#### Concepts


[Range Object](range-object-word.md)

