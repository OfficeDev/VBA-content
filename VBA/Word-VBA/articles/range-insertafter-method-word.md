---
title: Range.InsertAfter Method (Word)
keywords: vbawd10.chm157155432
f1_keywords:
- vbawd10.chm157155432
ms.prod: word
api_name:
- Word.Range.InsertAfter
ms.assetid: 25b2c0be-e9c7-1e42-09ea-308bbdcde7c6
ms.date: 06/08/2017
---


# Range.InsertAfter Method (Word)

Inserts the specified text at the end of a range.


## Syntax

 _expression_ . **InsertAfter**( **_Text_** )

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Text_|Required| **String**|The text to be inserted.|

## Remarks

After this method is applied, the range expands to include the new text.

You can insert characters such as quotation marks, tab characters, and nonbreaking hyphens by using the Visual Basic  **Chr** function with the **InsertAfter** method. You can also use the following Visual Basic constants: **vbCr** , **vbLf** , **vbCrLf** and **vbTab** .

If you use this method with a range that refers to an entire paragraph, the text is inserted after the ending paragraph mark (the text will appear at the beginning of the next paragraph). To insert text at the end of a paragraph, determine the ending point and subtract 1 from this location (the paragraph mark is one character), as shown in the following example.




```vb
Set doc = ActiveDocument 
Set rngRange = _ 
 doc.Range(doc.Paragraphs(1).Start, _ 
 doc.Paragraphs(1).End - 1) 
rngRange.InsertAfter _ 
 " This is now the last sentence in paragraph one."
```

However, if the range ends with a paragraph mark that also happens to be the end of the document, Microsoft Word inserts the text before the final paragraph mark rather than creating a new paragraph at the end of the document.

Also, if the range is a bookmark, Word inserts the specified text but does not extend the range or the bookmark to include the new text.


## Example

This example inserts text at the end of the active document. The  **Content** property returns a **Range** object.


```vb
ActiveDocument.Content.InsertAfter "end of document"
```

This example inserts text from an input box as the second paragraph in the active document.




```vb
response = InputBox("Type some text") 
With ActiveDocument.Paragraphs(1).Range 
 .InsertAfter "1." &; Chr(9) &; response 
 .InsertParagraphAfter 
End With
```


## See also


#### Concepts


[Range Object](range-object-word.md)

