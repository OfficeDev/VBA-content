---
title: Selection.InsertAfter Method (Word)
keywords: vbawd10.chm158662760
f1_keywords:
- vbawd10.chm158662760
ms.prod: word
api_name:
- Word.Selection.InsertAfter
ms.assetid: 21286a89-5e4e-56ae-27a5-f581a337bfbb
ms.date: 06/08/2017
---


# Selection.InsertAfter Method (Word)

Inserts the specified text at the end of a range or selection.


## Syntax

 _expression_ . **InsertAfter**( **_Text_** )

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Text_|Required| **String**|The text to be inserted.|

## Remarks

After using this method, the selection expands to include the new text.

You can insert characters such as quotation marks, tab characters, and nonbreaking hyphens by using the Visual Basic  **Chr** function with the **InsertAfter** method. You can also use the following Visual Basic constants: **vbCr** , **vbLf** , **vbCrLf** and **vbTab** .

If you use this method with a selection that refers to an entire paragraph, the text is inserted after the ending paragraph mark (the text will appear at the beginning of the next paragraph). To insert text at the end of a paragraph, determine the ending point and subtract 1 from this location (the paragraph mark is one character), as shown in the following example.




```vb
ActiveDocument.Range( _ 
 ActiveDocument.Paragraphs(1).Range.Start, _ 
 ActiveDocument.Paragraphs(1).Range.End - 1) _ 
 .Select 
 
Selection.InsertAfter _ 
 " This is now the last sentence in paragraph one."
```

However, if the selection ends with a paragraph mark that also happens to be the end of the document, Microsoft Word inserts the text before the final paragraph mark rather than creating a new paragraph at the end of the document. Also, if the selection is a bookmark, Word inserts the specified text but does not extend the selection or the bookmark to include the new text.


## Example

This example inserts text at the end of the selection and then collapses the selection to an insertion point.


```vb
With Selection 
 .InsertAfter "appended text" 
 .Collapse Direction:=wdCollapseEnd 
End With
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

