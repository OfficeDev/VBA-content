---
title: Selection.FormattedText Property (Word)
keywords: vbawd10.chm158662658
f1_keywords:
- vbawd10.chm158662658
ms.prod: word
api_name:
- Word.Selection.FormattedText
ms.assetid: b16da3f9-1aa6-e722-0a9c-8a4c30922450
ms.date: 06/08/2017
---


# Selection.FormattedText Property (Word)

Returns or sets a  **[Range](range-object-word.md)** object that includes the formatted text in the specified range or selection. Read/write.


## Syntax

 _expression_ . **FormattedText**

 _expression_ A variable that represents a **[Selection](selection-object-word.md)** object.


## Remarks

This property returns a  **[Range](range-object-word.md)** object with the character formatting and text from the specified range or selection. Paragraph formatting is included in the **[Range](range-object-word.md)** object if there is a paragraph mark in the range or selection.



When you set this property, the text in the range is replaced with formatted text. If you don't want to replace the existing text, use the  **Collapse** method before using this property (see the first example).




## Example

This example copies the first paragraph in the document, including its formatting, and inserts the formatted text at the insertion point.


```
Selection.Collapse Direction:=wdCollapseStart 
Selection.FormattedText = ActiveDocument.Paragraphs(1).Range
```

This example copies the text and formatting from the selection into a new document.




```vb
Set myRange = Selection.FormattedText 
Documents.Add.Content.FormattedText = myRange
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

