---
title: ListFormat.ConvertNumbersToText Method (Word)
keywords: vbawd10.chm163578042
f1_keywords:
- vbawd10.chm163578042
ms.prod: word
api_name:
- Word.ListFormat.ConvertNumbersToText
ms.assetid: 5ba6d823-dadb-1059-d439-0e556d91058f
ms.date: 06/08/2017
---


# ListFormat.ConvertNumbersToText Method (Word)

Changes the list numbers and LISTNUM fields in the specified  **ListFormat** object to text.


## Syntax

 _expression_ . **ConvertNumbersToText**

 _expression_ A variable that represents a **[ListFormat](listformat-object-word.md)** object.


## Example

This example converts the preset numbers in  _myRange_ to text without affecting any LISTNUM fields.


```vb
Set myDoc = ActiveDocumentSet myRange = _ 
    myDoc.Range(Start:=myDoc.Paragraphs(12).Range.Start, _ 
    End:=myDoc.Paragraphs(20).Range.End) 
myRange.ListFormat.ConvertNumbersToText wdNumberParagraph
```


## See also


#### Concepts


[ListFormat Object](listformat-object-word.md)

