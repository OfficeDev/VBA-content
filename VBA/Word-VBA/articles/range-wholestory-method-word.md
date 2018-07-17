---
title: Range.WholeStory Method (Word)
keywords: vbawd10.chm157155456
f1_keywords:
- vbawd10.chm157155456
ms.prod: word
api_name:
- Word.Range.WholeStory
ms.assetid: bb55c363-b3c0-e1aa-5e25-74cf2a1954c8
ms.date: 06/08/2017
---


# Range.WholeStory Method (Word)

Expands a range to include the entire story.


## Syntax

 _expression_ . **WholeStory**

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


## Remarks

The following instructions, where  _myRange_ is a valid **Range** object, are functionally equivalent:


```
myRange.WholeStory 
myRange.Expand Unit:=wdStory
```


## Example

This example expands  _myRange_ to include the entire story and then applies the Arial font to the range.


```vb
Set myRange = Selection.Range 
myRange.WholeStory 
myRange.Font.Name = "Arial"
```

This example expands  _myRange_ to include the entire comments story ( **wdCommentsStory** ) and then copies the comments into a new document.




```vb
If ActiveDocument.Comments.Count >= 1 Then 
 Set myRange = Activedocument.Comments(1).Range 
 myRange.WholeStory 
 myRange.Copy 
 Documents.Add.Content.Paste 
End If
```


## See also


#### Concepts


[Range Object](range-object-word.md)

