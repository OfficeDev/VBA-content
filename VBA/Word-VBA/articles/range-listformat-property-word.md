---
title: Range.ListFormat Property (Word)
keywords: vbawd10.chm157155396
f1_keywords:
- vbawd10.chm157155396
ms.prod: word
api_name:
- Word.Range.ListFormat
ms.assetid: 509365dc-0b93-96d9-6614-74f2d85bfd45
ms.date: 06/08/2017
---


# Range.ListFormat Property (Word)

Returns a  **[ListFormat](listformat-object-word.md)** object that represents all the list formatting characteristics of a range. Read-only.


## Syntax

 _expression_ . **ListFormat**

 _expression_ A variable that represents a **[Range](range-object-word.md)** object.


## Example

This example sets the variable myDoc to a range that includes paragraphs three through six of the active document. The example then either applies the default outline-numbered list format to the range or removes it, depending on whether or not the format was already applied to the range.


```vb
Set myDoc = ActiveDocument 
Set myRange = _ 
 myDoc.Range(Start:= myDoc.Paragraphs(3).Range.Start, _ 
 End:=myDoc.Paragraphs(6).Range.End) 
myRange.ListFormat.ApplyOutlineNumberDefault
```

This example applies the second list template on the  **Numbered** tab in the **Bullets and Numbering** dialog box to all the paragraphs in the selection.




```
Selection.Range.ListFormat.ApplyListTemplate _ 
 ListTemplate:=ListGalleries(wdNumberGallery).ListTemplates(2)
```


## See also


#### Concepts


[Range Object](range-object-word.md)

