---
title: Range.StoryLength Property (Word)
keywords: vbawd10.chm157155480
f1_keywords:
- vbawd10.chm157155480
ms.prod: word
api_name:
- Word.Range.StoryLength
ms.assetid: 0dd342e2-2a90-bbf9-2989-a2629fcf40a5
ms.date: 06/08/2017
---


# Range.StoryLength Property (Word)

Returns the number of characters in the story that contains the specified range. Read-only  **Long** .


## Syntax

 _expression_ . **StoryLength**

 _expression_ A variable that represents a **[Range](range-object-word.md)** object.


## Example

This example determines whether the header in the active document is empty. If the header story is not empty, a message box displays the contents of the header. If the document header is empty,  **StoryLength** returns 1 for the final paragraph mark.


```vb
Set myRange = ActiveDocument.Sections(1) _ 
 .Headers(wdHeaderFooterPrimary).Range 
If myRange.StoryLength > 1 Then MsgBox myRange.Text
```

This example closes the document without saving changes, if it is empty.




```vb
If ActiveDocument.Content.StoryLength = 1 Then _ 
 ActiveDocument.Close SaveChanges:=wdDoNotSaveChanges
```


## See also


#### Concepts


[Range Object](range-object-word.md)

