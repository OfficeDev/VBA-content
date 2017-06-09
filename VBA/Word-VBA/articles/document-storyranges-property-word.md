---
title: Document.StoryRanges Property (Word)
keywords: vbawd10.chm158007352
f1_keywords:
- vbawd10.chm158007352
ms.prod: word
api_name:
- Word.Document.StoryRanges
ms.assetid: 6afc9e1a-950c-e1b0-15d5-73afeb72fc59
ms.date: 06/08/2017
---


# Document.StoryRanges Property (Word)

Returns a  **[StoryRanges](storyranges-object-word.md)** collection that represents all the stories in the specified document. Read-only.


## Syntax

 _expression_ . **StoryRanges**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example steps through the  **StoryRanges** collection to determine whether **wdPrimaryFooterStory** is part of the **StoryRanges** collection.


```vb
For Each aStory In ActiveDocument.StoryRanges 
 If aStory.StoryType = wdEvenPagesFooterStory Then 
 MsgBox "Document includes an even page footer" 
 End If 
Next aStory
```

This example adds text to the primary header story and then displays the text.




```vb
ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary).Range _ 
 .Text = "Header text" 
MsgBox ActiveDocument.StoryRanges(wdPrimaryHeaderStory).Text
```


## See also


#### Concepts


[Document Object](document-object-word.md)

