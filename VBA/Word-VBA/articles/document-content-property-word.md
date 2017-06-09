---
title: Document.Content Property (Word)
keywords: vbawd10.chm158007337
f1_keywords:
- vbawd10.chm158007337
ms.prod: word
api_name:
- Word.Document.Content
ms.assetid: 80578329-a648-1d4b-f83d-4b2d289813fb
ms.date: 06/08/2017
---


# Document.Content Property (Word)

Returns a  **[Range](range-object-word.md)** object that represents the main document story. Read-only.


## Syntax

 _expression_ . **Content**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

The following two statements are equivalent:


```vb
Set mainStory = ActiveDocument.Content 
Set mainStory = ActiveDocument.StoryRanges(wdMainTextStory)
```


## Example

This example changes the font and font size of the text in the active document to Arial 10 point.


```vb
Set myRange = ActiveDocument.Content 
With myRange.Font 
 .Name = "Arial" 
 .Size = 10 
End With
```

This example inserts text at the end of the document named "Changes.doc." The  **For Each...Next** statement is used to determine whether the document is open.




```vb
For Each aDocument In Documents 
 If InStr(LCase$(aDocument.Name), "changes.doc") Then 
 Set myRange = Documents("Changes.doc").Content 
 myRange.InsertAfter "the end." 
 End If 
Next aDocument
```


## See also


#### Concepts


[Document Object](document-object-word.md)

