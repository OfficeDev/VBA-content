---
title: PageSetup.VerticalAlignment Property (Word)
keywords: vbawd10.chm158400622
f1_keywords:
- vbawd10.chm158400622
ms.prod: word
api_name:
- Word.PageSetup.VerticalAlignment
ms.assetid: d21c70de-2f3a-3a33-df3d-e1b0a89314f9
ms.date: 06/08/2017
---


# PageSetup.VerticalAlignment Property (Word)

Returns or sets the vertical alignment of text on each page in a document or section. Read/write  **[WdVerticalAlignment](wdverticalalignment-enumeration-word.md)** .


## Syntax

 _expression_ . **VerticalAlignment**

 _expression_ Required. A variable that represents a **[PageSetup](pagesetup-object-word.md)** object.


## Example

This example changes the vertical alignment of the first document so that the text is centered between the top and bottom margins.


```
Documents(1).PageSetup.VerticalAlignment = wdAlignVerticalCenter
```

This example creates a new document and then inserts the same paragraph 10 times. The vertical alignment of the new document is then set so that the 10 paragraphs are equally spaced (justified) between the top and bottom margins.




```vb
Set myDoc = Documents.Add 
With myDoc.Content 
 For i = 1 to 9 
 .InsertAfter "This is a sentence." 
 .InsertParagraphAfter 
 Next i 
 .InsertAfter "This is a sentence." 
End With 
myDoc.PageSetup.VerticalAlignment = wdAlignVerticalJustify
```


## See also


#### Concepts


[PageSetup Object](pagesetup-object-word.md)

