---
title: Document.FormattingShowClear Property (Word)
keywords: vbawd10.chm158007745
f1_keywords:
- vbawd10.chm158007745
ms.prod: word
api_name:
- Word.Document.FormattingShowClear
ms.assetid: e6a25cc8-29be-0ba4-21ba-763676cc2f90
ms.date: 06/08/2017
---


# Document.FormattingShowClear Property (Word)

 **True** for Microsoft Word to show clear formatting in the **Styles and Formatting** task pane. Read/write **Boolean** .


## Syntax

 _expression_ . **FormattingShowClear**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Example

This example disables display of the  **Clear Formatting** button in the list of styles.


```vb
Sub ShowClearFormatting() 
 With ActiveDocument 
 .FormattingShowClear = False 
 .FormattingShowFilter = wdShowFilterFormattingInUse 
 .FormattingShowFont = True 
 .FormattingShowNumbering = True 
 .FormattingShowParagraph = True 
 End With 
End Sub
```


## See also


#### Concepts


[Document Object](document-object-word.md)

