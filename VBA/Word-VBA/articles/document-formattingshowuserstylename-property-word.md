---
title: Document.FormattingShowUserStyleName Property (Word)
keywords: vbawd10.chm158007819
f1_keywords:
- vbawd10.chm158007819
ms.prod: word
api_name:
- Word.Document.FormattingShowUserStyleName
ms.assetid: 16bdfdcd-f550-9b15-d405-20bd391aa0e5
ms.date: 06/08/2017
---


# Document.FormattingShowUserStyleName Property (Word)

Returns or sets a  **Boolean** that represents whether to show user-defined styles. Read/write.


## Syntax

 _expression_ . **FormattingShowUserStyleName**

 _expression_ An expression that returns a **Document** object.


## Remarks

This property corresponds to the  **Hide built-in name when alternate names exists** check box in the **Styles Gallery Options** dialog box. Setting the **FormattingShowUserStyleName** property to **True** hides built-in styles when alternate user-defined styles exist. For example, if a user creates a heading style, such as Heading 1 or Heading 2, to take advantage of other built-in features of Microsoft Word, such as tables of contents, the user-defined style takes precedence, and the similarly named built-in style is not shown in the list of style names.


## See also


#### Concepts


[Document Object](document-object-word.md)

