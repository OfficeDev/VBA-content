---
title: Document.TrackMoves Property (Word)
keywords: vbawd10.chm158007795
f1_keywords:
- vbawd10.chm158007795
ms.prod: word
api_name:
- Word.Document.TrackMoves
ms.assetid: 6c94cd58-dd47-313c-c04f-f04fe6f86f02
ms.date: 06/08/2017
---


# Document.TrackMoves Property (Word)

Returns or sets a ** Boolean** that represents whether to mark moved text when Track Changes is turned on. Read/write.


## Syntax

 _expression_ . **TrackMoves**

 _expression_ An expression that returns a **Document** object.


## Remarks

By default, when Track Changes is turned on, moved text is marked as deleted and inserted. When  **TrackMoves** is **True** , moved text is marked as moved, with the from text being marked with strikethrough formatting. This property corresponds to the **Track Moves** check box in the **Track Change Options** dialog box.


## See also


#### Concepts


[Document Object](document-object-word.md)

