---
title: Document.OMathBreakSub Property (Word)
keywords: vbawd10.chm158007825
f1_keywords:
- vbawd10.chm158007825
ms.prod: word
api_name:
- Word.Document.OMathBreakSub
ms.assetid: a361f255-1392-eddc-7771-98e9db7c291a
ms.date: 06/08/2017
---


# Document.OMathBreakSub Property (Word)

Returns or sets a  **[WdOMathBreakSub](wdomathbreaksub-enumeration-word.md)** constant that represents how Microsoft Word handles a subtraction operator that falls before a line break. Read/write.


## Syntax

 _expression_ . **OMathBreakSub**

 _expression_ An expression that returns a **Document** object.


## Remarks

This property is used only when the  **[OMathBreakBin](document-omathbreakbin-property-word.md)** property is set to **wdOMathBreakBinRepeat** . Subtraction sometimes receives special treatment when a line break falls on a subtraction operator and the document setting is to repeat the subtraction operator on the following line, because two negatives make a positive. Some writers choose to convert one of the minus signs into a plus sign, and some choose to keep the two negatives.


## See also


#### Concepts


[Document Object](document-object-word.md)

