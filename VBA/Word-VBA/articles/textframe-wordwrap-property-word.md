---
title: TextFrame.WordWrap Property (Word)
keywords: vbawd10.chm162665362
f1_keywords:
- vbawd10.chm162665362
ms.prod: word
api_name:
- Word.TextFrame.WordWrap
ms.assetid: 70bef68b-3c37-9b4e-4cfe-ed0832a7934c
ms.date: 06/08/2017
---


# TextFrame.WordWrap Property (Word)

 **True** if Microsoft Word wraps Latin text in the middle of a word in the specified text frames. Read/write **Long** . .


## Syntax

 _expression_ . **WordWrap**

 _expression_ Required. A variable that represents a **[TextFrame](textframe-object-word.md)** object.


## Remarks

This property returns  **wdUndefined** if it's set to **True** for some of the specified text in the specified text frame and false for other text.


 **Note**  This property may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.


## See also


#### Concepts


[TextFrame Object](textframe-object-word.md)

