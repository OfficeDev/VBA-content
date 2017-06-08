---
title: TableStyle.TableDirection Property (Word)
keywords: vbawd10.chm244776972
f1_keywords:
- vbawd10.chm244776972
ms.prod: word
api_name:
- Word.TableStyle.TableDirection
ms.assetid: 3569f6a0-6339-b9ae-3e0d-dc1f1cadb777
ms.date: 06/08/2017
---


# TableStyle.TableDirection Property (Word)

Returns or sets the direction in which Microsoft Word orders cells in the specified table style. Read/write  **[WdTableDirection](wdtabledirection-enumeration-word.md)** .


## Syntax

 _expression_ . **TableDirection**

 _expression_ Required. A variable that represents a **[TableStyle](tablestyle-object-word.md)** object.


## Remarks

If the  **TableDirection** property is set to **wdTableDirectionLtr** , the selected rows are arranged with the first column in the leftmost position. If the **TableDirection** property is set to **wdTableDirectionRtl** , the selected rows are arranged with the first column in the rightmost position.


## See also


#### Concepts


[TableStyle Object](tablestyle-object-word.md)

