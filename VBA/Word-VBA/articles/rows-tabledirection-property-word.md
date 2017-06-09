---
title: Rows.TableDirection Property (Word)
keywords: vbawd10.chm155975784
f1_keywords:
- vbawd10.chm155975784
ms.prod: word
api_name:
- Word.Rows.TableDirection
ms.assetid: 02351774-13c0-ec82-c553-3b048eabb133
ms.date: 06/08/2017
---


# Rows.TableDirection Property (Word)

Returns or sets the direction in which Microsoft Word orders cells in the specified table or row. Read/write  **[WdTableDirection](wdtabledirection-enumeration-word.md)** .


## Syntax

 _expression_ . **TableDirection**

 _expression_ Required. A variable that represents a **[Rows](rows-object-word.md)** collection.


## Remarks

If the  **TableDirection** property is set to **wdTableDirectionLtr** , the selected rows are arranged with the first column in the leftmost position. If the **TableDirection** property is set to **wdTableDirectionRtl** , the selected rows are arranged with the first column in the rightmost position.


## Example

This example sets Microsoft Word to order cells in the selected row from right to left.


```
Selection.Rows.TableDirection = _ 
 wdTableDirectionRtl
```


## See also


#### Concepts


[Rows Collection Object](rows-object-word.md)

