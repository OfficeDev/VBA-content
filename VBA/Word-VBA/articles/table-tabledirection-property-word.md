---
title: Table.TableDirection Property (Word)
keywords: vbawd10.chm156303478
f1_keywords:
- vbawd10.chm156303478
ms.prod: word
api_name:
- Word.Table.TableDirection
ms.assetid: 3062731b-a334-927d-3871-f845cfb662ac
ms.date: 06/08/2017
---


# Table.TableDirection Property (Word)

Returns or sets the direction in which Microsoft Word orders cells in the specified table. Read/write  **[WdTableDirection](wdtabledirection-enumeration-word.md)** .


## Syntax

 _expression_ . **TableDirection**

 _expression_ Required. A variable that represents a **[Table](table-object-word.md)** object.


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


[Table Object](table-object-word.md)

