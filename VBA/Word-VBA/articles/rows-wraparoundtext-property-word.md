---
title: Rows.WrapAroundText Property (Word)
keywords: vbawd10.chm155975692
f1_keywords:
- vbawd10.chm155975692
ms.prod: word
api_name:
- Word.Rows.WrapAroundText
ms.assetid: 6d899cb2-f8af-1b20-3d8e-4ef353d4b762
ms.date: 06/08/2017
---


# Rows.WrapAroundText Property (Word)

Returns or sets whether text should wrap around the specified rows. Read/write  **Long** .


## Syntax

 _expression_ . **WrapAroundText**

 _expression_ An expression that returns a **Rows** object.


## Remarks

Returns  **wdUndefined** if only some of the specified rows have wrapping enabled. Can be set to **True** or **False** . Setting the **WrapAroundText** property to **False** also sets the **[AllowOverlap](rows-allowoverlap-property-word.md)** property to **False** . Setting the **AllowOverlap** property to **True** also sets the **WrapAroundText** property to **True** .


## Example

This example sets Microsoft Word to wrap text around the first table in the document.


```vb
ActiveDocument.Tables(1).Rows.WrapAroundText = True
```


## See also


#### Concepts


[Rows Collection Object](rows-object-word.md)

