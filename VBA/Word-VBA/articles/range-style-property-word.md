---
title: Range.Style Property (Word)
keywords: vbawd10.chm157155479
f1_keywords:
- vbawd10.chm157155479
ms.prod: word
api_name:
- Word.Range.Style
ms.assetid: aeceef42-cbdc-3d55-2f43-0afffd933cc2
ms.date: 06/08/2017
---


# Range.Style Property (Word)

Returns or sets the style for the specified object. Read/write  **Variant** .


## Syntax

 _expression_ . **Style**

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


## Remarks

To set this property, specify the local name of the style, an integer, a  **[WdBuiltinStyle](wdbuiltinstyle-enumeration-word.md)** constant, or an object that represents the style. When you return the style for a range that includes more than one style, only the first character or paragraph style is returned.


## Example

This example displays the style for each character in the selection. 


 **Note**  Each element of the  **Characters** collection is a **Range** object.


```vb
For Each c in Selection.Characters 
 MsgBox c.Style 
Next c
```


## See also


#### Concepts


[Range Object](range-object-word.md)

