---
title: Range.Bold Property (Word)
keywords: vbawd10.chm157155458
f1_keywords:
- vbawd10.chm157155458
ms.prod: word
api_name:
- Word.Range.Bold
ms.assetid: 04723b36-43bb-4721-90a5-33447a9b742e
ms.date: 06/08/2017
---


# Range.Bold Property (Word)

 **True** if the range is formatted as bold. Read/write **Long** .


## Syntax

 _expression_ . **Bold**

 _expression_ A variable that represents a **[Range](range-object-word.md)** object.


## Remarks

Returns  **True** , **False** or **wdUndefined** (a mixture of **True** and **False** ). Can be set to **True** , **False** , or **wdToggle** .


## Example

This example toggles the bold format for the selected text.


```vb
If Selection.Type = wdSelectionNormal Then 
 Selection.Range.Bold = wdToggle 
End If
```

This example makes the first paragraph in the active document bold.




```vb
ActiveDocument.Paragraphs(1).Range.Bold = True
```


## See also


#### Concepts


[Range Object](range-object-word.md)

