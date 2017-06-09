---
title: Font.Name Property (Word)
keywords: vbawd10.chm156369038
f1_keywords:
- vbawd10.chm156369038
ms.prod: word
api_name:
- Word.Font.Name
ms.assetid: 53928c78-c3f8-1b61-4cf4-fbe3cdc074c2
ms.date: 06/08/2017
---


# Font.Name Property (Word)

Returns or sets the name of the specified object. Read/write  **String** .


## Syntax

 _expression_ . **Name**

 _expression_ Required. A variable that represents a **[Font](font-object-word.md)** object.


## Example

This example formats the selection as Arial bold.


```vb
With Selection.Font 
 .Name = "Arial" 
 .Bold = True 
End With
```


## See also


#### Concepts


[Font Object](font-object-word.md)

