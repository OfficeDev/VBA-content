---
title: Font.Shrink Method (Word)
keywords: vbawd10.chm156368997
f1_keywords:
- vbawd10.chm156368997
ms.prod: word
api_name:
- Word.Font.Shrink
ms.assetid: 6a4ca959-07df-2b17-f59e-c6cf1f6236c6
ms.date: 06/08/2017
---


# Font.Shrink Method (Word)

Decreases the font size to the next available size.


## Syntax

 _expression_ . **Shrink**

 _expression_ A variable that represents a **[Font](font-object-word.md)** object.


## Remarks

If the selection or range contains more than one font size, each size is decreased to the next available setting.


## Example

This example inserts a line of increasingly smaller Z's in a new document.


```vb
Set myRange = Documents.Add.Content 
myRange.Font.Size = 45 
For Count = 1 To 5 
 myRange.InsertAfter "Z" 
 For Count2 = 1 to 3 
 myRange.Characters(Count).Font.Shrink 
 Next Count2 
Next Count
```


## See also


#### Concepts


[Font Object](font-object-word.md)

