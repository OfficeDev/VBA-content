---
title: RecentFiles.Maximum Property (Word)
keywords: vbawd10.chm157483010
f1_keywords:
- vbawd10.chm157483010
ms.prod: word
api_name:
- Word.RecentFiles.Maximum
ms.assetid: 3bdab716-106f-6e18-abe1-863450397ab9
ms.date: 06/08/2017
---


# RecentFiles.Maximum Property (Word)

Returns or sets the maximum number of recently used files that can appear on the  **File** menu. Can be a number from 0 (zero) through 9. Read/write **Long** .


## Syntax

 _expression_ . **Maximum**

 _expression_ Required. A variable that represents a **[RecentFiles](recentfiles-object-word.md)** collection.


## Example

This example disables the list of most recently used files.


```
RecentFiles.Maximum = 0
```

This example increases the number of items on the list of most recently used files by 1.




```vb
num = RecentFiles.Maximum 
If num <> 9 Then RecentFiles.Maximum = num + 1
```


## See also


#### Concepts


[RecentFiles Collection Object](recentfiles-object-word.md)

