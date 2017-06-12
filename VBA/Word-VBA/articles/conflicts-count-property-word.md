---
title: Conflicts.Count Property (Word)
keywords: vbawd10.chm174391301
f1_keywords:
- vbawd10.chm174391301
ms.prod: word
api_name:
- Word.Conflicts.Count
ms.assetid: 7a9488a5-d29c-16af-cab0-cbc2fe7fba96
ms.date: 06/08/2017
---


# Conflicts.Count Property (Word)

Returns the number of items in the  **Conflicts** collection. Read-only.


## Syntax

 _expression_ . **Count**

 _expression_ An expression that returns a **Conflicts** object.


## Example

The following code example gets the number of  **Conflict** objects in the active document.


```vb
Dim confCount as Long 
 
confCount = ActiveDocument.CoAuthoring.Conflicts.Count 

```


## See also


#### Concepts


[Conflicts Object](conflicts-object-word.md)

