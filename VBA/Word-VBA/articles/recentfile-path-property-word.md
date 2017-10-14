---
title: RecentFile.Path Property (Word)
keywords: vbawd10.chm157548547
f1_keywords:
- vbawd10.chm157548547
ms.prod: word
api_name:
- Word.RecentFile.Path
ms.assetid: 5fa6c504-0168-ea5b-8455-bb617a3ee236
ms.date: 06/08/2017
---


# RecentFile.Path Property (Word)

Returns the disk or Web path to the specified object. Read-only  **String** .


## Syntax

 _expression_ . **Path**

 _expression_ Required. A variable that represents a **[RecentFile](recentfile-object-word.md)** object.


## Remarks

The path doesn't include a trailing character — for example, "C:\MSOffice" or "http://MyServer". Use the  **[PathSeparator](application-pathseparator-property-word.md)** property to add the character that separates folders and drive letters. Use the **[Name](recentfile-name-property-word.md)** property to return the file name without the path.


## See also


#### Concepts


[RecentFile Object](recentfile-object-word.md)

