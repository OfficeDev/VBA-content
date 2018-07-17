---
title: Dictionary.Path Property (Word)
keywords: vbawd10.chm162332673
f1_keywords:
- vbawd10.chm162332673
ms.prod: word
api_name:
- Word.Dictionary.Path
ms.assetid: 1fd2d6ac-e112-9d13-0e41-2584e6841b73
ms.date: 06/08/2017
---


# Dictionary.Path Property (Word)

Returns the path to the specified dictionary. Read-only  **String** .


## Syntax

 _expression_ . **Path**

 _expression_ Required. A variable that represents a **[Dictionary](dictionary-object-word.md)** object.


## Remarks

The path doesn't include a trailing character — for example, "C:\MSOffice" or "http://MyServer". Use the  **PathSeparator** property to add the character that separates folders and drive letters. Use the **Name** property to return the file name without the path and use the **FullName** property to return the file name and the path together.


 **Note**  You can use the  **PathSeparator** property to build Web addresses even though they contain forward slashes (/) and the **PathSeparator** property defaults to a backslash (\).


## See also


#### Concepts


[Dictionary Object](dictionary-object-word.md)

