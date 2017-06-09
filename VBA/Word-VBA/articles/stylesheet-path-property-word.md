---
title: StyleSheet.Path Property (Word)
keywords: vbawd10.chm166658052
f1_keywords:
- vbawd10.chm166658052
ms.prod: word
api_name:
- Word.StyleSheet.Path
ms.assetid: 96a68487-b1b8-4c45-1869-b066874df9e5
ms.date: 06/08/2017
---


# StyleSheet.Path Property (Word)

Returns the disk or Web path to the specified style sheet. Read-only  **String** .


## Syntax

 _expression_ . **Path**

 _expression_ Required. A variable that represents a **[StyleSheet](stylesheet-object-word.md)** object.


## Remarks

The path doesn't include a trailing character—for example, "C:\MSOffice" or "http://MyServer". Use the  **[PathSeparator](application-pathseparator-property-word.md)** property to add the character that separates folders and drive letters, and use the **[Name](stylesheet-name-property-word.md)** property to return the file name without the path.


 **Note**  You can use the  **PathSeparator** property to build Web addresses even though they contain forward slashes (/) and the **PathSeparator** property defaults to a backslash (\).


## See also


#### Concepts


[StyleSheet Object](stylesheet-object-word.md)

