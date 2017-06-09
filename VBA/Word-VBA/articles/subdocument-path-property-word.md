---
title: Subdocument.Path Property (Word)
keywords: vbawd10.chm159973380
f1_keywords:
- vbawd10.chm159973380
ms.prod: word
api_name:
- Word.Subdocument.Path
ms.assetid: d27bc7ce-5346-b9a7-cd29-b42b0e8c98eb
ms.date: 06/08/2017
---


# Subdocument.Path Property (Word)

Returns the disk or Web path to the specified subdocument. Read-only  **String** .


## Syntax

 _expression_ . **Path**

 _expression_ Required. A variable that represents a **[Subdocument](subdocument-object-word.md)** object.


## Remarks

The path doesn't include a trailing character — for example, "C:\MSOffice" or "http://MyServer". Use the  **[PathSeparator](application-pathseparator-property-word.md)** property to add the character that separates folders and drive letters. Use the **[Name](subdocument-name-property-word.md)** property to return the file name without the path.


 **Note**  You can use the  **PathSeparator** property to build Web addresses even though they contain forward slashes (/) and the **PathSeparator** property defaults to a backslash (\).


## See also


#### Concepts


[Subdocument Object](subdocument-object-word.md)

