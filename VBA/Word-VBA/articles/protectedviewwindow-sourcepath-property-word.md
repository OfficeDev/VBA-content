---
title: ProtectedViewWindow.SourcePath Property (Word)
keywords: vbawd10.chm231735307
f1_keywords:
- vbawd10.chm231735307
ms.prod: word
api_name:
- Word.ProtectedViewWindow.SourcePath
ms.assetid: 05b4e601-894a-de8f-1119-565183b244b7
ms.date: 06/08/2017
---


# ProtectedViewWindow.SourcePath Property (Word)

Returns the path of the source file for the specified protected view window. Read-only  **String** .


## Syntax

 _expression_ . **SourcePath**

 _expression_ An expression that returns a **ProtectedViewWindow** object.


## Remarks

The path does not include a trailing character (for example, "C:\MSOffice"). Use the [PathSeparator](application-pathseparator-property-word.md) property to add the character that separates folders and drive letters. Use the[SourceName](linkformat-sourcename-property-word.md) property to return the file name without the path.


## Example

The following code example returns the path and name of the document associated with the specified protected view window.


```vb
MsgBox ActiveProtectedViewWindow.SourcePath &; Application.PathSeparator _ 
 &; ActiveProtectedViewWindow.SourceName 

```


## See also


#### Concepts


[ProtectedViewWindow Object](protectedviewwindow-object-word.md)

