---
title: Frameset.ParentFrameset Property (Word)
keywords: vbawd10.chm165807083
f1_keywords:
- vbawd10.chm165807083
ms.prod: word
api_name:
- Word.Frameset.ParentFrameset
ms.assetid: aa2759c6-4072-00c6-0c4f-ef12ecc19bd6
ms.date: 06/08/2017
---


# Frameset.ParentFrameset Property (Word)

Returns a  **Frameset** object that represents the parent of the specified **Frameset** object on a frames page.


## Syntax

 _expression_ . **ParentFrameset**

 _expression_ An expression that returns a **[Frameset](frameset-object-word.md)** object.


## Remarks

For more information on creating frames pages, see [Creating Frames Pages](http://msdn.microsoft.com/library/0245564e-b2df-83cd-1e32-e63079970dc1%28Office.15%29.aspx).


## Example

This example returns the number of child  **Frameset** objects belonging to the parent **Frameset** object of the specified frame.


```vb
MsgBox ActiveDocument.ActiveWindow.ActivePane _ 
 .Frameset.ParentFrameset.ChildFramesetCount
```


## See also


#### Concepts


[Frameset Object](frameset-object-word.md)

