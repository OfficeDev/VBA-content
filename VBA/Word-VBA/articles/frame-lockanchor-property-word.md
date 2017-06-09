---
title: Frame.LockAnchor Property (Word)
keywords: vbawd10.chm153747462
f1_keywords:
- vbawd10.chm153747462
ms.prod: word
api_name:
- Word.Frame.LockAnchor
ms.assetid: 654dc51d-12bb-4168-f737-69f8de7da17a
ms.date: 06/08/2017
---


# Frame.LockAnchor Property (Word)

 **True** if the specified frame is locked. Read/write **Boolean** .


## Syntax

 _expression_ . **LockAnchor**

 _expression_ Required. A variable that represents a **[Frame](frame-object-word.md)** object.


## Remarks

The frame anchor indicates where the frame will appear in Normal view. You cannot reposition a locked frame anchor.


## Example

This example locks the anchor of the first frame in section two of the active document.


```vb
Set myRange = ActiveDocument.Sections(2).Range 
If TypeName(myRange) <> "Nothing" And myRange.Frames.Count > 0 Then 
 myRange.Frames(1).LockAnchor = True 
End If
```


## See also


#### Concepts


[Frame Object](frame-object-word.md)

