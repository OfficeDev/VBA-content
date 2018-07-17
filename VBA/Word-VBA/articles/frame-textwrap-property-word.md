---
title: Frame.TextWrap Property (Word)
keywords: vbawd10.chm153747468
f1_keywords:
- vbawd10.chm153747468
ms.prod: word
api_name:
- Word.Frame.TextWrap
ms.assetid: 457175c6-4b32-539a-c78d-889647459724
ms.date: 06/08/2017
---


# Frame.TextWrap Property (Word)

 **True** if document text wraps around the specified frame. Read/write **Boolean** .


## Syntax

 _expression_ . **TextWrap**

 _expression_ An expression that returns a **[Frame](frame-object-word.md)** object.


## Example

This example causes text to not wrap around the first frame in the active document.


```vb
If ActiveDocument.Frames.Count >= 1 Then 
 ActiveDocument.Frames(1).TextWrap = False 
End If
```

This example causes text to wrap around all frames in the active document.




```vb
For Each aFrame In ActiveDocument.Frames 
 aFrame.TextWrap = True 
Next aFrame
```


## See also


#### Concepts


[Frame Object](frame-object-word.md)

