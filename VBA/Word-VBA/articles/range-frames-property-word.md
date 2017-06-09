---
title: Range.Frames Property (Word)
keywords: vbawd10.chm157155394
f1_keywords:
- vbawd10.chm157155394
ms.prod: word
api_name:
- Word.Range.Frames
ms.assetid: c30bb71d-3998-42fe-2850-a76c3975418b
ms.date: 06/08/2017
---


# Range.Frames Property (Word)

Returns a  **[Frames](frames-object-word.md)** collection that represents all the frames in a range. Read-only.


## Syntax

 _expression_ . **Frames**

 _expression_ A variable that represents a **[Range](range-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example causes text to wrap around frames in the first section in the active document.


```vb
For Each aFrame In ActiveDocument.Sections(1).Range.Frames 
 aFrame.TextWrap = True 
Next aFrame
```


## See also


#### Concepts


[Range Object](range-object-word.md)

