---
title: Selection.Frames Property (Word)
keywords: vbawd10.chm158662722
f1_keywords:
- vbawd10.chm158662722
ms.prod: word
api_name:
- Word.Selection.Frames
ms.assetid: cc589559-858a-2ebb-00dd-64f97966859f
ms.date: 06/08/2017
---


# Selection.Frames Property (Word)

Returns a  **[Frames](frames-object-word.md)** collection that represents all the frames in a selection. Read-only.


## Syntax

 _expression_ . **Frames**

 _expression_ A variable that represents a **[Selection](selection-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example causes text to wrap around frames in the first section in the active document.


```vb
For Each aFrame In ActiveDocument.Sections(1).Range.Frames 
 aFrame.TextWrap = True 
Next aFrame
```

This example adds a frame around the selection and returns a frame object to the myFrame variable.




```vb
Set myFrame = ActiveDocument.Frames.Add(Range:=Selection.Range)
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

