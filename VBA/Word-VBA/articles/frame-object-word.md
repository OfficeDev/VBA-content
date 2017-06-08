---
title: Frame Object (Word)
keywords: vbawd10.chm2346
f1_keywords:
- vbawd10.chm2346
ms.prod: word
api_name:
- Word.Frame
ms.assetid: d36d3361-9e93-7dd9-b8c9-0ce503e03810
ms.date: 06/08/2017
---


# Frame Object (Word)

Represents a frame. The  **Frame** object is a member of the **Frames** collection. The **[Frames](frames-object-word.md)** collection includes all frames in a selection, range, or document.


## Remarks

Use  **Frames** (Index), where Index is the index number, to return a single **Frame** object. The index number represents the position of the frame in the selection, range, or document. The following example allows text to wrap around the first frame in the active document.


```vb
ActiveDocument.Frames(1).TextWrap = True
```

Use the  **Add** method to add a frame around a range. The following example adds a frame around the first paragraph in the active document.




```vb
ActiveDocument.Frames.Add _ 
 Range:=ActiveDocument.Paragraphs(1).Range
```

You can wrap text around  **Shape** or **ShapeRange** objects by using the **WrapFormat** property. You can position a **Shape** or **ShapeRange** object by using the **Top** and **Left** properties.


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


