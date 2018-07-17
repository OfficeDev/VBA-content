---
title: Application.UsableWidth Property (Word)
keywords: vbawd10.chm158335009
f1_keywords:
- vbawd10.chm158335009
ms.prod: word
api_name:
- Word.Application.UsableWidth
ms.assetid: e5ea64f7-2a7a-fdaa-20ff-fdf6196de874
ms.date: 06/08/2017
---


# Application.UsableWidth Property (Word)

Returns the maximum width (in points) to which you can set the width of a Microsoft Word document window. Read-only  **Long** .


## Syntax

 _expression_ . **UsableWidth**

 _expression_ A variable that represents an **[Application](application-object-word.md)** object.


## Example

This example sets the size of the active document window to one quarter of the maximum allowable screen area.


```vb
With ActiveDocument.ActiveWindow 
 .WindowState = wdWindowStateNormal 
 .Top = 5 
 .Left = 5 
 .Height = (Application.UsableHeight*0.5) 
 .Width = (Application.UsableWidth*0.5) 
End With
```

This example displays the size of the working area in the active document window.




```vb
With ActiveDocument.ActiveWindow 
 MsgBox "Working area height = " _ 
 &; .UsableHeight &; vbLf _ 
 &; "Working area width = " _ 
 &; .UsableWidth 
End With
```


## See also


#### Concepts


[Application Object](application-object-word.md)

