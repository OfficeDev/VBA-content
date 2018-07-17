---
title: Application.ScreenUpdating Property (Word)
keywords: vbawd10.chm158335002
f1_keywords:
- vbawd10.chm158335002
ms.prod: word
api_name:
- Word.Application.ScreenUpdating
ms.assetid: 660284d1-2b00-eba0-28bc-36bc6400fcf4
ms.date: 06/08/2017
---


# Application.ScreenUpdating Property (Word)

 **True** if screen updating is turned on. Read/write **Boolean** .


## Syntax

 _expression_ . **ScreenUpdating**

 _expression_ An expression that returns an **[Application](application-object-word.md)** object.


## Remarks

The  **ScreenUpdating** property controls most display changes on the monitor while a procedure is running. When screen updating is turned off, toolbars remain visible and Word still allows the procedure to display or retrieve information using status bar prompts, input boxes, dialog boxes, and message boxes. You can increase the speed of some procedures by keeping screen updating turned off. You must set the **ScreenUpdating** property to **True** when the procedure finishes or when it stops after an error.


## Example

This example turns off screen updating and then adds a new document. Five hundred lines of text are added to the document. At every fiftieth line, the macro selects the line and refreshes the screen.


```vb
Application.ScreenUpdating = False 
Documents.Add 
For x = 1 To 500 
 With ActiveDocument.Content 
 .InsertAfter "This is line " &; x &; "." 
 .InsertParagraphAfter 
 End With 
If x Mod 50 = 0 Then 
 ActiveDocument.Paragraphs(x).Range.Select 
 Application.ScreenRefresh 
End If 
Next x 
Application.ScreenUpdating = True
```


## See also


#### Concepts


[Application Object](application-object-word.md)

