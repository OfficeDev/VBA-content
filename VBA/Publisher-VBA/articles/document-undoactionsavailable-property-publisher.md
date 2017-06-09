---
title: Document.UndoActionsAvailable Property (Publisher)
keywords: vbapb10.chm196726
f1_keywords:
- vbapb10.chm196726
ms.prod: publisher
api_name:
- Publisher.Document.UndoActionsAvailable
ms.assetid: 1dd20295-3987-c36d-ccc1-9e18a7887f33
ms.date: 06/08/2017
---


# Document.UndoActionsAvailable Property (Publisher)

Returns the number of actions available on the undo stack. Read-only  **Long**.


## Syntax

 _expression_. **UndoActionsAvailable**

 _expression_A variable that represents an  **Document** object.


### Return Value

Long


## Example

The following example adds a rectangle that contains a text frame to the fourth page of the active publication. Some font properties and the text of the text frame are set. A test is then run to determine whether the font in the text frame is Courier. If so, the  **[Undo](document-undo-method-publisher.md)** method is used with the value of the **UndoActionsAvailable** property passed as a parameter to specify that all previous actions be undone.

The  **[Redo](document-redo-method-publisher.md)** method is then used with the value of the **[RedoActionsAvailable](document-redoactionsavailable-property-publisher.md)** property minus 2 passed as a parameter to redo all actions except for the last two. A new font is specified for the text in the text frame, in addition to new text.

This example assumes the active document contains at least four pages.




```vb
Dim thePage As page 
Dim theShape As Shape 
Dim theDoc As Publisher.Document 
 
Set theDoc = ActiveDocument 
Set thePage = theDoc.Pages(4) 
 
With theDoc 
 With thePage 
 Set theShape = .Shapes.AddShape(msoShapeRectangle, _ 
 75, 75, 190, 30) 
 With theShape.TextFrame.TextRange 
 .Font.Size = 12 
 .Font.Name = "Courier" 
 .Text = "This font is Courier." 
 End With 
 End With 
 
 If thePage.Shapes(1).TextFrame.TextRange.Font.Name = "Courier" Then 
 .Undo (.UndoActionsAvailable) 
 .Redo (.RedoActionsAvailable - 2) 
 With theShape.TextFrame.TextRange 
 .Font.Name = "Verdana" 
 .Text = "This font is Verdana." 
 End With 
 End If 
End With
```


