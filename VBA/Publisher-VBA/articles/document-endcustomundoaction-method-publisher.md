---
title: Document.EndCustomUndoAction Method (Publisher)
keywords: vbapb10.chm196710
f1_keywords:
- vbapb10.chm196710
ms.prod: publisher
api_name:
- Publisher.Document.EndCustomUndoAction
ms.assetid: 5b703366-8d0e-1bbc-3320-a2fea99468c3
ms.date: 06/08/2017
---


# Document.EndCustomUndoAction Method (Publisher)

Specifies the endpoint of a group of actions that are wrapped to create a single undo action. The  ** [BeginCustomUndoAction Method](document-begincustomundoaction-method-publisher.md)** method is used to specify the starting point and label (textual description) of the actions used to create the single undo action. The wrapped group of actions can be undone with a single undo.


## Syntax

 _expression_. **EndCustomUndoAction**

 _expression_A variable that represents a  **Document** object.


## Remarks

The  **BeginCustomUndoAction** method must be called before the **EndCustomUndoAction** method is called. A run-time error is returned if **EndCustomUndoAction** is called before **BeginCustomUndoAction**.


## Example

The following example contains two custom undo actions. The first one is created on page four of the active publication. The  **BeginCustomUndoAction** method is used to specify the point at which the custom undo action should begin. Six individual actions are performed, and then they are wrapped into one action with the call to **EndCustomUndoAction**. 

The text in the text frame that was created within the first custom undo action is then tested to determine whether the font is Verdana. If not, the  **[Undo](document-undo-method-publisher.md)** method is called with **[UndoActionsAvailable](document-undoactionsavailable-property-publisher.md)** passed as a parameter. In this case there is only one undo action available. So, the call to **Undo** will undo only one action, but this one action has wrapped six actions into one.

A second undo action is then created, and it could also be undone later with a single undo operation.

This example assumes that the active publication contains at least four pages.




```vb
Dim thePage As page 
Dim theShape As Shape 
Dim theDoc As Publisher.Document 
 
Set theDoc = ActiveDocument 
Set thePage = theDoc.Pages(4) 
 
With theDoc 
 ' The following six of actions are wrapped to create one 
 ' custom undo action named "Add Rectangle and Courier Text". 
 .BeginCustomUndoAction ("Add Rectangle and Courier Text") 
 With thePage 
 Set theShape = .Shapes.AddShape(msoShapeRectangle, _ 
 75, 75, 190, 30) 
 With theShape.TextFrame.TextRange 
 .Font.Size = 14 
 .Font.Bold = msoTrue 
 .Font.Name = "Courier" 
 .Text = "This font is Courier." 
 End With 
 End With 
 .EndCustomUndoAction 
 
 If Not thePage.Shapes(1).TextFrame.TextRange.Font.Name = "Verdana" Then 
 ' This call to Undo will undo all actions that are available. 
 ' In this case, there is only one action that can be undone. 
 .Undo (.UndoActionsAvailable) 
 ' A new custom undo action is created with a name of 
 ' "Add Balloon and Verdana Text". 
 .BeginCustomUndoAction ("Add Balloon and Verdana Text") 
 With thePage 
 Set theShape = .Shapes.AddShape(msoShapeBalloon, _ 
 75, 75, 190, 30) 
 With theShape.TextFrame.TextRange 
 .Font.Size = 11 
 .Font.Name = "Verdana" 
 .Text = "This font is Verdana." 
 End With 
 End With 
 .EndCustomUndoAction 
 End If 
End With
```


