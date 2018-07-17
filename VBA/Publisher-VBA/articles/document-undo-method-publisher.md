---
title: Document.Undo Method (Publisher)
keywords: vbapb10.chm196704
f1_keywords:
- vbapb10.chm196704
ms.prod: publisher
api_name:
- Publisher.Document.Undo
ms.assetid: 8cfd09a0-8a0d-2870-f833-a35ff1fc21b4
ms.date: 06/08/2017
---


# Document.Undo Method (Publisher)

Undoes the last action or a specified number of actions. Corresponds to the list of items that appears when you click the arrow beside the  **Undo** button on the **Standard** toolbar.


## Syntax

 _expression_. **Undo**( **_Count_**)

 _expression_A variable that represents a  **Document** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Count|Optional| **Long**|Specifies the number of actions to be undone. Default is 1, meaning that if omitted, only the last action will be undone.|

## Remarks

If called when there are no actions on the undo stack, or when  **_Count_** is greater than the number of actions that currently reside on the stack, the **Undo** method will undo as many actions as possible and ignore the rest.

The maximum number of actions that can be undone in one call to  **Undo** is 20.


## Example

The following example uses the  **Undo** method to undo actions that do not meet specific criteria.

Part 1 of the example adds a rectangular callout shape to the fourth page of the active publication, and text is added to the callout. This process creates three actions. 

Part 2 of the example tests whether the font of the text added to the callout is Verdana. If not, then the  **Undo** method is used to undo all available actions (the value of the **[UndoActionsAvailable](document-undoactionsavailable-property-publisher.md)** property is used to specify that all actions be undone). This clears all actions from the stack. A new rectangle shape and text frame are then added and the text frame is populated with Verdana text.




```vb
Dim thePage As page 
Dim theShape As Shape 
Dim theDoc As Publisher.Document 
 
Set theDoc = ActiveDocument 
Set thePage = theDoc.Pages(4) 
 
With theDoc 
 ' Part 1 
 With thePage 
 ' Setting the shape creates the first action 
 Set theShape = .Shapes.AddShape(msoShapeRectangularCallout, _ 
 75, 75, 120, 30) 
 ' Setting the text range creates the second action 
 With theShape.TextFrame.TextRange 
 ' Setting the text creates the third action 
 .Text = "This text is not Verdana." 
 End With 
 End With 
 
 ' Part 2 
 If Not thePage.Shapes(1).TextFrame.TextRange.Font.Name = "Verdana" Then 
 ' UndoActionsAvailable = 3 
 .Undo (.UndoActionsAvailable) 
 With thePage 
 Set theShape = .Shapes.AddShape(msoShapeRectangle, _ 
 75, 75, 120, 30) 
 With theShape.TextFrame.TextRange 
 .Font.Name = "Verdana" 
 .Text = "This text is Verdana." 
 End With 
 End With 
 End If 
End With
```


