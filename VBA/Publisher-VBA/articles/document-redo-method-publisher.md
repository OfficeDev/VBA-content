---
title: Document.Redo Method (Publisher)
keywords: vbapb10.chm196708
f1_keywords:
- vbapb10.chm196708
ms.prod: publisher
api_name:
- Publisher.Document.Redo
ms.assetid: 4b76aeaa-77f7-5f22-ff80-77479b0f0702
ms.date: 06/08/2017
---


# Document.Redo Method (Publisher)

Redoes the last action or a specified number of actions. Corresponds to the list of items that appears when you click the arrow beside the  **Redo** button on the **Standard** toolbar. Calling this method reverses the ** [Undo Method](document-undo-method-publisher.md)** method.


## Syntax

 _expression_. **Redo**( **_Count_**)

 _expression_A variable that represents a  **Document** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Count|Optional| **Long**|Specifies the number of actions to be redone. Default is 1, meaning that if omitted, only the last action will be redone.|

### Return Value

Nothing


## Remarks

If called when there are no actions on the redo stack, or when  **_Count_** is greater than the number of actions that currently reside on the stack, the **Redo** method will redo as many actions as possible and ignore the rest.

The maximum number of actions that can be redone in one call to  **Redo** is 20.


## Example

The following example uses the  **Redo** method to redo a subset of the actions that were undone using the **Undo** method.

Part 1 creates a rectangle that contains a text frame on the fourth page of the active publication. Various font properties are set, and text is added to the text frame. In this case, the text "This font is Courier" is set to 12 point bold Courier font. 

Part 2 tests whether the text in the text frame is Verdana font. If not, then the  **Undo** method is used to undo the last four actions on the undo stack. The **Redo** method is then used to redo the the first two of the last four actions that were just undone. In this case, the third action (setting the font size) and the fourth action (setting the font to bold) are redone. The font name is then changed to Verdana, and the text is modified.




```vb
Dim thePage As page 
Dim theShape As Shape 
Dim theDoc As Publisher.Document 
 
Set theDoc = ActiveDocument 
Set thePage = theDoc.Pages(4) 
 
' Part 1 
With theDoc 
 With thePage 
 ' Setting the shape creates the first action 
 Set theShape = .Shapes.AddShape(msoShapeRectangle, _ 
 75, 75, 190, 30) 
 ' Setting the text range creates the second action 
 With theShape.TextFrame.TextRange 
 ' Setting the font size creates the third action 
 .Font.Size = 12 
 ' Setting the font to bold creates the fourth action 
 .Font.Bold = msoTrue 
 ' Setting the font name creates the fifth action 
 .Font.Name = "Courier" 
 ' Setting the text creates the sixth action 
 .Text = "This font is Courier." 
 End With 
 End With 
 
 ' Part 2 
 If Not thePage.Shapes(1).TextFrame.TextRange.Font.Name = "Verdana" Then 
 .Undo (4) 
 With thePage 
 With theShape.TextFrame.TextRange 
 ' Redo redoes the first two of the four actions that were just undone 
 theDoc.Redo (2) 
 .Font.Name = "Verdana" 
 .Text = "This font is Verdana." 
 End With 
 End With 
 End If 
End With
```


