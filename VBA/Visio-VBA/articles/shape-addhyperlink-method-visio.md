---
title: Shape.AddHyperlink Method (Visio)
ms.prod: visio
api_name:
- Visio.Shape.AddHyperlink
ms.assetid: fbf77a65-88a1-e710-60a2-efde9e7df968
ms.date: 06/08/2017
---


# Shape.AddHyperlink Method (Visio)

Adds a  **Hyperlink** object to a Microsoft Visio shape.


## Syntax

 _expression_ . **AddHyperlink**

 _expression_ A variable that represents a **Shape** object.


### Return Value

Hyperlink


## Remarks

Using the  **AddHyperlink** method is equivalent to adding a hyperlink to a shape by clicking **Hyperlink** on the **Insert** tab.

If a  **Hyperlink** object already exists for the shape, the method returns a reference to the existing **Hyperlink** object.


## Example

This example shows how to use the  **AddHyperlink** method to add a hyperlink to a shape. It also shows how to trap errors that arise when you try to access nonexistent hyperlinks. It first attempts to access a hyperlink that does not exist, thereby throwing an error. Then it adds the hyperlink, and when it attempts to access the hyperlink a second time, no error is thrown. Before running this example, replace _address_ with a valid Internet or intranet address.


```vb
 
Sub AddHyperlink_Example() 
 
 Dim vsoShape As Visio.Shape 
 Dim vsoHyperlink As Visio.Hyperlink 
 Dim blsCaught As Boolean 
 
 'Draw a rectangle shape on the active page. 
 Set vsoShape = ActivePage.DrawRectangle(1, 2, 2, 1) 
 
 'A shape that has no hyperlink should raise an exception 
 'when the Hyperlink property is accessed. 
 On Error GoTo lblCatch 
 
 blsCaught = False 
 Set vsoHyperlink = vsoShape.Hyperlink 
 
 If Not blsCaught Then 
 Debug.Print "ERROR - Hyperlink didn't throw an exception!" 
 End If 
 
 'Add a hyperlink to a shape. 
 Set vsoHyperlink = vsoShape.AddHyperlink 
 
 'Return an absolute address. 
 vsoHyperlink.Address = "address " 
 
 Exit Sub 
 
 lblCatch: 
 Debug.Print "Error was thrown : " &; Err.Description 
 blsCaught = True 
 Resume Next 
End Sub
```


