
# Shape.SendToBack Method (Visio)

 **Last modified:** July 28, 2015

 _**Applies to:** Visio 2013 Preview_

Moves the shape or selected shapes to the back of the z-order.


## Syntax

 _expression_. **SendToBack**

 _expression_A variable that represents a  **Shape** object.


### Return Value

Nothing


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to move a shape to the back of the z-order on a page.


```
 
Public Sub SendToBack_Example() 
 
 Dim vsoShape1 As Visio.Shape 
 Dim vsoShape2 As Visio.Shape 
 Dim vsoShape3 As Visio.Shape 
 
 'Draw three rectangles. 
 Set vsoShape1 = ActivePage.DrawRectangle(1, 1, 5, 5) 
 vsoShape1.Text = "1" 
 Set vsoShape2 = ActivePage.DrawRectangle(2, 2, 6, 6) 
 vsoShape2.Text = "2" 
 Set vsoShape3 = ActivePage.DrawRectangle(3, 3, 7, 7) 
 vsoShape3.Text = "3" 
 
 'Move vsoShape3 to the back of the z-order. 
 vsoShape3.SendToBack 
 
End Sub 

```

