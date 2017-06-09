---
title: Form.Recalc Method (Access)
keywords: vbaac10.chm13502
f1_keywords:
- vbaac10.chm13502
ms.prod: access
api_name:
- Access.Form.Recalc
ms.assetid: 61786e64-dc17-b685-f427-fc7952d0320f
ms.date: 06/08/2017
---


# Form.Recalc Method (Access)

The  **Recalc** method immediately updates all calculated controls on a form.


## Syntax

 _expression_. **Recalc**

 _expression_ A variable that represents a **Form** object.


### Return Value

Nothing


## Remarks

Using this method is equivalent to pressing the F9 key when a form has the focus. You can use this method to recalculate the values of controls that depend on other fields for which the contents may have changed.


## Example

The following example uses the  **Recalc** method to update controls on an Orders form. This form includes the Freight text box, which displays the freight cost, and a calculated control that displays the total cost of an order including freight. If the statement containing the **Recalc** method is placed in the AfterUpdate event procedure for the Freight text box, the total cost of an order is recalculated every time a new freight amount is entered.


```vb
Sub Freight_AfterUpdate() 
 Me.Recalc 
End Sub
```


## See also


#### Concepts


[Form Object](form-object-access.md)

