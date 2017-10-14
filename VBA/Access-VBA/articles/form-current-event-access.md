---
title: Form.Current Event (Access)
keywords: vbaac10.chm13634
f1_keywords:
- vbaac10.chm13634
ms.prod: access
api_name:
- Access.Form.Current
ms.assetid: 44961599-2b0a-874e-be64-1e29f47f839f
ms.date: 06/08/2017
---


# Form.Current Event (Access)

Occurs when the focus moves to a record, making it the current record, or when the form is refreshed or requeried.


## Syntax

 _expression_. **Current**

 _expression_ A variable that represents a **Form** object.


## Remarks

To run a macro or event procedure when this event occurs, set the  **OnCurrent** property to the name of the macro or to [Event Procedure].

This event occurs both when a form is opened and whenever the focus leaves one record and moves to another. Microsoft Access runs the  **Current** macro or event procedure before the first or next record is displayed.

By running a macro or event procedure when a form's  **Current** event occurs, you can display a message or synchronize records in another form related to the current record. For example, when a customer record becomes current, you can display the customer's previous order. When a supplier record becomes current, you can display the products manufactured by the supplier in a Suppliers form. You can also perform calculations based on the current record or change the form in response to data in the current record.

If your macro or event procedure runs a  **GoToControl** or **GoToRecord** action or the corresponding method of the **DoCmd** object in response to an **Open** event, the **Current** event occurs.

The  **Current** event also occurs when you refresh a form or requery the form's underlying table or query— for example, when you click Remove Filter/Sort on the Records menu or use the Requery action in a macro or the **Requery** method in Visual Basic code.

When you first open a form, the following events occur in this order:

Open → Load → Resize → Activate → Current


## Example

In the following example, a  **Current** event procedure checks the status of an option button called Discontinued. If the button is selected, the example sets the background color of the ProductName field to red to indicate that the product has been discontinued.

To try the example, add the following event procedure to a form that contains an option called Discontinued and a text box called ProductName.




```vb
Private Sub Form_Current() 
 If Me!Discontinued Then 
 Me!ProductName.BackColor = 255 
 EndIf 
End Sub
```


## See also


#### Concepts


[Form Object](form-object-access.md)

