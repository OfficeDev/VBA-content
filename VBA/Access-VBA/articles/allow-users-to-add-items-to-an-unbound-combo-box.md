---
title: Allow Users to Add Items to an Unbound Combo Box
ms.prod: access
ms.assetid: 654cefc7-cbd4-5e8e-adc7-919c6977ac6a
ms.date: 06/08/2017
---


# Allow Users to Add Items to an Unbound Combo Box

The following example uses the  **NotInList** event to add an item to a combo box.

To try this example, create a combo box called  **Colors** on a form. Set the combo box's **[LimitToList](combobox-limittolist-property-access.md)** property to **Yes**. To populate the combo box, set the combo box's **[RowSourceType](combobox-rowsourcetype-property-access.md)** property to **Value List**, and supply a list of values separated by semicolons as the setting for the **[RowSource](combobox-rowsource-property-access.md)** property. For example, you might supply the following values as the setting for this property: Red; Green; Blue.

Next, add the following event procedure to the form. Switch to Form view and enter a new value in the text portion of the combo box. 




```vb
Private Sub Colors_NotInList(NewData As String, _ 
        Response As Integer) 
    Dim ctl As Control 
     
    ' Return Control object that points to combo box. 
    Set ctl = Me!Colors 
    ' Prompt user to verify they want to add new value. 
    If MsgBox("Value is not in list. Add it?", _ 
         vbOKCancel) = vbOK Then 
        ' Set Response argument to indicate that data 
        ' is being added. 
        Response = acDataErrAdded 
        ' Add string in NewData argument to row source. 
        ctl.RowSource = ctl.RowSource &; ";" &; NewData 
    Else 
    ' If user chooses Cancel, suppress error message 
    ' and undo changes. 
        Response = acDataErrContinue 
        ctl.Undo 
    End If 
End Sub
```


 **Note**  This example adds an item to an unbound combo box. When you add an item to a bound combo box, you add a value to a field in the underlying data source. In most cases you cannot simply add one field in a new record. Depending on the structure of data in the table, you probably will need to add one or more fields to fulfill data requirements. For instance, a new record must include values for any fields comprising the primary key. If you need to add items to a bound combo box dynamically, you must prompt the user to enter data for all required fields, save the new record, and then requery the combo box to display the new value. 


