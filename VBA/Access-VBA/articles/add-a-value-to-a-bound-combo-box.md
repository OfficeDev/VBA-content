---
title: Add a Value to a Bound Combo Box
ms.prod: access
ms.assetid: a34fddd2-eef6-10e2-c141-609053d1dd90
ms.date: 06/08/2017
---


# Add a Value to a Bound Combo Box

Combo boxes are commonly used to display a list of values in a table or query. By responding to the [NotInList](combobox-notinlist-event-access.md) event, you can provide a way for the user to add values that are not in the list.

Often the value displayed in a combo box is looked up from a record in a related table. Because the list is derived from a table or query, you must provide a way for the user to enter a new record in the underlying table. Then you can use the [Requery](combobox-requery-method-access.md) method to requery the list, so it contains the new value.

When a user types a value in a combo box that is not in the list, the  **NotInList** event of the combo box occurs as long as the combo box's [LimitToList](combobox-limittolist-property-access.md) property is set to **Yes**, or a column other than the combo box's bound column is displayed in the box. You can write an event procedure for the **NotInList** event that provides a way for the user to add a new record to the table that supplies the list's values. The **NotInList** event procedure includes a string argument named _NewData_ that Access uses to pass the text the user enters to the event procedure.

The  **NotInList** event procedure also has a _Response_ argument, in which you tell Access what to do after the procedure runs. Depending on what action you take in the event procedure, you set the _Response_ argument to one of three predefined constant values:


|**Constant**|**Description**|
|:-----|:-----|
|**acDataErrAdded**|If your event procedure enters the new value in the record source for the list or provides a way for the user to do so, set the  _Response_ argument to **acDataErrAdded**. Access then requeries the combo box for you, adding the new value to the list.|
|**acDataErrDisplay**|If you do not add the new value and want Access to display the default error message, set the  _Response_ argument to **acDataErrDisplay**. Access requires the user to enter a valid value from the list.|
|**acDataErrContinue**|If you display your own message in the event procedure, set the  _Response_ argument to **acDataErrContinue**. Access does not display its default error message, but still requires the user to enter a value in the field. If you do not want the user to select an existing value from the list, you can undo changes to the field by using the **Undo** method.|
For example, the following event procedure asks the user whether to add a value to a list, adds the value, and then uses the  _Response_ argument to tell Access to requery the list:



```vb
Private Sub ShipperID_NotInList(NewData As String, Response As Integer)

   Dim dbsOrders As DAO.Database
   Dim rstShippers As DAO.Recordset
   Dim intAnswer As Integer

On Error GoTo ErrorHandler

   intAnswer = MsgBox("Add " &; NewData &; " to the list of shippers?", _
      vbQuestion + vbYesNo)

   If intAnswer = vbYes Then

      ' Add shipper stored in NewData argument to the Shippers table.
      Set dbsOrders = CurrentDb
      Set rstShippers = dbsOrders.OpenRecordset("Shippers")
      rstShippers.AddNew
      rstShippers!CompanyName = NewData
      rstShippers.Update

      Response = acDataErrAdded         ' Requery the combo box list.
   Else
      Response = acDataErrDisplay       ' Require the user to select
                                        ' an existing shipper.
   End If

   rstShippers.Close
   dbsOrders.Close

   Set rstShippers = Nothing
   Set dbsOrders = Nothing

   Exit Sub

ErrorHandler:
   MsgBox "Error #: " &; Err.Number &; vbCrLf &; vbCrLf &; Err.Description
End Sub
```


