---
title: ComboBox.NotInList Event (Access)
keywords: vbaac10.chm14214
f1_keywords:
- vbaac10.chm14214
ms.prod: access
api_name:
- Access.ComboBox.NotInList
ms.assetid: 1c8a73e1-ca69-ae31-c86a-c1dc6cb3e860
ms.date: 06/08/2017
---


# ComboBox.NotInList Event (Access)

The  **NotInList** event occurs when the user enters a value in the text box portion of a combo box that isn't in the combo box list.


## Syntax

 _expression_. **NotInList**( **_NewData_**, **_Response_** )

 _expression_ A variable that represents a **ComboBox** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _NewData_|Required|**String**|A string that Microsoft Access uses to pass the text the user entered in the text box portion of the combo box to the event procedure.|
| _Response_|Required|**Integer**|The setting indicates how the  **NotInList** event was handled. The _Response_ argument can be one of the following intrinsic constants: <ul><li>**acDataErrDisplay** (Default) Displays the default message to the user. You can use this when you don't want to allow the user to add a new value to the combo box list.</li><li>**acDataErrContinue** Doesn't display the default message to the user. You can use this when you want to display a custom message to the user. For example, the event procedure could display a custom dialog box asking if the user wanted to save the new entry. If the response is Yes, the event procedure would add the new entry to the list and set the **Response** argument to **acDataErrAdded**. If the response is No, the event procedure would set the **Response** argument to **acDataErrContinue**.</li><li>**acDataErrAdded** Doesn't display a message to the user but enables you to add the entry to the combo box list in the **NotInList**  event procedure. After the entry is added, Microsoft Access updates the list by re-querying the combo box. Microsoft Access then rechecks the string against the combo box list, and saves the value in the **NewData** argument in the field the combo box is bound to. If the string is not in the list, then Microsoft Access displays an error message.</li></ul>|

## Remarks

To run a macro or event procedure when this event occurs, set the  **[OnNotInList](combobox-onnotinlist-property-access.md)** property to the name of the macro or to [Event Procedure].

This event enables the user to add a new value to the combo box list.

The  **[LimitToList](combobox-limittolist-property-access.md)** property must be set to Yes for the **NotInList** event to occur.

The  **NotInList** event doesn't trigger the **Error** event.

The  **NotInList** event occurs for combo boxes whose **LimitToList** property is set to Yes, after you enter a value that isn't in the list and attempt to move to another control or save the record. The event occurs after all the **Change** events for the combo box.

When the  **[AutoExpand](combobox-autoexpand-property-access.md)** property is set to Yes, Microsoft Access selects matching values in the list as the user enters characters in the text box portion of the combo box. If the characters the user types match the first characters of a value in the list (for example, the user types "Smith" and "Smithson" is a value in the list), the **NotInList** event will not occur when the user moves to another control or saves the record. However, the characters that Microsoft Access adds to the characters the user types (in the example, "son") are selected in the text box portion of the combo box. If the user wants the **NotInList** event to fire in such cases (for example, the user wants to add the new name "Smith" to the combo box list), the user can enter a **SPACE**,  **BACKSPACE**, or  **DEL** character after the last character in the new value.

When the  **LimitToList** property is set to Yes and the combo box list is dropped down, Microsoft Access selects matching values in the list as the user enters characters in the text box portion of the combo box, even if the **AutoExpand** property is set to No. If the user presses **ENTER** or moves to another control or record, the selected value appears in the combo box. In this case, the **NotInList** event will not fire. To allow the **NotInList** event to fire, the user should not drop down the combo box list.


## Example

The following example uses the  **NotInList** event to add an item to a combo box.

To try this example, create a combo box called Colors on a form. Set the combo box's  **LimitToList** property to Yes. To populate the combo box, set the combo box's **RowSourceType** property to Value List, and supply a list of values separated by semicolons as the setting for the **RowSource** property. For example, you might supply the following values as the setting for this property: Red; Green; Blue.

Next add the following event procedure to the form. Switch to Form view and enter a new value in the text portion of the combo box. 

|**Note**|
|:-----|
|This example adds an item to an unbound combo box. When you add an item to a bound combo box, you add a value to a field in the underlying data source. In most cases you can't simply add one field in a new record ? depending on the structure of data in the table, you probably will need to add one or more fields to fulfill data requirements. For instance, a new record must include values for any fields comprising the primary key. If you need to add items to a bound combo box dynamically, you must prompt the user to enter data for all required fields, save the new record, and then re-query the combo box to display the new value.|



```vb
Private Sub Colors_NotInList(NewData As String, _ 
        Response As Integer) 
    Dim ctl As Control 
     
    ' Return Control object that points to combo box. 
    Set ctl = Me!Colors 
    ' Prompt user to verify they wish to add new value. 
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



The following example shows how to use the  **NotInList** event to add an item to a bound combo box.

 **Sample code provided by:** Bill Jelen, [MrExcel.com](http://www.mrexcel.com/)




```vb
Private Sub cboDept_NotInList(NewData As String, Response As Integer)
    Dim oRS As DAO.Recordset, i As Integer, sMsg As String
    Dim oRSClone As DAO.Recordset

    Response = acDataErrContinue

    If MsgBox("Add dept?", vbYesNo) = vbYes Then
        Set oRS = CurrentDb.OpenRecordset("tblDepartments", dbOpenDynaset)
        oRS.AddNew
        oRS.Fields(1) = NewData
        For i = 2 To oRS.Fields.Count - 1
            sMsg = "What do you want for " &; oRS(i).Name
            oRS(i).Value = InputBox(sMsg, , oRS(i).DefaultValue)
        Next i
        oRS.Update
        cboDept = Null
        cboDept.Requery
        DoCmd.OpenTable "tblDepartments", acViewNormal, acReadOnly
        DoCmd.GoToRecord acDataTable, "tblDepartments", acLast
    End If
End Sub
```

The following example shows how to add an item to a bound combo box.

 **Sample code provided by:** The [Microsoft Access 2010 Programmer's Reference](http://www.wrox.com/WileyCDA/WroxTitle/Access-2010-Programmer-s-Reference.productCd-0470591668.html)




```vb
Private Sub cboMainCategory_NotInList(NewData As String, Response As Integer)

    On Error GoTo Error_Handler
    Dim intAnswer As Integer
    intAnswer = MsgBox("""" &; NewData &; """ is not an approved category. " &; vbcrlf _
        &; "Do you want to add it now?" _ vbYesNo + vbQuestion, "Invalid Category")

    Select Case intAnswer
        Case vbYes
            DoCmd.SetWarnings False
            DoCmd.RunSQL "INSERT INTO tlkpCategoryNotInList (Category) "
                &; _ "Select """ &; NewData &; """;"
            DoCmd.SetWarnings True
            Response = acDataErrAdded
        Case vbNo
            MsgBox "Please select an item from the list.", _
                vbExclamation + vbOKOnly, "Invalid Entry"
            Response = acDataErrContinue

    End Select

    Exit_Procedure:
        DoCmd.SetWarnings True
        Exit Sub

    Error_Handler:
        MsgBox Err.Number &; ", " &; Error Description
        Resume Exit_Procedure
        Resume

End Sub
```


## About the Contributors
<a name="AboutContributors"> </a>

Holy Macro! Books publishes entertaining books for people who use Microsoft Office. See the complete catalog at MrExcel.com. 

Wrox Press is driven by the Programmer to Programmer philosophy. Wrox books are written by programmers for programmers, and the Wrox brand means authoritative solutions to real-world programming problems. 


## See also
<a name="AboutContributors"> </a>


#### Concepts


[ComboBox Object](combobox-object-access.md)

