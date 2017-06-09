---
title: Form.Error Event (Access)
keywords: vbaac10.chm13658
f1_keywords:
- vbaac10.chm13658
ms.prod: access
api_name:
- Access.Form.Error
ms.assetid: ed8229fb-4169-8be5-dc2e-a543ca3bfff3
ms.date: 06/08/2017
---


# Form.Error Event (Access)

The Error event occurs when a run-time error is produced in Microsoft Access when a form has the focus.


## Syntax

 _expression_. **Error**( ** _DataErr_**, ** _Response_** )

 _expression_ A variable that represents a **Form** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _DataErr_|Required|**Integer**|The error code returned by the Err object when an error occurs. You can use the DataErr argument with the Error function to map the number to the corresponding error message. |
| _Response_|Required|**Integer**|The setting determines whether or not an error message is displayed. The Response argument can be one of the following intrinsic constants.<ul><li><p><b>acDataErrContinue</b>  Ignore the error and continue without displaying the default Microsoft Access error message. You can supply a custom error message in place of the default error message.</p></li><li><p><b>acDataErrDisplay</b>  (Default) Display the default Microsoft Access error message.</p></li></ul>|

## Remarks

This includes Microsoft Access database engine errors, but not run-time errors in Visual Basic or errors from ADO.

To run a macro or event procedure when this event occurs, set the  **OnError** property to the name of the macro or to [Event Procedure].

By running an event procedure or a macro when an Error event occurs, you can intercept a Microsoft Access error message and display a custom message that conveys a more specific meaning for your application.

 **Link provided by:**
![Community Member Icon](images/8b9774c4-6c97-470e-b3a2-56d8f786444c.png) The[UtterAccess](http://www.utteraccess.com) community


- [Handling Access Errors with VBA](http://www.utteraccess.com/wiki/index.php/Error_Handling)
    

## Example

The following example shows how you can replace a default error message with a custom error message. When Microsoft Access returns an error message indicating it has found a duplicate key (error code 3022), this event procedure displays a message that gives more application-specific information to users.

To try the example, add the following event procedure to a form that is based on a table with a unique employee ID number as the key for each record.




```vb
Private Sub Form_Error(DataErr As Integer, Response As Integer) 
    Const conDuplicateKey = 3022 
    Dim strMsg As String 
 
    If DataErr = conDuplicateKey Then 
        Response = acDataErrContinue 
        strMsg = "Each employee record must have a unique " _ 
            &; "employee ID number. Please recheck your data." 
        MsgBox strMsg 
    End If 
End Sub
```



The following example shows how you can replace a default error message with a custom error message.

 **Sample code provided by:**
![MVP Contributor](images/odc_OfficeTA_33px_MVPContrib.jpg) Bill Jelen, [MrExcel.com](http://www.mrexcel.com/)




```vb
Private Sub Form_Error(DataErr As Integer, Response As Integer)
    Select Case DataErr
        Case 2113
            MsgBox "Only numbers are acceptable in this box", vbCritical, "Call 1-800-123-4567"
            Response = acDataErrContinue
        Case 2237
            MsgBox "You can only choose from the dropdown box"
            Response = acDataErrContinue
        Case 3022
            MsgBox "You entered a value that exists already in another record"
            Response = acDataErrContinue
            SSN.Value = SSN.OldValue
        Case 3314
            MsgBox "The DOH is required, so you cannot leave this field empty"
            Response = acDataErrContinue
        Case Else
            Response = acDataErrDisplay
    End Select
    ActiveControl.Undo
End Sub
```


## About the Contributors
<a name="AboutContributors"> </a>

UtterAccess is the premier Microsoft Access wiki and help forum. 

Holy Macro! Books publishes entertaining books for people who use Microsoft Office. See the complete catalog at MrExcel.com. 


## See also
<a name="AboutContributors"> </a>


#### Concepts


[Form Object](form-object-access.md)

