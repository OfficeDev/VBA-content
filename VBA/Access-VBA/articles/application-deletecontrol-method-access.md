---
title: Application.DeleteControl Method (Access)
keywords: vbaac10.chm12522
f1_keywords:
- vbaac10.chm12522
ms.prod: access
api_name:
- Access.Application.DeleteControl
ms.assetid: f59f9368-0d7a-8e5f-5140-86e2d2c18c22
ms.date: 06/08/2017
---


# Application.DeleteControl Method (Access)

The  **DeleteControl** method deletes a specified control from a form.


## Syntax

 _expression_. **DeleteControl**( ** _FormName_**, ** _ControlName_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FormName_|Required|**String**|The name of the form containing the control you want to delete.|
| _ControlName_|Required|**String**|The name of the control you want to delete.|

### Return Value

Nothing


## Remarks

For example, suppose you have a procedure that must be run the first time each user logs onto your database. You can set the  **OnClick** property of a button on the form to this procedure. Once the user has logged on and run the procedure, you can use the **DeleteControl** method to dynamically remove the command button from the form.

The  **DeleteControl** method is available only in form Design view or report Design view, respectively.


 **Note**  If you are building a wizard that deletes a control from a form or report, your wizard must open the form or report in Design view before it can delete the control.


## Example

The following example creates a form with a command button and displays a message that asks if the user wants to delete the command button. If the user clicks Yes, the command button is deleted.


```vb
Sub DeleteCommandButton() 
 Dim frm As Form, ctlNew As Control 
 Dim strMsg As String, intResponse As Integer, _ 
 intDialog As Integer 
 
 ' Create new form and get pointer to it. 
 Set frm = CreateForm 
 ' Create new command button. 
 Set ctlNew = CreateControl(frm.Name, acCommandButton) 
 ' Restore form. 
 DoCmd.Restore 
 ' Set caption. 
 ctlNew.Caption = "New Command Button" 
 ' Size control. 
 ctlNew.SizeToFit 
 ' Prompt user to delete control. 
 strMsg = "About to delete " &; ctlNew.Name &;". Continue?" 
 ' Define buttons to be displayed in dialog box. 
 intDialog = vbYesNo + vbCritical + vbDefaultButton2 
 intResponse = MsgBox(strMsg, intDialog) 
 If intResponse = vbYes Then 
 ' Delete control. 
 DeleteControl frm.Name, ctlNew.Name 
 End If 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-access.md)

