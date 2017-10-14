---
title: Application.CodeContextObject Property (Access)
keywords: vbaac10.chm12497
f1_keywords:
- vbaac10.chm12497
ms.prod: access
api_name:
- Access.Application.CodeContextObject
ms.assetid: b675d334-33e6-b845-0dd9-6dca36f7b4ab
ms.date: 06/08/2017
---


# Application.CodeContextObject Property (Access)

You can use the  **CodeContextObject** property to determine the object in which a macro or Visual Basic code is executing. Read-only **Object**.


## Syntax

 _expression_. **CodeContextObject**

 _expression_ A variable that represents an **Application** object.


## Remarks

The  **CodeContextObject** property is set by Microsoft Access and is read-only in all views.

The  **[ActiveControl](screen-activecontrol-property-access.md)**, **[ActiveDatasheet](screen-activedatasheet-property-access.md)**, **[ActiveForm](screen-activeform-property-access.md)**, and **[ActiveReport](screen-activereport-property-access.md)** properties of the **[Screen](screen-object-access.md)** object always return the object that currently has the focus. The object with the focus may or may not be the object where a macro or Visual Basic code is currently running, for example, when Visual Basic code runs in the **[Timer](form-timer-event-access.md)** event on a hidden form.


## Example

In the following example the  **CodeContextObject** property is used in a function to identify the name of the object in which an error occurred. The object name is then used in the message box title as well as in the body of the error message. The **Error** statement is used in the command button's click event to generate the error for this example.


```vb
Private Sub Command1_Click() 
 On Error GoTo Command1_Err 
 Error 11 ' Generate divide-by-zero error. 
 Exit Sub 
 
 Command1_Err: 
 If ErrorMessage("Command1_Click() Event", vbYesNo + _ 
 vbInformation, Err) = vbYes Then 
 Exit Sub 
 Else 
 Resume 
 End If 
End Sub 
 
Function ErrorMessage(strText As String, intType As Integer, _ 
 intErrVal As Integer) As Integer 
 Dim objCurrent As Object 
 Dim strMsgboxTitle As String 
 Set objCurrent = CodeContextObject 
 strMsgboxTitle = "Error in " &; objCurrent.Name 
 strText = strText &; "Error #" &; intErrVal _ 
 &; " occured in " &; objCurrent.Name 
 ErrorMessage = MsgBox(strText, intType, strMsgboxTitle) 
 Err = 0 
End Function
```


## See also


#### Concepts


[Application Object](application-object-access.md)

