---
title: Application.MailMergeWizardStateChange Event (Word)
keywords: vbawd10.chm4000023
f1_keywords:
- vbawd10.chm4000023
ms.prod: word
api_name:
- Word.Application.MailMergeWizardStateChange
ms.assetid: d112d3f1-7fe7-1db6-891b-917598eea2ef
ms.date: 06/08/2017
---


# Application.MailMergeWizardStateChange Event (Word)

Occurs when a user changes from a specified step to a specified step in the Mail Merge Wizard.


## Syntax

 _expression_ . **Private Sub object_MailMergeWizardStateChange**( **_ByVal Doc As Document_** , **_FromState As Long_** , **_ToState As Long_** , **_Handled As Boolean_** )

 _expression_ A variable that represents an **[Application](application-object-word.md)** object that has been declared with events in a class module. For information about using events with the **Application** object, see[Using Events with the Application Object](http://msdn.microsoft.com/library/784c4c61-7e47-3dbf-46f6-da655f786ca1%28Office.15%29.aspx).


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Doc_|Required| **Document**|The mail merge main document.|
| _FromState_|Required| **Long**|The Mail Merge Wizard step from which a user is moving.|
| _ToState_|Required| **Long**|The Mail Merge Wizard step to which a user is moving.|
| _Handled_|Required| **Boolean**| **True** moves the user to the next step. **False** for the user to remain at the current step.|

## Example

This example displays a message when a user moves from step three of the Mail Merge Wizard to step four. Based on the answer to the message, the user will either move to step four or remain at step three. This example assumes that you have declared an application variable called MailMergeApp in your general declarations and have set the variable equal to the Word Application object.


```vb
Private Sub MailMergeApp_MailMergeWizardStateChange(ByVal Doc As Document, _ 
 FromState As Long, ToState As Long, Handled As Boolean) 
 
 Dim intVBAnswer As Integer 
 FromState = 3 
 ToState = 4 
 
 'Display a message when moving from step three to step four 
 intVBAnswer = MsgBox("Have you selected all of your recipients?", _ 
 vbYesNo, "Wizard State Event!") 
 
 If intVBAnswer = vbYes Then 
 'Continue on to step four 
 Handled = True 
 Else 
 'Return to step three 
 MsgBox "Please select all recipients to whom " &; _ 
 "you want to send this letter." 
 Handled = False 
 End If 
 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-word.md)

