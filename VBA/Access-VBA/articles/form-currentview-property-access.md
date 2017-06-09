---
title: Form.CurrentView Property (Access)
keywords: vbaac10.chm13466
f1_keywords:
- vbaac10.chm13466
ms.prod: access
api_name:
- Access.Form.CurrentView
ms.assetid: d173222e-99d1-704e-ee3c-246263725706
ms.date: 06/08/2017
---


# Form.CurrentView Property (Access)

You can use the  **CurrentView** property to determine how a form is currently displayed. Read/write **Integer**.


## Syntax

 _expression_. **CurrentView**

 _expression_ A variable that represents a **Form** object.


## Remarks

The  **CurrentView** property uses the following settings.



|**Setting**|**Form Displayed In:**|
|:-----|:-----|
|0|Design view|
|1|Form view|
|2|Datasheet view|
Use this property to perform different tasks depending on the current view. For example, an event procedure could determine which view the form is displayed in and perform one task if the form is displayed in Form view or another task if it's displayed in Datasheet view.


## Example

The following example uses the GetCurrentView subroutine to determine whether a form is in Form or Datasheet view. If it's in Form view, a message to the user is displayed in a text box on the form; if it's in Datasheet view, the same message is displayed in a message box.


```vb
GetCurrentView Me, "Please contact system administrator." 
 
Sub GetCurrentView(frm As Form, strDisplayMsg As String) 
 Const conFormView = 1 
 Const conDataSheet = 2 
 Dim intView As Integer 
 intView = frm.CurrentView 
 Select Case intView 
 Case conFormView 
 frm!MessageTextBox.SetFocus 
 ' Display message in text box. 
 frm!MessageTextBox = strDisplayMsg 
 Case conDataSheet 
 ' Display message in message box. 
 MsgBox strDisplayMsg 
 End Select 
End Sub
```


## See also


#### Concepts


[Form Object](form-object-access.md)

