---
title: Form.OpenArgs Property (Access)
keywords: vbaac10.chm13428
f1_keywords:
- vbaac10.chm13428
ms.prod: access
api_name:
- Access.Form.OpenArgs
ms.assetid: f18ed66f-01e0-b8a3-a15b-687e738aafe6
ms.date: 06/08/2017
---


# Form.OpenArgs Property (Access)

Determines the string expression specified by the  _OpenArgs_ argument of the **OpenForm** method that opened a form. Read/write **Variant**.


## Syntax

 _expression_. **OpenArgs**

 _expression_ A variable that represents a **Form** object.


## Remarks

This property is available only by using a macro or by using Visual Basic with the  **OpenForm** method of the **DoCmd** object. This property setting is read-only in all views.

To use the  **OpenArgs** property, open a form by using the **OpenForm** method of the **DoCmd** object and set the _OpenArgs_ argument to the desired string expression. The **OpenArgs** property setting can then be used in code for the form, such as in an **Open** event procedure. You can also refer to the property setting in a macro, such as an **Open** macro, or an expression, such as an expression that sets the **ControlSource** property for a control on the form.

For example, suppose that the form you open is a continuous-form list of clients. If you want the focus to move to a specific client record when the form opens, you can set the  **OpenArgs** property to the client's name, and then use the **FindRecord** action in an Open macro to move the focus to the record for the client with the specified name.


## Example

The following example uses the  **OpenArgs** property to open the Employees form to a specific employee record and demonstrates how the **OpenForm** method sets the **OpenArgs** property. You can run this procedure as appropriate — for example, when the **AfterUpdate** event occurs for a custom dialog box used to enter new information about an employee.


```vb
Sub OpenToCallahan() 
    DoCmd.OpenForm "Employees", acNormal, , , acReadOnly, _ 
     , "Callahan" 
End Sub 
 
Sub Form_Open(Cancel As Integer) 
    Dim strEmployeeName As String 
    ' If OpenArgs property contains employee name, find 
    ' corresponding employee record and display it on form. For 
    ' example,if the OpenArgs property contains "Callahan", 
    ' move to first "Callahan" record. 
    strEmployeeName = Forms!Employees.OpenArgs 
    If Len(strEmployeeName) > 0 Then 
        DoCmd.GoToControl "LastName" 
        DoCmd.FindRecord strEmployeeName, , True, , True, , True 
    End If 
End Sub
```

The following example shows how to use the  **OpenArgs** property to prevent a form from being opened from the Navigation Pane.

 **Sample code provided by:** The[Microsoft Access 2010 Programmer's Reference](http://www.wrox.com/WileyCDA/WroxTitle/Access-2010-Programmer-s-Reference.productCd-0470591668.html)

```vb
Private Sub Form_Open(Cancel As Integer)

If Me.OpenArgs() <> "Valid User" Then
    MsgBox "You are not authorized to use this form!", _
        vbExclamation + vbOKOnly, "Invalid Access"
    Cancel = True
End If
End Sub
```


## About the Contributors
<a name="AboutContributors"> </a>

Wrox Press is driven by the Programmer to Programmer philosophy. Wrox books are written by programmers for programmers, and the Wrox brand means authoritative solutions to real-world programming problems. 


## See also
<a name="AboutContributors"> </a>


#### Concepts


[Form Object](form-object-access.md)

