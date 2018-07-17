---
title: Forms.Count Property (Access)
keywords: vbaac10.chm12359
f1_keywords:
- vbaac10.chm12359
ms.prod: access
api_name:
- Access.Forms.Count
ms.assetid: 915dcb5c-bab5-956f-329e-63a6bf934991
ms.date: 06/08/2017
---


# Forms.Count Property (Access)

You can use the  **Count** property to determine the number of items in a specified collection. Read-only **Long**.


## Syntax

 _expression_. **Count**

 _expression_ A variable that represents a **Forms** object.


## Example

For example, if you want to determine the number of forms currently open or existing on the database, you would use the following code strings


```vb
' Determine the number of open forms. 
 
forms.count 
 
' Determine the number of forms (open or closed) 
' in the current database. 
 
currentproject.allforms.count
```

The following example uses the  **Count** property to control a loop that prints information about all open forms and their controls.




```vb
Sub Print_Form_Controls() 
 Dim frm As Form, intI As Integer 
 Dim intJ As Integer 
 Dim intControls As Integer, intForms As Integer 
 intForms = Forms.Count ' Number of open forms. 
 If intForms > 0 Then 
 For intI = 0 To intForms - 1 
 Set frm = Forms(intI) 
 Debug.Print frm.Name 
 intControls = frm.Count 
 If intControls > 0 Then 
 For intJ = 0 To intControls - 1 
 Debug.Print vbTab; frm(intJ).Name 
 Next intJ 
 Else 
 Debug.Print vbTab; "(no controls)" 
 End If 
 Next intI 
 Else 
 MsgBox "No open forms.", vbExclamation, "Form Controls" 
 End If 
End Sub
```


## See also


#### Concepts


[Forms Collection](forms-object-access.md)

