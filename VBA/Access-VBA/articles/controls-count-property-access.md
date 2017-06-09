---
title: Controls.Count Property (Access)
keywords: vbaac10.chm10180
f1_keywords:
- vbaac10.chm10180
ms.prod: access
api_name:
- Access.Controls.Count
ms.assetid: 531c1674-4782-aa8f-64f5-0493a29886e3
ms.date: 06/08/2017
---


# Controls.Count Property (Access)

You can use the  **Count** property to determine the number of items in a specified collection. Read-only **Long**.


## Syntax

 _expression_. **Count**

 _expression_ A variable that represents a **Controls** object.


## Example

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


[Controls Collection](controls-object-access.md)

