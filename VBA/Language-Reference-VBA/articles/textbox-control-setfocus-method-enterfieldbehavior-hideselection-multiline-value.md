---
title: TextBox Control, SetFocus Method, EnterFieldBehavior, HideSelection, MultiLine, Value Properties Example
keywords: fm20.chm5225188
f1_keywords:
- fm20.chm5225188
ms.prod: office
ms.assetid: 144cad11-7ddb-0f46-96fe-8b4da3f665e4
ms.date: 06/08/2017
---


# TextBox Control, SetFocus Method, EnterFieldBehavior, HideSelection, MultiLine, Value Properties Example

The following example demonstrates the  **HideSelection** property in the context of either a single form or more than one form. The user can select text in a **TextBox** and TAB to other controls on a form, as well as transfer the focus to a second form. This code sample also uses the **SetFocus** method, and the **EnterFieldBehavior**, **MultiLine**, and **Value** properties.

To use this example, follow these steps:




1. Copy this sample code (except for the last event subroutine) to the Declarations portion of a form.
    
2. Add a large  **TextBox** named TextBox1, a **ToggleButton** named ToggleButton1, and a **CommandButton** named CommandButton1.
    
3. Insert a second form into this project named UserForm2.
    
4. Paste the last event subroutine of this listing into the Declarations section of UserForm2.
    
5. In this form, add a  **CommandButton** named CommandButton1.
    
6. Run UserForm1.
    




```vb
' ***** Code for UserForm1 ***** 
Private Sub CommandButton1_Click() 
 TextBox1.SetFocus 
 UserForm2.Show 'Bring up the second form. 
End Sub
```




```vb
Private Sub ToggleButton1_Click() 
 If ToggleButton1.Value = True Then 
 TextBox1.HideSelection = False 
 ToggleButton1.Caption = "Selection Visible" 
 Else 
 TextBox1.HideSelection = True 
 ToggleButton1.Caption = "Selection Hidden" 
 End If 
End Sub
```




```vb
Private Sub UserForm_Initialize() 
 TextBox1.MultiLine = True 
 TextBox1.EnterFieldBehavior = fmEnterFieldBehaviorRecallSelection 
 
'Fill the TextBox 
 TextBox1.Text = "SelText indicates the starting " _ 
 &; "point of selected text, or the insertion " _ 
 &; point if no text is selected." &; Chr$(10) _ 
 &; Chr$(13) &; "The SelStart property is " _ 
 &; "always valid, even when the control does " _ 
 &; "not have focus. Setting SelStart to a " _ 
 &; "value less than zero creates an error. " _ 
 &; Chr$(10) &; Chr$(13) &; "Changing the value " _ 
 &; "of SelStart cancels any existing " _ 
 &; "selection in the control, places " _ 
 &; "an insertion point in the text, and sets " _ 
 &; "the SelLength property to zero." 
 
 TextBox1.HideSelection = True 
 ToggleButton1.Caption = "Selection Hidden" 
 ToggleButton1.Value = False 
 
End Sub
```vb
'
' ***** Code for UserForm2 *****



```vb
Private Sub CommandButton1_Click() 
 UserForm2.Hide 
End Sub
```


