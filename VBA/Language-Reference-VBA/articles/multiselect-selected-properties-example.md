---
title: MultiSelect, Selected Properties Example
keywords: fm20.chm5225122
f1_keywords:
- fm20.chm5225122
ms.prod: office
ms.assetid: 643b50df-51b2-bf19-7f18-4429a3f363f2
ms.date: 06/08/2017
---


# MultiSelect, Selected Properties Example

The following example uses the  **MultiSelect** and **Selected** properties to demonstate how the user can select one or more items in a **ListBox**. The user specifies a selection method by choosing an option button and then selects an item(s) from the **ListBox**. The user can display the selected items in a second **ListBox** by clicking the **CommandButton**.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:




- Two  **ListBox** controls named ListBox1 and ListBox2.
    
- A  **CommandButton** named CommandButton1.
    
- Three  **OptionButton** controls named OptionButton1 through OptionButton3.
    




```vb
Dim i As Integer 
 
Private Sub CommandButton1_Click() 
 ListBox2.Clear 
 
 For i = 0 To 9 
 If ListBox1.Selected(i) = True Then 
 ListBox2.AddItem ListBox1.List(i) 
 End If 
 Next i 
 
End Sub 
 
Private Sub OptionButton1_Click() 
 ListBox1.MultiSelect = fmMultiSelectSingle 
End Sub 
 
Private Sub OptionButton2_Click() 
 ListBox1.MultiSelect = fmMultiSelectMulti 
End Sub 
 
Private Sub OptionButton3_Click() 
 ListBox1.MultiSelect = fmMultiSelectExtended 
End Sub 
 
Private Sub UserForm_Initialize() 
 For i = 0 To 9 
 ListBox1.AddItem "Choice " &; (ListBox1.ListCount + 1) 
 Next i 
 
 OptionButton1.Caption = "Single Selection" 
 ListBox1.MultiSelect = fmMultiSelectSingle 
 OptionButton1.Value = True 
 
 OptionButton2.Caption = "Multiple Selection" 
 OptionButton3.Caption = "Extended Selection" 
 
 CommandButton1.Caption = "Show selections" 
 CommandButton1.AutoSize = True 
End Sub
```


