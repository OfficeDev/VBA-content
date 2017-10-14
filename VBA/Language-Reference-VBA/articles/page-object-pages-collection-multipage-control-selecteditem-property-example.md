---
title: Page Object, Pages Collection, MultiPage Control, SelectedItem Property Example
keywords: fm20.chm5225175
f1_keywords:
- fm20.chm5225175
ms.prod: office
ms.assetid: 85bf4dd6-a291-27b4-7f67-811e28ade6e9
ms.date: 06/08/2017
---


# Page Object, Pages Collection, MultiPage Control, SelectedItem Property Example

The following example accesses an individual page of a  **MultiPage** in several ways:



- Using the  **Pages** collection with a numeric index.
    
- Using the  **Pages** collection with a string index.
    
- Using the  **Pages** collection with the **Item** method.
    
- Using the name of the individual page in the  **MultiPage**.
    
- Using the  **SelectedItem** property.
    

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains a  **MultiPage** named MultiPage1.



```vb
Private Sub UserForm_Initialize() 
 Dim PageName As String 
 
 For i = 0 To MultiPage1.Count - 1 
 'Use index (numeric or string) 
 MsgBox "MultiPage1.Pages(i).Caption = " _ 
 &; MultiPage1.Pages(i).Caption 
 MsgBox "MultiPage1.Pages.Item(i).Caption = " _ 
 &; MultiPage1.Pages.Item(i).Caption 
 
 PageName = MultiPage1.Pages(i).Name 
 MsgBox "PageName = " &; PageName 
 
 MsgBox "MultiPage1.Pages(PageName)" _ 
 &; ".Caption = "_ 
 &; MultiPage1.Pages(PageName).Caption 
 MsgBox "MultiPage1.Pages.Item(PageName)" _ 
 &; ".Caption = " &; MultiPage1.Pages _ 
 .Item(PageName).Caption 
 
 'Use Page object without referring to 
 'Pages collection 
 If i = 0 Then 
 MsgBox "MultiPage1.Page1.Caption= " _ 
 &; MultiPage1.Page1.Caption 
 ElseIf i = 1 Then 
 MsgBox "MultiPage1.Page2.Caption = " _ 
 &; MultiPage1.Page2.Caption 
 End If 
 
 'Use SelectedItem Property 
 MultiPage1.Value = i 
 MsgBox "MultiPage1.SelectedItem.Caption = " _ 
 &; MultiPage1.SelectedItem.Caption 
 Next i 
End Sub
```


