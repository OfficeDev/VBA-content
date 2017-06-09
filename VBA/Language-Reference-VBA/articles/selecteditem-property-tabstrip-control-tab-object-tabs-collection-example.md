---
title: SelectedItem Property, TabStrip Control, Tab Object, Tabs Collection Example
keywords: fm20.chm5225157
f1_keywords:
- fm20.chm5225157
ms.prod: office
ms.assetid: 7480356d-77dd-c925-a784-d7388e2bfba9
ms.date: 06/08/2017
---


# SelectedItem Property, TabStrip Control, Tab Object, Tabs Collection Example

The following example accesses an individual tab of a  **TabStrip** in several ways:



- Using the  **Tabs** collection with a numeric index.
    
- Using the  **Tabs** collection with a string index.
    
- Using the  **Tabs** collection with the **Item** method.
    
- Using the name of the individual  **Tab**.
    
- Using the  **SelectedItem** property.
    

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains a  **TabStrip** named TabStrip1.



```vb
Private Sub UserForm_Initialize() 
 Dim TabName As String 
 
 For i = 0 To TabStrip1.Count - 1 
 'Using index (numeric or string) 
 MsgBox "TabStrip1.Tabs(i).Caption = " _ 
 &; TabStrip1.Tabs(i).Caption 
 MsgBox "TabStrip1.Tabs.Item(i).Caption = " _ 
 &; TabStrip1.Tabs.Item(i).Caption 
 
 TabName = TabStrip1.Tabs(i).Name 
 MsgBox "TabName = " &; TabName 
 
 MsgBox "TabStrip1.Tabs(TabName).Caption = " _ 
 &; TabStrip1.Tabs(TabName).Caption 
 MsgBox "TabStrip1.Tabs.Item(TabName)_ 
 .Caption = " _ 
 &; TabStrip1.Tabs.Item(TabName).Caption 
 
 'Use Tab object without referring to Tabs 
 'collection 
 If i = 0 Then 
 MsgBox "TabStrip1.Tab1.Caption = " _ 
 &; TabStrip1.Tab1.Caption 
 ElseIf i = 1 Then 
 MsgBox "TabStrip1.Tab2.Caption = " _ 
 &; TabStrip1.Tab2.Caption 
 EndIf 
 
 'Use SelectedItem Property 
 TabStrip1.Value = i 
 MsgBox "TabStrip1.SelectedItem.Caption = " _ 
 &; TabStrip1.SelectedItem.Caption 
 Next i 
End Sub
```


