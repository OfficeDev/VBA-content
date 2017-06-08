---
title: Index Property Example
keywords: fm20.chm5225145
f1_keywords:
- fm20.chm5225145
ms.prod: office
ms.assetid: 7e2a502c-386d-cc3d-842e-8fbbe95e2518
ms.date: 06/08/2017
---


# Index Property Example

The following example uses the  **Index** property to change the order of the pages and tabs in a **MultiPage** and **TabStrip**. The user chooses CommandButton1 to move the third page and tab to the front of the **MultiPage** and **TabStrip**. The user chooses CommandButton2 to move the selected page and tab to the back of the **MultiPage** and **TabStrip**.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:




- Two  **CommandButton** controls named CommandButton1 and CommandButton2.
    
- A  **MultiPage** named MultiPage1.
    
- A  **TabStrip** named TabStrip1.
    




```vb
Dim MyPageOrTab As Object 
 
Private Sub CommandButton1_Click() 
'Move third page and tab to front of control 
 MultiPage1.page3.Index = 0 
 TabStrip1.Tab3.Index = 0 
End Sub 
 
Private Sub CommandButton2_Click() 
'Move selected page and tab to back of control 
 Set MyPageOrObject = MultiPage1.SelectedItem 
 MsgBox "MultiPage1.SelectedItem = " _ 
 &; MultiPage1.SelectedItem.Name 
 MyPageOrObject.Index = 4 
 
 Set MyPageOrObject = TabStrip1.SelectedItem 
 MsgBox "TabStrip1.SelectedItem = " _ 
 &; TabStrip1.SelectedItem.Caption 
 MyPageOrObject.Index = 4 
End Sub 
 
Private Sub UserForm_Initialize() 
 MultiPage1.Width = 200 
 MultiPage1.Pages.Add 
 MultiPage1.Pages.Add 
 MultiPage1.Pages.Add 
 
 TabStrip1.Width = 200 
 TabStrip1.Tabs.Add 
 TabStrip1.Tabs.Add 
 TabStrip1.Tabs.Add 
 
 CommandButton1.Caption = _ 
 "Move third page/tab to front" 
 CommandButton1.Width = 120 
 
 CommandButton2.Caption = _ 
 "Move selected item to back" 
 CommandButton2.Width = 120 
 End Sub
```


