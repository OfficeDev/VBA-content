---
title: "How to: Change the Order of Pages and Tabs"
keywords: olfm10.chm3077194
f1_keywords:
- olfm10.chm3077194
ms.prod: outlook
ms.assetid: 5180c30b-e5bb-48b9-ece7-02d5b8d41af0
ms.date: 06/08/2017
---


# How to: Change the Order of Pages and Tabs

The following example uses the  **Index** property to change the order of the pages and tabs in a **[MultiPage](multipage-object-outlook-forms-script.md)** and **[TabStrip](tabstrip-object-outlook-forms-script.md)**. The user chooses CommandButton1 to move the third page and tab to the front of the  **MultiPage** and **TabStrip**. The user chooses CommandButton2 to move the selected page and tab to the back of the  **MultiPage** and **TabStrip**.

To use this example, copy this sample code to the Script Editor of a form. Make sure that the form contains:

- Two  **[CommandButton](commandbutton-object-outlook-forms-script.md)** controls named CommandButton1 and CommandButton2.
    
- A  **MultiPage** named MultiPage1.
    
- A  **TabStrip** named TabStrip1.
    



```vb
Dim MyPageOrTab 
Dim MultiPage1 
Dim TabStrip1 
 
Sub CommandButton1_Click() 
'Move third page and tab to front of control 
 MultiPage1.page3.Index = 0 
 TabStrip1.Tab3.Index = 0 
End Sub 
 
Sub CommandButton2_Click() 
'Move selected page and tab to back of control 
 Set MyPageOrObject = MultiPage1.SelectedItem 
 MsgBox "MultiPage1.SelectedItem = " &; MultiPage1.SelectedItem.Name 
 MyPageOrObject.Index = 4 
 
 Set MyPageOrObject = TabStrip1.SelectedItem 
 MsgBox "TabStrip1.SelectedItem = " &; TabStrip1.SelectedItem.Caption 
 MyPageOrObject.Index = 4 
End Sub 
 
Sub Item_Open() 
 Set MultiPage1 = Item.GetInspector.ModifiedFormPages.Item("P.2").Controls("MultiPage1") 
 Set TabStrip1 = Item.GetInspector.ModifiedFormPages.Item("P.2").Controls("TabStrip1") 
 Set CommandButton1 = Item.GetInspector.ModifiedFormPages.Item("P.2").Controls("CommandButton1") 
 Set CommandButton2 = Item.GetInspector.ModifiedFormPages.Item("P.2").Controls("CommandButton2") 
 
 MultiPage1.Width = 200 
 MultiPage1.Pages.Add 
 MultiPage1.Pages.Add 
 MultiPage1.Pages.Add 
 
 TabStrip1.Width = 200 
 TabStrip1.Tabs.Add 
 TabStrip1.Tabs.Add 
 TabStrip1.Tabs.Add 
 
 CommandButton1.Caption = "Move third page/tab to front" 
 CommandButton1.Width = 120 
 
 CommandButton2.Caption = "Move selected item to back" 
 CommandButton2.Width = 120 
 End Sub
```


