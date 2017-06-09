---
title: "How to: Add, Cut, and Paste a Control on a Page"
keywords: olfm10.chm3077170
f1_keywords:
- olfm10.chm3077170
ms.prod: outlook
ms.assetid: f20fb2d9-0ee2-2cf5-173c-9fdd6201bdca
ms.date: 06/08/2017
---


# How to: Add, Cut, and Paste a Control on a Page

The following example uses the Microsoft Forms 2.0  **Controls**collection, and the  **Controls.Add**,  **Controls.Cut**, and  **[Page.Paste](page-paste-method-outlook-forms-script.md)** methods to add, cut, and paste a control on a **[Page](page-object-outlook-forms-script.md)** of a **[MultiPage](multipage-object-outlook-forms-script.md)**. The control involved in the cut and paste operations is dynamically added to the form.

This example assumes the user will add, then cut, then paste the new control.

To use this example, copy this sample code to the Script Editor of a form. To run the code you need to open the form so the  **Open** event will activate. Make sure that the form contains:


- Three  **[CommandButton](commandbutton-object-outlook-forms-script.md)** controls named CommandButton1 through CommandButton3.
    
- A  **MultiPage** named MultiPage1.
    



```vb
Dim CommandButton1 
Dim CommandButton2 
Dim CommandButton3 
Dim MultiPage1 
Dim MyTextBox 
 
Sub CommandButton1_Click() 
 Set MyTextBox = MultiPage1.Pages(MultiPage1.Value).Controls.Add("Forms.TextBox.1", "MyTextBox", 1) 
 CommandButton2.Enabled = True 
 CommandButton1.Enabled = False 
End Sub 
 
Sub CommandButton2_Click() 
 MultiPage1.Pages(MultiPage1.Value).Controls.Cut 
 CommandButton3.Enabled = True 
 CommandButton2.Enabled = False 
End Sub 
 
Sub CommandButton3_Click() 
 Dim MyPage 
 Set MyPage = MultiPage1.Pages.Item(MultiPage1.Value) 
 
 MyPage.Paste 
 CommandButton3.Enabled = False 
End Sub 
 
Sub Item_Open() 
 Set CommandButton1 = Item.GetInspector.ModifiedFormPages.Item("P.2").Controls("CommandButton1") 
 Set CommandButton2 = Item.GetInspector.ModifiedFormPages.Item("P.2").Controls("CommandButton2") 
 Set CommandButton3 = Item.GetInspector.ModifiedFormPages.Item("P.2").Controls("CommandButton3") 
 Set MultiPage1 = Item.GetInspector.ModifiedFormPages.Item("P.2").Controls("MultiPage1") 
 
 CommandButton1.Caption = "Add" 
 CommandButton2.Caption = "Cut" 
 CommandButton3.Caption = "Paste" 
 
 CommandButton1.Enabled = True 
 CommandButton2.Enabled = False 
 CommandButton3.Enabled = False 
End Sub
```


