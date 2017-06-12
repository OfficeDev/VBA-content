---
title: "How to: Add a Control to a MultiPage Control"
keywords: olfm10.chm3077153
f1_keywords:
- olfm10.chm3077153
ms.prod: outlook
ms.assetid: 9fd9a559-ece9-26dd-047c-c3c649347257
ms.date: 06/08/2017
---


# How to: Add a Control to a MultiPage Control

The following example uses the  **Add**,  **Clear**, and  **Remove** methods of the Microsoft Forms 2.0 **Controls** collection to add a control to and remove a control from a **[Page](page-object-outlook-forms-script.md)** of a **[MultiPage](multipage-object-outlook-forms-script.md)** at run time.

To use this example, copy this sample code to the Script Editor of a form. To run the code you need to open the form so the  **Open** event will activate. Make sure that the form contains:

- A  **MultiPage** named MultiPage1.
    
- Three  **[CommandButton](commandbutton-object-outlook-forms-script.md)** controls named CommandButton1 through CommandButton3.
    



```vb
Dim MyTextBox 
Dim MultiPage1 
 
Sub Item_Open() 
 Set MyPage = Item.GetInspector.ModifiedFormPages("P.2") 
 Set MultiPage1 = MyPage.MultiPage1 
 MyPage.CommandButton1.Caption = "Add control" 
 MyPage.CommandButton2.Caption = "Clear controls" 
 MyPage.CommandButton3.Caption = "Remove control" 
End Sub 
 
Sub CommandButton1_Click() 
 Set MyTextBox = MultiPage1.Pages(0).Controls.Add("Forms.TextBox.1", "MyTextBox", 1) 
End Sub 
 
Sub CommandButton2_Click() 
 MultiPage1.Pages(0).Controls.Clear 
End Sub 
 
Sub CommandButton3_Click() 
 If MultiPage1.Pages(0).Controls.Count > 0 Then 
 MultiPage1.Pages(0).Controls.Remove "MyTextBox" 
 End If 
End Sub
```


