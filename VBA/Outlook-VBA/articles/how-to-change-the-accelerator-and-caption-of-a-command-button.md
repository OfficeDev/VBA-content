---
title: "How to: Change the Accelerator and Caption of a Command Button"
keywords: olfm10.chm3077148
f1_keywords:
- olfm10.chm3077148
ms.prod: outlook
ms.assetid: 5f763d6a-e376-1088-04c8-fbd3a43de4e4
ms.date: 06/08/2017
---


# How to: Change the Accelerator and Caption of a Command Button

This example changes the  **[Accelerator](commandbutton-accelerator-property-outlook-forms-script.md)** and **[Caption](commandbutton-caption-property-outlook-forms-script.md)** properties of a **[CommandButton](commandbutton-object-outlook-forms-script.md)** each time the user clicks the button by using the mouse or the accelerator key. The **[Click](commandbutton-click-event-outlook-forms-script.md)** event contains the code to change the **Accelerator** and **Caption** properties.

To try this example, paste the code into the Script Editor of a form containing a  **CommandButton** named CommandButton1. To run the code you need to open the form so the **Open** event will activate.



```vb
Dim CommandButton1 
 
Sub Item_Open() 
 Set CommandButton1 = Item.GetInspector.ModifiedFormPages("P.2").CommandButton1 
 CommandButton1.Accelerator = "C" 'Set Accelerator key to ALT + C 
End Sub 
 
Sub CommandButton1_Click () 
 If CommandButton1.Caption = "OK" Then 'Check caption, then change it. 
 CommandButton1.Caption = "Clicked" 
 CommandButton1.Accelerator = "C" 'Set Accelerator key to ALT + C 
 Else 
 CommandButton1.Caption = "OK" 
 CommandButton1.Accelerator = "O" 'Set Accelerator key to ALT + O 
 End If 
End Sub
```


