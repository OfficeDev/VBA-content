---
title: "How to: Display the Number of Pages and Tabs in MultiPage and TabStrip Controls on a Form"
keywords: olfm10.chm3077169
f1_keywords:
- olfm10.chm3077169
ms.prod: outlook
ms.assetid: 9d49b6b3-7650-d96e-9a47-00b508fc6006
ms.date: 06/08/2017
---


# How to: Display the Number of Pages and Tabs in MultiPage and TabStrip Controls on a Form

The following example displays the  **Count** property of the Microsoft Forms 2.0 **Controls**collection for the form, and the  **Count** property identifying the number of pages and tabs of each **[MultiPage](multipage-object-outlook-forms-script.md)** and **[TabStrip](tabstrip-object-outlook-forms-script.md)**.

To use this example, copy this sample code to the Script Editor of a form. To run the code you need to open the form so the  **Open** event will activate. The form can contain any number of controls, with the following restrictions:

- Names of  **MultiPage** controls must start with "MultiPage".
    
- Names of  **TabStrip** controls must start with "TabStrip".
    

 **Note**  You can add pages to a  **MultiPage** or add tabs to a **TabStrip** while in design mode. Double-click on the control, then right click in the tab area of the control and choose **New Page** from the shortcut menu.




```vb
Sub Item_Open 
 Dim Controls 
 Dim MyControl 
 
 Set Controls = Item.GetInspector.ModifiedFormPages.Item("P.2").Controls 
 MsgBox "Controls.Count = " &; Controls.Count 
 For i = 0 to Controls.Count -1 
 Set MyControl = Item.GetInspector.ModifiedFormPages.Item("P.2").Controls.Item(i) 
 If (MyControl.Name = "MultiPage1") Then 
 MsgBox MyControl.Name &; ".Pages.Count = " &; MyControl.Pages.Count 
 ElseIf (MyControl.Name = "TabStrip1") Then 
 MsgBox MyControl.Name &; ".Tabs.Count = " &; MyControl.Tabs.Count 
 End If 
 Next 
End Sub
```


