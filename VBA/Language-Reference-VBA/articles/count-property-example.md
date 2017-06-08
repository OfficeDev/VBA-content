---
title: Count Property Example
keywords: fm20.chm5225150
f1_keywords:
- fm20.chm5225150
ms.prod: office
ms.assetid: acf6e338-c85e-dacc-0ef7-696bb430b3f5
ms.date: 06/08/2017
---


# Count Property Example

The following example displays the  **Count** property of the **Controls** collection for the form, and the **Count** property identifying the number of pages and tabs of each **MultiPage** and **TabStrip**.

To use this example, copy this sample code to the Declarations portion of a form. The form can contain any number of controls, with the following restrictions:




- Names of  **MultiPage** controls must start with "MultiPage".
    
- Names of  **TabStrip** controls must start with "TabStrip".
    


 **Note**  You can add pages to a  **MultiPage** or add tabs to a **TabStrip**. In Windows, double-click the control, then right-click in the tab area of the control and choose **New Page** from the shortcut menu.




```vb
Private Sub UserForm_Initialize() 
 Dim MyControl As Control 
 
 MsgBox "UserForm1.Controls.Count = " _ 
 &; Controls.Count 
 
 For Each MyControl In Controls 
 If (MyControl.Name Like "MultiPage*") Then 
 MsgBox MyControl.Name _ 
 &; ".Pages.Count = " _ 
 &; MyControl.Pages.Count 
 ElseIf (MyControl.Name Like "TabStrip*") Then 
 MsgBox MyControl.Name &; ".Tabs.Count = " _ 
 &; MyControl.Tabs.Count 
 End If 
 Next 
 
End Sub
```


