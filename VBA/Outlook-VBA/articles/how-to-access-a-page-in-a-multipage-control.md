---
title: "How to: Access a Page in a MultiPage Control"
keywords: olfm10.chm3077150
f1_keywords:
- olfm10.chm3077150
ms.prod: outlook
ms.assetid: dd67169f-2f3e-c93d-b5c5-512e8e83bb7a
ms.date: 06/08/2017
---


# How to: Access a Page in a MultiPage Control

The following example accesses an individual page of a  **[MultiPage](multipage-object-outlook-forms-script.md)** in several ways:


- Using the  **[Pages](pages-object-outlook-forms-script.md)** collection with a numeric index.
    
- Using the name of the individual page in the  **MultiPage**.
    
- Using the  **[SelectedItem](multipage-selecteditem-property-outlook-forms-script.md)** property.
    

To use this example, copy this sample code to the Script Editor of a form. Make sure that the form contains a  **MultiPage** named MultiPage1 and a **[CommandButton](commandbutton-object-outlook-forms-script.md)** named CommandButton1.




```vb
Sub CommandButton1_Click 
 Dim PageName 
 Set MultiPage1 = Item.GetInspector.ModifiedFormPages("P.2").MultiPage1 
 
 For i = 0 To MultiPage1.Count - 1 
 'Use index (numeric or string) 
 MsgBox "MultiPage1.Pages(i).Caption = " &; MultiPage1.Pages(i).Caption 
 MsgBox "MultiPage1.Pages.Item(i).Caption = " &; MultiPage1.Pages.Item(i).Caption 
 
 'Use Page object without referring to Pages collection 
 If i = 0 Then 
 MsgBox "MultiPage1.Page1.Caption = " &; MultiPage1.Page1.Caption 
 ElseIf i = 1 Then 
 MsgBox "MultiPage1.Page2.Caption = " &; MultiPage1.Page2.Caption 
 End If 
 
 'Use SelectedItem Property 
 MultiPage1.Value = i 
 MsgBox "MultiPage1.SelectedItem.Caption = " &; MultiPage1.SelectedItem.Caption 
 Next 
End Sub
```


