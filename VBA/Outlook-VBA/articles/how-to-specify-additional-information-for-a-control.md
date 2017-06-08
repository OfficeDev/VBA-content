---
title: "How to: Specify Additional Information for a Control"
keywords: olfm10.chm3077252
f1_keywords:
- olfm10.chm3077252
ms.prod: outlook
ms.assetid: dcbdfec2-ae0c-27d7-6713-9c99fa6e82d6
ms.date: 06/08/2017
---


# How to: Specify Additional Information for a Control

The following example uses the  **Tag** property to store additional information about each control on the Microsoft Forms 2.0 **UserForm**. The user clicks a control and then clicks the  **[CommandButton](commandbutton-object-outlook-forms-script.md)**. The contents of  **Tag** for the appropriate control are returned in the **[TextBox](textbox-object-outlook-forms-script.md)**.

To use this example, copy this sample code to the Script Editor of a form. To run the code you need to open the form so the  **Open** event will activate. Make sure that the form contains:

- A  **TextBox** named TextBox1.
    
- A  **CommandButton** named CommandButton1.
    
- A  **[ScrollBar](scrollbar-object-outlook-forms-script.md)** named ScrollBar1.
    
- A  **[ComboBox](combobox-object-outlook-forms-script.md)** named ComboBox1.
    
- A  **[MultiPage](multipage-object-outlook-forms-script.md)** named MultiPage1.
    



```vb
Sub CommandButton1_Click() 
 Set TextBox1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("TextBox1") 
 Set MultiPage1= Item.GetInspector.ModifiedFormPages("P.2").Controls("MultiPage1") 
 
 TextBox1.Text = Item.GetInspector.ModifiedFormPages("P.2").ActiveControl.Tag 
End Sub 
 
Sub Item_Open() 
 Set TextBox1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("TextBox1") 
 Set CommandButton1= Item.GetInspector.ModifiedFormPages("P.2").Controls("CommandButton1") 
 Set ComboBox1= Item.GetInspector.ModifiedFormPages("P.2").Controls("ComboBox1") 
 Set ScrollBar1= Item.GetInspector.ModifiedFormPages("P.2").Controls("ScrollBar1") 
 Set MultiPage1= Item.GetInspector.ModifiedFormPages("P.2").Controls("MultiPage1") 
 
 TextBox1.Locked = True 
 TextBox1.Tag = "Display area for Tag properties." 
 TextBox1.AutoSize = True 
 
 CommandButton1.Caption = "Show Tag of Current Control." 
 CommandButton1.AutoSize = True 
 CommandButton1.WordWrap = True 
 CommandButton1.TakeFocusOnClick = False 
 CommandButton1.Tag = "Shows tag of control that has the focus." 
 
 ComboBox1.Style = fmStyleDropDownList 
 ComboBox1.Tag = "ComboBox Style is that of a ListBox." 
 
 ScrollBar1.Max = 100 
 ScrollBar1.Min = -273 
 ScrollBar1.Tag = "Max = " &; ScrollBar1.Max &; " , Min = " &; ScrollBar1.Min 
 
 MultiPage1.Pages.Add 
 MultiPage1.Pages.Add 
 MultiPage1.Tag = "This MultiPage has " &; MultiPage1.Pages.Count &; " pages." 
End Sub
```


