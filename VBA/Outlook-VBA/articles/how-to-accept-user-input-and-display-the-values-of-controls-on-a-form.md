---
title: "How to: Accept User Input and Display the Values of Controls on a Form"
keywords: olfm10.chm3077264
f1_keywords:
- olfm10.chm3077264
ms.prod: outlook
ms.assetid: 5966b34a-7334-a82a-afbc-55d466c06d53
ms.date: 06/08/2017
---


# How to: Accept User Input and Display the Values of Controls on a Form

The following example demonstrates the values that the different types of controls can have by displaying the  **Value** property of a selected control. The user chooses a control by pressing TAB or by clicking on the control. Depending on the type of control, the user can also specify a value for the control by typing in the text area of the control, by clicking one or more times on the control, or by selecting an item, page, or tab within the control. The user can display the value of the selected control by clicking the appropriately labeled **[CommandButton](commandbutton-object-outlook-forms-script.md)**.

To use this example, copy this sample code to the Script Editor of a form. To run the code you need to open the form so the  **Open** event will activate. Make sure that the form contains:

- A  **CommandButton** named CommandButton1.
    
- A  **[TextBox](textbox-object-outlook-forms-script.md)** named TextBox1.
    
- A  **[CheckBox](checkbox-object-outlook-forms-script.md)** named CheckBox1.
    
- A  **[ComboBox](combobox-object-outlook-forms-script.md)** named ComboBox1.
    
- A  **CommandButton** named CommandButton2.
    
- A  **[ListBox](listbox-object-outlook-forms-script.md)** named ListBox1.
    
- A  **[MultiPage](multipage-object-outlook-forms-script.md)** named MultiPage1.
    
- Two  **[OptionButton](optionbutton-object-outlook-forms-script.md)** controls named OptionButton1 and OptionButton2.
    
- A  **[ScrollBar](scrollbar-object-outlook-forms-script.md)** named ScrollBar1.
    
- A  **[SpinButton](spinbutton-object-outlook-forms-script.md)** named SpinButton1.
    
- A  **[TabStrip](tabstrip-object-outlook-forms-script.md)** named TabStrip1.
    
- A  **TextBox** named TextBox2.
    
- A  **[ToggleButton](togglebutton-object-outlook-forms-script.md)** named ToggleButton1.
    



```vb
Sub CommandButton1_Click() 
 Set TextBox1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("TextBox1") 
 Set Form = Item.GetInspector.ModifiedFormPages("P.2") 
 TextBox1.Text = "Value of " &; Form.ActiveControl.Name &; " is " &; Form.ActiveControl.Value 
End Sub 
 
Sub Item_Open() 
 Set CommandButton1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("CommandButton1") 
 Set TextBox1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("TextBox1") 
 Set ComboBox1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("ComboBox1") 
 Set ListBox1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("ListBox1") 
 Set CheckBox1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("CheckBox1") 
 Set ToggleButton1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("ToggleButton1") 
 Set TextBox2 = Item.GetInspector.ModifiedFormPages("P.2").Controls("TextBox2") 
 
 CommandButton1.Caption = "Get value of current control" 
 CommandButton1.AutoSize = True 
 CommandButton1.TakeFocusOnClick = False 
 CommandButton1.TabStop = False 
 
 TextBox1.AutoSize = True 
 
 For i = 0 To 10 
 ComboBox1.AddItem "Choice " &; (i + 1) 
 ListBox1.AddItem "Selection " &; (100 - i) 
 Next 
 
 CheckBox1.TripleState = True 
 ToggleButton1.TripleState = True 
 
 TextBox2.Text = "Enter text here." 
End Sub
```


