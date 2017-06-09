---
title: "How to: Set the Input Style for a Combo Box"
keywords: olfm10.chm3077245
f1_keywords:
- olfm10.chm3077245
ms.prod: outlook
ms.assetid: 3cc10ca8-4c9a-93e5-da88-460198c48c48
ms.date: 06/08/2017
---


# How to: Set the Input Style for a Combo Box

The following example uses the  **[Style](combobox-style-property-outlook-forms-script.md)** property to change the style of user input of a **[ComboBox](combobox-object-outlook-forms-script.md)**. The user chooses a style by selecting an  **[OptionButton](optionbutton-object-outlook-forms-script.md)** control and then types into the **ComboBox** to select an item. When **Style** is _StyleDropDownList_, the user must choose an item from the drop-down list. When  **Style** is _StyleDropDownCombo_, the user can type into the text area to specify an item in the drop-down list.

To use this example, copy this sample code to the Script Editor of a form. To run the code you need to open the form so the  **Open** event will activate. Make sure that the form contains:

- Two  **OptionButton** controls named OptionButton1 and OptionButton2.
    
- A  **ComboBox** named ComboBox1.
    



```vb
Sub OptionButton1_Click() 
 Set ComboBox1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("ComboBox1") 
 ComboBox1.Style = 0 '0=fmStyleDropDownCombo 
End Sub 
 
Sub OptionButton2_Click() 
 Set ComboBox1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("ComboBox1") 
 ComboBox1.Style = 2 '2=fmStyleDropDownList 
End Sub 
 
Sub Item_Open() 
 Set ComboBox1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("ComboBox1") 
 Set OptionButton1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("OptionButton1") 
 Set OptionButton2 = Item.GetInspector.ModifiedFormPages("P.2").Controls("OptionButton2") 
 
 For i = 1 To 10 
 ComboBox1.AddItem "Choice " &; i 
 Next 
 
 OptionButton1.Caption = "Select like ComboBox" 
 OptionButton1.Value = True 
 ComboBox1.Style = 0 '0=fmStyleDropDownCombo 
 
 OptionButton2.Caption = "Select like ListBox" 
End Sub
```


