---
title: "How to: Display the List Portion of a Combo Box"
keywords: olfm10.chm3077183
f1_keywords:
- olfm10.chm3077183
ms.prod: outlook
ms.assetid: 9edcd472-eeaa-c7ef-7d15-369f50c9fe31
ms.date: 06/08/2017
---


# How to: Display the List Portion of a Combo Box

The following example uses the  **[DropDown](combobox-dropdown-method-outlook-forms-script.md)** method to display the list in a **[ComboBox](combobox-object-outlook-forms-script.md)**. The user can display the list of a  **ComboBox** by clicking the **[CommandButton](commandbutton-object-outlook-forms-script.md)**.

To use this example, copy this sample code to the Script Editor of a form. To run the code you need to open the form so the  **Open** event will activate. Make sure that the form contains:

- A  **ComboBox** named ComboBox1.
    
- A  **CommandButton** named CommandButton1.
    



```vb
Dim ComboBox1 
 
Sub CommandButton1_Click() 
 ComboBox1.DropDown 
End Sub 
 
Sub Item_Open() 
 Set ComboBox1 = Item.GetInspector.ModifiedFormPages.Item("P.2").Controls("ComboBox1") 
 ComboBox1.AddItem "Turkey" 
 ComboBox1.AddItem "Chicken" 
 ComboBox1.AddItem "Duck" 
 ComboBox1.AddItem "Goose" 
 ComboBox1.AddItem "Grouse" 
End Sub
```


