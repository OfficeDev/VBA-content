---
title: "How to: Control Character Matching in a Combo Box and List Box"
keywords: olfm10.chm3077212
f1_keywords:
- olfm10.chm3077212
ms.prod: outlook
ms.assetid: 09557015-d581-a09f-499e-86bd4211588e
ms.date: 06/08/2017
---


# How to: Control Character Matching in a Combo Box and List Box

The following example uses the  **MatchEntry** property to demonstrate character matching that is available for **[ComboBox](combobox-object-outlook-forms-script.md)** and **[ListBox](listbox-object-outlook-forms-script.md)**. In this example, the user can set the type of matching with the  **[OptionButton](optionbutton-object-outlook-forms-script.md)** controls and then type into the **ComboBox** to specify an item from its list.

To use this example, copy this sample code to the Script Editor of a form. To run the code you need to open the form so the  **Open** event will activate. Make sure that the form contains:

- Three  **OptionButton** controls named OptionButton1 through OptionButton3.
    
- A  **ComboBox** named ComboBox1.
    



```vb
Sub OptionButton1_Click() 
 Set ComboBox1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("ComboBox1") 
 ComboBox1.MatchEntry = 2 '2=fmMatchEntryNone 
End Sub 
 
Sub OptionButton2_Click() 
 Set ComboBox1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("ComboBox1") 
 ComboBox1.MatchEntry = 0 '0=fmMatchEntryFirstLetter 
End Sub 
 
Sub OptionButton3_Click() 
 Set ComboBox1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("ComboBox1") 
 ComboBox1.MatchEntry = 1 '1=fmMatchEntryComplete 
End Sub 
 
Sub Item_Open() 
 Set ComboBox1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("ComboBox1") 
 Set OptionButton1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("OptionButton1") 
 Set OptionButton2 = Item.GetInspector.ModifiedFormPages("P.2").Controls("OptionButton2") 
 Set OptionButton3 = Item.GetInspector.ModifiedFormPages("P.2").Controls("OptionButton3") 
 
 For i = 1 To 9 
 ComboBox1.AddItem "Choice " &; i 
 Next 
 ComboBox1.AddItem "Chocoholic" 
 
 OptionButton1.Caption = "No matching" 
 OptionButton1.Value = True 
 
 OptionButton2.Caption = "Basic matching" 
 OptionButton3.Caption = "Extended matching" 
End Sub
```


