---
title: "How to: Allow Multiple Selections in a List Box"
keywords: olfm10.chm3077221
f1_keywords:
- olfm10.chm3077221
ms.prod: outlook
ms.assetid: da59a685-c118-00b1-8a08-b9d19a15aa77
ms.date: 06/08/2017
---


# How to: Allow Multiple Selections in a List Box

The following example uses the  **[MultiSelect](listbox-multiselect-property-outlook-forms-script.md)** and **[Selected](listbox-selected-property-outlook-forms-script.md)** properties to demonstrate how the user can select one or more items in a **[ListBox](listbox-object-outlook-forms-script.md)**. The user specifies a selection method by choosing an option button and then selects an item(s) from the  **ListBox**. The user can display the selected items in a second  **ListBox** by clicking the **[CommandButton](commandbutton-object-outlook-forms-script.md)**.

To use this example, copy this sample code to the Script Editor of a form. To run the code you need to open the form so the  **Open** event will activate. Make sure that the form contains:

- Two  **ListBox** controls named ListBox1 and ListBox2.
    
- A  **CommandButton** named CommandButton1.
    
- Three  **[OptionButton](optionbutton-object-outlook-forms-script.md)** controls named OptionButton1 through OptionButton3.
    



```vb
Sub CommandButton1_Click() 
 Set ListBox1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("ListBox1") 
 Set ListBox2 = Item.GetInspector.ModifiedFormPages("P.2").Controls("ListBox2") 
 ListBox2.Clear 
 
 For i = 0 To 9 
 If ListBox1.Selected(i) = True Then 
 ListBox2.AddItem ListBox1.List(i) 
 End If 
 Next 
 
End Sub 
 
Sub OptionButton1_Click() 
 Set ListBox1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("ListBox1") 
 ListBox1.MultiSelect = 0 '0=fmMultiSelectSingle 
End Sub 
 
Sub OptionButton2_Click() 
 Set ListBox1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("ListBox1") 
 ListBox1.MultiSelect = 1 '1=fmMultiSelectMulti 
End Sub 
 
Sub OptionButton3_Click() 
 Set ListBox1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("ListBox1") 
 ListBox1.MultiSelect = 2 '2=fmMultiSelectExtended 
End Sub 
 
Sub Item_Open() 
 Set ListBox1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("ListBox1") 
 Set OptionButton1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("OptionButton1") 
 Set OptionButton2 = Item.GetInspector.ModifiedFormPages("P.2").Controls("OptionButton2") 
 Set OptionButton3 = Item.GetInspector.ModifiedFormPages("P.2").Controls("OptionButton3") 
 Set CommandButton1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("CommandButton1") 
 
 For i = 0 To 9 
 ListBox1.AddItem "Choice " &; (ListBox1.ListCount + 1) 
 Next 
 
 OptionButton1.Caption = "Single Selection" 
 ListBox1.MultiSelect = 0 '0=fmMultiSelectSingle 
 OptionButton1.Value = True 
 
 OptionButton2.Caption = "Multiple Selection" 
 OptionButton3.Caption = "Extended Selection" 
 
 CommandButton1.Caption = "Show selections" 
 CommandButton1.AutoSize = True 
End Sub
```


