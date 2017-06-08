---
title: "How to: Control the Style and Selection Behavior of a List Box"
keywords: olfm10.chm3077208
f1_keywords:
- olfm10.chm3077208
ms.prod: outlook
ms.assetid: f4b7c003-55c4-4908-77d0-d6184f6ec786
ms.date: 06/08/2017
---


# How to: Control the Style and Selection Behavior of a List Box

The following example uses the  **[ListStyle](listbox-liststyle-property-outlook-forms-script.md)** and **[MultiSelect](olklistbox-multiselect-property-outlook.md)** properties to control the appearance of a **[ListBox](listbox-object-outlook-forms-script.md)**. The user chooses a value for  **ListStyle** using the **[ToggleButton](togglebutton-object-outlook-forms-script.md)** and chooses an **[OptionButton](optionbutton-object-outlook-forms-script.md)** for one of the **MultiSelect** values. The appearance of the **ListBox** changes accordingly, as well as the selection behavior within the **ListBox**.

To use this example, copy this sample code to the Script Editor of a form. To run the code you need to open the form so the  **Open** event will activate. Make sure that the form contains:

- A  **ListBox** named ListBox1.
    
- A  **[Label](label-object-outlook-forms-script.md)** named Label1.
    
- Three  **OptionButton** controls named OptionButton1 through OptionButton3.
    
- A  **ToggleButton** named ToggleButton1.
    



```vb
Sub Item_Open() 
 Set ListBox1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("ListBox1") 
 Set Label1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("Label1") 
 Set OptionButton1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("OptionButton1") 
 Set OptionButton2 = Item.GetInspector.ModifiedFormPages("P.2").Controls("OptionButton2") 
 Set OptionButton3 = Item.GetInspector.ModifiedFormPages("P.2").Controls("OptionButton3") 
 Set ToggleButton1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("ToggleButton1") 
 
 For i = 1 To 8 
 ListBox1.AddItem "Choice" &; (ListBox1.ListCount + 1) 
 Next 
 
 Label1.Caption = "MultiSelect Choices" 
 Label1.AutoSize = True 
 
 ListBox1.MultiSelect = 0 '0=fmMultiSelectSingle 
 OptionButton1.Caption = "Single entry" 
 OptionButton1.Value = True 
 OptionButton2.Caption = "Multiple entries" 
 OptionButton3.Caption = "Extended entries" 
 
 ToggleButton1.Caption = "ListStyle - Plain" 
 ToggleButton1.Value = True 
 ToggleButton1.Width = 90 
 ToggleButton1.Height = 30 
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
 
Sub ToggleButton1_Click() 
 Set ToggleButton1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("ToggleButton1") 
 Set ListBox1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("ListBox1") 
 
 If ToggleButton1.Value = True Then 
 ToggleButton1.Caption = "Plain ListStyle" 
 ListBox1.ListStyle = 0 '0=fmListStylePlain 
 Else 
 ToggleButton1.Caption = "OptionButton or CheckBox" 
 ListBox1.ListStyle = 1 '1=fmListStyleOption 
 End If 
End Sub
```


