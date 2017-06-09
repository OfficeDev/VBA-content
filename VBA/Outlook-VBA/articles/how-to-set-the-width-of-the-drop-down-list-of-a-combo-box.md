---
title: "How to: Set the Width of the Drop-Down List of a Combo Box"
keywords: olfm10.chm3077209
f1_keywords:
- olfm10.chm3077209
ms.prod: outlook
ms.assetid: 0e36a1fa-482a-9016-9a32-265193bb741e
ms.date: 06/08/2017
---


# How to: Set the Width of the Drop-Down List of a Combo Box

The following example uses a  **[SpinButton](spinbutton-object-outlook-forms-script.md)** to control the width of the drop-down list of a **[ComboBox](combobox-object-outlook-forms-script.md)**. The user changes the value of the  **SpinButton**, then clicks on the drop-down arrow of the  **ComboBox** to display the list.

To use this example, copy this sample code to the Script Editor of a form. To run the code you need to open the form so the  **Open** event will activate. Make sure that the form contains:

- A  **ComboBox** named ComboBox1.
    
- A  **SpinButton** named SpinButton1 that is bound to a custom number field named SpinButtonValue.
    
- A  **[Label](label-object-outlook-forms-script.md)** named Label1.
    



```vb
Sub Item_Open() 
 Set ComboBox1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("ComboBox1") 
 Set SpinButton1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("SpinButton1") 
 Set Label1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("Label1") 
 
 For i = 1 To 20 
 ComboBox1.AddItem "Choice " &; (ComboBox1.ListCount + 1) 
 Next 
 SpinButton1.Min = 0 
 SpinButton1.Max = 130 
 
 'convert listwidth value from '122 pt' to an integer 
 intpos = instr(combobox1.listwidth," ") 
 intwidth = left(combobox1.listwidth,intpos-1) 
 SpinButton1.Value = intwidth 
 SpinButton1.SmallChange = 5 
 Label1.Caption = "ListWidth = " &; SpinButton1.Value 
End Sub 
 
Sub Item_CustomPropertyChange(byval pname) 
 If pname = "SpinButtonValue" Then 
 Set ComboBox1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("ComboBox1") 
 Set SpinButton1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("SpinButton1") 
 Set Label1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("Label1") 
 
 ComboBox1.ListWidth = SpinButton1.Value 
 Label1.Caption = "ListWidth = " &; SpinButton1.Value 
 End If 
End Sub
```


