---
title: "How to: Set Entry Matching in a Combo Box"
keywords: olfm10.chm3077213
f1_keywords:
- olfm10.chm3077213
ms.prod: outlook
ms.assetid: 1e47c76a-a152-30a4-96a6-f95122209ff1
ms.date: 06/08/2017
---


# How to: Set Entry Matching in a Combo Box

The following example uses the  **[MatchFound](combobox-matchfound-property-outlook-forms-script.md)** and **[MatchRequired](combobox-matchrequired-property-outlook-forms-script.md)** properties to demonstrate additional character matching for **[ComboBox](combobox-object-outlook-forms-script.md)**. The matching verification occurs in the  **Change** event.

In this example, the user specifies whether the text portion of a  **ComboBox** must match one of the listed items in the **ComboBox**. The user can specify whether matching is required by using a  **[CheckBox](checkbox-object-outlook-forms-script.md)** and then type into the **ComboBox** to specify an item from its list.

To use this example, copy this sample code to the Script Editor of a form. To run the code you need to open the form so the  **Open** event will activate. Make sure that the form contains:


- A  **ComboBox** named ComboBox1 that is bound to the Subject field.
    
- A  **CheckBox** named CheckBox1.
    



```vb
Sub CheckBox1_Click() 
 Set ComboBox1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("ComboBox1") 
 Set CheckBox1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("CheckBox1") 
 
 If CheckBox1.Value = True Then 
 ComboBox1.MatchRequired = True 
 MsgBox "To move the focus from the ComboBox, you must match an entry in the list or press ESC." 
 Else 
 ComboBox1.MatchRequired = False 
 MsgBox " To move the focus from the ComboBox, just tab to or click another control. Matching is optional." 
 End If 
End Sub 
 
Sub Item_PropertyChange(byval pname) 
 if pname = "Subject" then 
 Set ComboBox1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("ComboBox1") 
 If ComboBox1.MatchRequired = True Then 
 'MSForms handles this case automatically 
 Else 
 If ComboBox1.MatchFound = True Then 
 MsgBox "Match Found; matching optional." 
 Else 
 MsgBox "Match not Found; matching optional." 
 End If 
 End If 
 end if 
End Sub 
 
Sub Item_Open() 
 Set ComboBox1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("ComboBox1") 
 Set CheckBox1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("CheckBox1") 
 
 For i = 1 To 9 
 ComboBox1.AddItem "Choice " &; i 
 Next 
 ComboBox1.AddItem "Chocoholic" 
 
 CheckBox1.Caption = "MatchRequired" 
 CheckBox1.Value = True 
End Sub
```


