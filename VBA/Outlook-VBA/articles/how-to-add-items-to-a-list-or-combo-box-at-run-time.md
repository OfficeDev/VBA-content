---
title: "How to: Add Items To a List or Combo Box at Run Time"
keywords: olfm10.chm3077344
f1_keywords:
- olfm10.chm3077344
ms.prod: outlook
ms.assetid: 5dd25eb3-8c36-3e71-30ae-b35638ef6943
ms.date: 06/08/2017
---


# How to: Add Items To a List or Combo Box at Run Time

In a  **[ListBox](listbox-object-outlook-forms-script.md)** or **[ComboBox](combobox-object-outlook-forms-script.md)** with a single column, use the **AddItem** method to add an individual entry to the list.

In a multicolumn list box or combo box, you can use the  **List** and **Column** properties to load the list from a two-dimensional array, as shown in the following steps.

1. Create a multicolumn  **ListBox** or **ComboBox** control.
    
2. In VBScript, create a two-dimensional array that contains the items you want to put in the list.
    
3. Set the  **ColumnCount** property of the list box or combo box to match the number of entries in the list. To set the property, click the property and enter a value in the **Apply** box.
    
4. Do one of the following:
    
      - Assign the array as the value of the  **List** property. The contents of the list box will match the contents of the array exactly.
    
  - Assign the array as the value of the  **Column** property. **Column** transposes rows and columns, so each row of the list box matches the corresponding column of the array.
    

