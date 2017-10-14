---
title: Synchronize Two Combo Boxes on a Form
ms.prod: access
ms.assetid: fcfc692b-b1c0-c44f-26cd-7d1de732eb6f
ms.date: 06/08/2017
---


# Synchronize Two Combo Boxes on a Form

This topic illustrates how to synchronize two combo boxes so that when you select an item in the first combo box, the selection limits the choices in the second combo box. For example, you may want the products displayed in a combo box to be limited to the category selected in another combo box.

In this example, the second combo box is filled with the results of an SQL statement. This SQL statement finds all the products that have a CategoryID that matches the category selected in the first combo box.

Whenever a category is selected in the first combo box, its  **[AfterUpdate](combobox-afterupdate-event-access.md)** event procedure sets the second combo box's **[RowSourceType](combobox-rowsourcetype-property-access.md)** property. This refreshes the list of available products in the second combo box. Without this procedure, the contents of the second combo box would not change.




```vb
Private Sub cboCategories_AfterUpdate() 
 
    ' Update the row source of the cboProducts combo box 
    ' when the user makes a selection in the cboCategories 
    ' combo box. 
    Me.cboProducts.RowSource = "SELECT ProductName FROM" &; _ 
                            " tblProducts WHERE CategoryID = " &; Me.cboCategories &; _ 
                            " ORDER BY ProductName" 
                             
    Me.cboProducts = Me.cboProducts.ItemData(0) 
End Sub
```


