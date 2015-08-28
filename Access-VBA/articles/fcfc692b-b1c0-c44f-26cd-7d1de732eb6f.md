
# Synchronize Two Combo Boxes on a Form

 **Last modified:** July 28, 2015

 _**Applies to:** Access 2013_

This topic illustrates how to synchronize two combo boxes so that when you select an item in the first combo box, the selection limits the choices in the second combo box. For example, you may want the products displayed in a combo box to be limited to the category selected in another combo box.

In this example, the second combo box is filled with the results of an SQL statement. This SQL statement finds all the products that have a CategoryID that matches the category selected in the first combo box.
Whenever a category is selected in the first combo box, its  ** [AfterUpdate](89b45f0c-5ab1-889e-bd26-a34281b49b9e.md)** event procedure sets the second combo box's ** [RowSourceType](dd1d6ea8-5479-4bf9-3317-0b95282c7d74.md)** property. This refreshes the list of available products in the second combo box. Without this procedure, the contents of the second combo box would not change.



```
Private Sub cboCategories_AfterUpdate() 
 
    ' Update the row source of the cboProducts combo box 
    ' when the user makes a selection in the cboCategories 
    ' combo box. 
    Me.cboProducts.RowSource = "SELECT ProductName FROM" &amp; _ 
                            " tblProducts WHERE CategoryID = " &amp; Me.cboCategories &amp; _ 
                            " ORDER BY ProductName" 
                             
    Me.cboProducts = Me.cboProducts.ItemData(0) 
End Sub
```

