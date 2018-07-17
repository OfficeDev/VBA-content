---
title: Assign a Control a Value From a Table
ms.prod: access
ms.assetid: d9bba7e3-bca0-00df-3753-dc99ae767759
ms.date: 06/08/2017
---


# Assign a Control a Value From a Table

You can use the  **DLookup** function to display the value of a field that is not in the record source for your form or report. For example, suppose you have a form based on an Order Details table. The form displays the OrderID, ProductID, UnitPrice, Quantity, and Discount fields. However, the UnitPrice field is in another table: Products. You could use the **DLookup** function in a calculated control to display the UnitPrice on the same form when the user selects a product.

The following example populates the UnitPrice text box with the price of the product currently selected in the ProductID combo box.



```vb
Private Sub ProductID_AfterUpdate() 
 
 ' Evaluate filter before it is passed to DLookup function. 
 strFilter = "ProductID = " &; Me!ProductID 
 
 ' Look up product's unit price and assign it to the UnitPrice control. 
 Me!UnitPrice = DLookup("UnitPrice", "Products", strFilter) 
 
End Sub
```

The  **DLookup** function has three arguments. The first specifies the field you are looking up (UnitPrice); the second specifies the table (Products); and the third specifies which value to find (the value for the record where the ProductID is the same as the ProductID on the current record in the Orders subform).

