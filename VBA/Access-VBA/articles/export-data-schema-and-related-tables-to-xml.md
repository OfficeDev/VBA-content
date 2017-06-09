---
title: Export Data, Schema, and Related Tables to XML
ms.prod: access
ms.assetid: 4f84813a-bc39-ac03-f04f-624f74eed190
ms.date: 06/08/2017
---


# Export Data, Schema, and Related Tables to XML

The  **[ExportXML](application-exportxml-method-access.md)** method can be used to export the data and formatting contained in a table, along with any additonal data that you specify.

To specify the additional data to export, you must must use the  **[CreateAdditionalData](application-createadditionaldata-method-access.md)** method to create an **[AdditionalData](additionaldata-object-access.md)** object. Then, use the **[Add](additionaldata-add-method-access.md)** method to add additonal tables to export along with the main table.

The following procedure illustrates how to include additional data when exporting a table to XML. The Orders table is exported along with several other tables. The schema and the formatting are also exported as separate .xsd and .xsl files, respectively.




```vb
Private Sub ExportRelTables() 
   ' Purpose: Exports the Orders table as well as  
   ' a number of related databases to an XML file. 
   ' XSD and XSL files are also created. 
 
   Dim objAD As AdditionalData 
 
   ' Create the AdditionalData object. 
   Set objAD = Application.CreateAdditionalData 
 
   ' Add the related tables to the object. 
   With objAD 
      .Add "Order Details" 
      objAD(Item:="Order Details").Add "Order Details Details" 
      .Add "Customers" 
      .Add "Shippers" 
      .Add "Employees" 
      .Add "Products" 
      objAD(Item:="Products").Add "Product Details" 
      objAD(Item:="Products")(Item:="Product Details").Add _ 
         "Product Details Details" 
      .Add "Suppliers" 
      .Add "Categories" 
   End With 
 
   ' Export the Orders table along with the addtional data. 
   Application.ExportXml acExportTable, "Orders", _ 
       "C:\Orders.xml", "C:\OrdersSchema.xsd", _ 
       "C:\OrdersStyle.xsl", AdditionalData:= objAD 
End Sub
```


