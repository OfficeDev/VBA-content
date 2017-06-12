---
title: Application.CreateAdditionalData Method (Access)
keywords: vbaac10.chm12607
f1_keywords:
- vbaac10.chm12607
ms.prod: access
api_name:
- Access.Application.CreateAdditionalData
ms.assetid: d27df827-1bcc-eb1e-00d2-46eebd265440
ms.date: 06/08/2017
---


# Application.CreateAdditionalData Method (Access)

Creates an  **[AdditionalData](additionaldata-object-access.md)** object that can be used to add additional tables and queries to the parent table that is being exported by the **[ExportXML](application-exportxml-method-access.md)** method.


## Syntax

 _expression_. **CreateAdditionalData**

 _expression_ A variable that represents an **Application** object.


### Return Value

AdditionalData


## Example

The following example exports the contents of the Customers table in the Northwind Traders sample database, along with the contents of the Orders and Orders Details tables, to an XML data file named Customer Orders.xml.


```vb
Sub ExportCustomerOrderData() 
 Dim objOrderInfo As AdditionalData 
 
 Set objOrderInfo = Application.CreateAdditionalData() 
 
 ' Add the Orders and Order Details tables to the data to be exported. 
 objOrderInfo.Add "Orders" 
 objOrderInfo.Add "Order Details" 
 
 ' Export the contents of the Customers table. The Orders and Order 
 ' Details tables will be included in the XML file. 
 Application.ExportXML ObjectType:=acExportTable, DataSource:="Customers", _ 
 DataTarget:="Customer Orders.xml", _ 
 AdditionalData:=objOrderInfo 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-access.md)

