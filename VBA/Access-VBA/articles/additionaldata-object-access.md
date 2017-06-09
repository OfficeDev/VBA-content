---
title: AdditionalData Object (Access)
keywords: vbaac10.chm13253
f1_keywords:
- vbaac10.chm13253
ms.prod: access
api_name:
- Access.AdditionalData
ms.assetid: 2677072b-c2ca-3bcd-fef4-f6b1cadb0379
ms.date: 06/08/2017
---


# AdditionalData Object (Access)

Represents the collection of tables and queries that will be included with the parent table that is exported by the  **[ExportXML](application-exportxml-method-access.md)** method.


## Remarks

To create an  **AdditionalData** object, use the **[CreateAdditionalData](application-createadditionaldata-method-access.md)** method of the **[Application](application-object-access.md)** object.

To add a table to an existing  **AdditionalData** object, use the **Add** method.


## Example

The following example exports the contents of the Customers table in the Northwind Traders sample database, along with the contents of the Orders and Orders Details tables, to an XML data file named Customer Orders.xml.


```
Sub ExportCustomerOrderData() 
 Dim objOrderInfo As AdditionalData 
 
 Set objOrderInfo = Application.CreateAdditionalData 
 
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


## Methods



|**Name**|
|:-----|
|[Add](additionaldata-add-method-access.md)|

## Properties



|**Name**|
|:-----|
|[Count](additionaldata-count-property-access.md)|
|[Item](additionaldata-item-property-access.md)|
|[Name](additionaldata-name-property-access.md)|

## See also


#### Other resources


[Access Object Model Reference](http://msdn.microsoft.com/library/2de134a4-6c5c-d2a3-8377-f4dd973ba650%28Office.15%29.aspx)
