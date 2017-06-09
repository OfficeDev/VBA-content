---
title: Application.ExportXML Method (Access)
keywords: vbaac10.chm12609
f1_keywords:
- vbaac10.chm12609
ms.prod: access
api_name:
- Access.Application.ExportXML
ms.assetid: 47627677-d311-c2e1-7532-e8a8a9beef29
ms.date: 06/08/2017
---


# Application.ExportXML Method (Access)

The  **ExportXML** method allows developers to export XML data, schemas, and presentation information from Microsoft SQL Server 2000 Desktop Engine (MSDE 2000), Microsoft SQL Server 6.5 or later, or the Microsoft Access database engine.


## Syntax

 _expression_. **ExportXML**( ** _ObjectType_**, ** _DataSource_**, ** _DataTarget_**, ** _SchemaTarget_**, ** _PresentationTarget_**, ** _ImageTarget_**, ** _Encoding_**, ** _OtherFlags_**, ** _WhereCondition_**, ** _AdditionalData_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ObjectType_|Required|**AcExportXMLObjectType**| A **[AcExportXMLObjectType](acexportxmlobjecttype-enumeration-access.md)** that represents the type of **[AccessObject](accessobject-object-access.md)** object to export.|
| _DataSource_|Required|**String**| The name of the **AccessObject** object to export. The default is the currently open object of the type specified by the _ObjectType_ argument.|
| _DataTarget_|Optional|**String**|The file name and path for the exported data. If this argument is omitted, data is not exported.|
| _SchemaTarget_|Optional|**String**|The file name and path for the exported schema information. If this argument is omitted, schema information is not exported to a separate XML file.|
| _PresentationTarget_|Optional|**String**|The file name and path for the exported presentation information. If this argument is omitted, presentation information is not exported.|
| _ImageTarget_|Optional|**String**|The path for exported images. If this argument is omitted, images are not exported.|
| _Encoding_|Optional|**AcExportXMLEncoding**|A  **[AcExportXMLEncoding](acexportxmlencoding-enumeration-access.md)** constant that specifies the text encoding to use for the exported XML. The default value is **acUTF8**.|
| _OtherFlags_|Optional|**AcExportXMLOtherFlags**|A bit mask that specifies other behaviors associated with exporting to XML. Can be a combination of  **[AcExportXMLOtherFlags](acexportxmlotherflags-enumeration-access.md)** constants.|
| _WhereCondition_|Optional|**String**|Specifies a subset of records to be exported.|
| _AdditionalData_|Optional|**Variant**|Specifies additional tables to export. This argument is ignored if the  _OtherFlags_ argument is set to **acLiveReportSource**.|

### Return Value

Nothing


## Remarks

Although the  _DataTarget_,  _SchemaTarget_, and  _PresentationTarget_ arguments are all optional, at least one must be specified when you are using this method. When the **ExportXML** method is called from within an **AccessObject** object, the default behavior is to overwrite any existing files specified in any of the arguments.


## Example

The following example exports the contents of the Customers table in the Northwind Traders sample database, along with the contents of the Orders and Orders Details tables, to an XML data file named Customer Orders.xml.


```vb
Sub ExportCustomerOrderData() 
 Dim objOrderInfo As AdditionalData 
 Dim objOrderDetailsInfo As AdditionalData 
 
 Set objOrderInfo = Application.CreateAdditionalData 
 
 ' Add the Orders and Order Details tables to the data to be exported. 
 Set objOrderDetailsInfo = objOrderInfo.Add("Orders") 
 objOrderDetailsInfo.Add "Order Details" 
 
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

