---
title: OfficeDataSourceObject Object (Office)
keywords: vbaof11.chm232000
f1_keywords:
- vbaof11.chm232000
ms.prod: office
api_name:
- Office.OfficeDataSourceObject
ms.assetid: d5e5401b-643e-c12c-2648-f281af481f45
ms.date: 06/08/2017
---


# OfficeDataSourceObject Object (Office)

Represents the mail merge data source in a mail merge operation.


## Remarks

To work with the  **OfficeDataSourceObject** object, dimension a variable as an **OfficeDataSourceObject** object. You can then work with the different properties and methods associated with the object. Use the **SetSortOrder** method to specify how to sort the records in a data source.


## Example

 The following example sorts the data source first according to Postal Code in descending order, then on last name and first name in ascending order.


```
Sub SetDataSortOrder() 
 Dim appOffice As OfficeDataSourceObject 
 
 Set appOffice = Application.OfficeDataSourceObject 
 appOffice.Open bstrConnect:="DRIVER=SQL Server;SERVER=ServerName;" &amp; _ 
 "UID=user;PWD=;DATABASE=Northwind", bstrTable:="Employees" 
 
 appOffice.SetSortOrder SortField1:="ZipCode", _ 
 SortAscending1:=False, SortField2:="LastName", _ 
 SortField3:="FirstName" 
End Sub
```

Use the  **Column**, **Comparison**, **CompareTo**, and **Conjunction** properties to return or set the data source query criterion. The following example changes an existing filter to remove from the mail merge all records that do not have a Region field equal to "WA".




```
Sub SetQueryCriterion() 
 Dim appOffice As Office.OfficeDataSourceObject 
 Dim intItem As Integer 
 
 appOffice.Open bstrConnect:="DRIVER=SQL Server;SERVER=ServerName;" &amp; _ 
 "UID=user;PWD=;DATABASE=Northwind", bstrTable:="Employees" 
 
 With appOffice.Filters 
 For intItem = 1 To .Count 
 With .Item(intItem) 
 If .Column = "Region" Then 
 .Comparison = msoFilterComparisonNotEqual 
 .CompareTo = "WA" 
 If .Conjunction = "Or" Then .Conjunction = "And" 
 End If 
 End With 
 Next intItem 
 End With 
End Sub
```


## Methods



|**Name**|
|:-----|
|[ApplyFilter](officedatasourceobject-applyfilter-method-office.md)|
|[Move](officedatasourceobject-move-method-office.md)|
|[Open](officedatasourceobject-open-method-office.md)|
|[SetSortOrder](officedatasourceobject-setsortorder-method-office.md)|

## Properties



|**Name**|
|:-----|
|[Columns](officedatasourceobject-columns-property-office.md)|
|[ConnectString](officedatasourceobject-connectstring-property-office.md)|
|[DataSource](officedatasourceobject-datasource-property-office.md)|
|[Filters](officedatasourceobject-filters-property-office.md)|
|[RowCount](officedatasourceobject-rowcount-property-office.md)|
|[Table](officedatasourceobject-table-property-office.md)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
