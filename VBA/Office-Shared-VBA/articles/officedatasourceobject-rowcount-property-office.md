---
title: OfficeDataSourceObject.RowCount Property (Office)
keywords: vbaof11.chm232005
f1_keywords:
- vbaof11.chm232005
ms.prod: office
api_name:
- Office.OfficeDataSourceObject.RowCount
ms.assetid: 5360a399-e2f8-b331-f62c-c110884b3c92
ms.date: 06/08/2017
---


# OfficeDataSourceObject.RowCount Property (Office)

Gets a  **Long** that represents the number of records in the specified data source. Read-only.


## Syntax

 _expression_. **RowCount**

 _expression_ A variable that represents an **OfficeDataSourceObject** object.


## Example

This example adds a new filter that removes all records with a blank  **Region** field and then applies the filter to the active publication.


```
Sub OfficeFilters() 
 Dim appOffice As OfficeDataSourceObject 
 Dim appFilters As ODSOFilters 
 
 Set appOffice = Application.OfficeDataSourceObject 
 appOffice.Open bstrConnect:="DRIVER=SQL Server;SERVER=ServerName;" &amp; _ 
 "UID=user;PWD=;DATABASE=Northwind", bstrTable:="Employees" 
 
 Set appFilters = appOffice.Filters 
 
 MsgBox appOffice.RowCount 
 
 appFilters.Add Column:="Region", Comparison:=msoFilterComparisonEqual, _ 
 Conjunction:=msoFilterConjunctionAnd, bstrCompareTo:="WA" 
 appOffice.ApplyFilter 
 
 MsgBox appOffice.RowCount 
 
End Sub
```


## See also


#### Concepts


[OfficeDataSourceObject Object](officedatasourceobject-object-office.md)
#### Other resources


[OfficeDataSourceObject Object Members](officedatasourceobject-members-office.md)

