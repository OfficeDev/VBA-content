---
title: ODSOFilters Object (Office)
keywords: vbaof11.chm241000
f1_keywords:
- vbaof11.chm241000
ms.prod: office
api_name:
- Office.ODSOFilters
ms.assetid: e706745d-3890-81e8-6c9a-4c6bf67387ee
ms.date: 06/08/2017
---


# ODSOFilters Object (Office)

Represents all the filters to apply to the data source attached to the mail merge publication. The  **ODSOFilters** object is composed of **ODSOFilter** objects.


## Remarks

Use the  **Add** method of the **ODSOFilters** object to add a new filter criterion to the query.


## Example

This example adds a new line to the query string and then applies the combined filter to the data source.


```
Sub SetQueryCriterion() 
 Dim appOffice As OfficeDataSourceObject 
 
 Set appOffice = Application.OfficeDataSourceObject 
 appOffice.Open bstrConnect:="DRIVER=SQL Server;SERVER=ServerName;" &amp; _ 
 "UID=user;PWD=;DATABASE=Northwind", bstrTable:="Employees" 
 
 With appOffice.Filters 
 .Add Column:="Region", _ 
 Comparison:=msoFilterComparisonIsBlank, _ 
 Conjunction:=msoFilterConjunctionAnd 
 .ApplyFilter 
 End With 
End Sub
```

Use the  **Item** method to access an individual filter criterion. This example loops through all the filter criterion and if it finds one with a value of "Region", changes it to remove from the mail merge all records that are not equal to "WA".




```
Sub SetQueryCriterion() 
 Dim appOffice As Office.OfficeDataSourceObject 
 Dim intItem As Integer 
 
 Set appOffice = Application.OfficeDataSourceObject 
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
|[Add](odsofilters-add-method-office.md)|
|[Delete](odsofilters-delete-method-office.md)|
|[Item](odsofilters-item-method-office.md)|

## Properties



|**Name**|
|:-----|
|[Application](odsofilters-application-property-office.md)|
|[Count](odsofilters-count-property-office.md)|
|[Creator](odsofilters-creator-property-office.md)|
|[Parent](odsofilters-parent-property-office.md)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
