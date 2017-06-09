---
title: ODSOColumn Object (Office)
keywords: vbaof11.chm233000
f1_keywords:
- vbaof11.chm233000
ms.prod: office
api_name:
- Office.ODSOColumn
ms.assetid: f8fe41bd-c9bd-fb5b-8ca7-27940c9c0996
ms.date: 06/08/2017
---


# ODSOColumn Object (Office)

Represents a field in a data source. The  **ODSOColumn** object is a member of the **ODSOColumns** collection.


## Remarks

The  **ODSOColumns** collection includes all the data fields in a mail merge data source (for example, Name, Address, and City).

You cannot add fields to the  **ODSOColumns** collection. All data fields in a data source are automatically included in the **ODSOColumns** collection.

Use [Columns](officedatasourceobject-columns-property-office.md)( _index_ ), where _index_ is the data field name or index number, to return a single **ODSOColumn** object. The index number represents the position of the data field in the mail merge data source.


## Example

This example retrieves the name and value of the first field of the first record in the data source attached to the active publication.


```
Sub GetDataFromSource() 
 
 Dim appOffice As OfficeDataSourceObject 
 
 Set appOffice = Application.OfficeDataSourceObject 
 appOffice.Open bstrConnect:="DRIVER=SQL Server;SERVER=ServerName;" &amp; _ 
 "UID=user;PWD=;DATABASE=Northwind", bstrTable:="Employees" 
 
 With appOffice.Columns 
 MsgBox "Field Name: " &amp; .Item(1).Name &amp; vbLf &amp; _ 
 "Value: " &amp; .Item(1).Value 
 End With 
End Sub
```


## Properties



|**Name**|
|:-----|
|[Application](odsocolumn-application-property-office.md)|
|[Creator](odsocolumn-creator-property-office.md)|
|[Index](odsocolumn-index-property-office.md)|
|[Name](odsocolumn-name-property-office.md)|
|[Parent](odsocolumn-parent-property-office.md)|
|[Value](odsocolumn-value-property-office.md)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
