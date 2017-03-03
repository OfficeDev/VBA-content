---
title: ODSOColumns Object (Office)
keywords: vbaof11.chm234000
f1_keywords:
- vbaof11.chm234000
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.ODSOColumns
ms.assetid: eaac6cd2-45ff-72ea-c9c9-a22f24214756
---


# ODSOColumns Object (Office)

A collection of  **ODSOColumn** objects that represent the data fields in a mail merge data source.


## Example

Use the  **Columns** property to return the **ODSOColumns** collection. The following example displays the field names in the data source attached to the active publication.


```vb
Sub ShowFieldNames() 
 Dim appOffice As OfficeDataSourceObject 
 Dim intCount As Integer 
 
 Set appOffice = Application.OfficeDataSourceObject 
 appOffice.Open bstrConnect:="DRIVER=SQL Server;SERVER=ServerName;" &; _ 
 "UID=user;PWD=;DATABASE=Northwind", bstrTable:="Employees" 
 
 With appOffice.Columns 
 For intCount = 1 To .Count 
 MsgBox "Column Name: " &; .Item(intCount).Name 
 Next 
 End With 
End Sub
```

Use ** Columns** ( _index_ ), where _index_ is the data field name or the index number, to return a single **ODSOColumn** object. The index number represents the position of the data field in the mail merge data source. This example retrieves the name of the first field and value of the first record of the FirstName field in the data source attached to the active publication.




```vb
Sub GetDataFromSource() 
 Dim appOffice As OfficeDataSourceObject 
 
 Set appOffice = Application.OfficeDataSourceObject 
 appOffice.Open bstrConnect:="DRIVER=SQL Server;SERVER=ServerName;" &; _ 
 "UID=user;PWD=;DATABASE=Northwind", bstrTable:="Employees" 
 
 With appOffice.Columns 
 MsgBox "Field Name: " &; .Columns(1).Name &; _ 
 "Value: " &; .Columns("FirstName").Value 
 End With 
End Sub
```


## See also


#### Concepts


[Object Model Reference](../../Office-Shared-VBA/articles/reference-object-library-reference-for-office.md)

