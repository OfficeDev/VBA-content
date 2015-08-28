
# OfficeDataSourceObject.ApplyFilter Method (Office)

 **Last modified:** July 28, 2015

Applies a filter to a mail merge data source to filter specified records meeting specified criteria.

## Syntax

 _expression_. **ApplyFilter**

 _expression_A variable that represents an  **OfficeDataSourceObject** object.


## Example

This example adds a new filter that removes all records with a blank Region field and then applies the filter to the active publication.


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


 [OfficeDataSourceObject Object](d5e5401b-643e-c12c-2648-f281af481f45.md)
#### Other resources


 [OfficeDataSourceObject Object Members](57ba0dc6-80e7-04a9-a619-2a3e6aa2cdff.md)
