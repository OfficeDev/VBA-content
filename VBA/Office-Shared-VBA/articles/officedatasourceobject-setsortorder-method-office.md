---
title: OfficeDataSourceObject.SetSortOrder Method (Office)
keywords: vbaof11.chm232008
f1_keywords:
- vbaof11.chm232008
ms.prod: office
api_name:
- Office.OfficeDataSourceObject.SetSortOrder
ms.assetid: 427d3a81-1863-4e52-02d4-7485553a4d2f
ms.date: 06/08/2017
---


# OfficeDataSourceObject.SetSortOrder Method (Office)

Sets the sort order for mail merge data.


## Syntax

 _expression_. **SetSortOrder**( **_SortField1_**, **_SortAscending1_**, **_SortField2_**, **_SortAscending2_**, **_SortField3_**, **_SortAscending3_** )

 _expression_ A variable that represents an **OfficeDataSourceObject** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SortField1_|Required|**String**|The first field on which to sort the mail merge data.|
| _SortAscending1_|Optional|**Boolean**|True (default) to perform an ascending sort on SortField1;  **False** to perform a descending sort.|
| _SortField2_|Optional|**String**|The second field on which to sort the mail merge data. Default is an empty string.|
| _SortAscending2_|Optional|**Boolean**|True (default) to perform an ascending sort on SortField2;  **False** to perform a descending sort.|
| _SortField3_|Optional|**String**|The third field on which to sort the mail merge data. Default is an empty string.|
| _SortAscending3_|Optional|**Boolean**|True (default) to perform an ascending sort on SortField3;  **False** to perform a descending sort.|

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


## See also


#### Concepts


[OfficeDataSourceObject Object](officedatasourceobject-object-office.md)
#### Other resources


[OfficeDataSourceObject Object Members](officedatasourceobject-members-office.md)

