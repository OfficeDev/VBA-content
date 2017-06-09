---
title: ODSOFilter.Comparison Property (Office)
keywords: vbaof11.chm240004
f1_keywords:
- vbaof11.chm240004
ms.prod: office
api_name:
- Office.ODSOFilter.Comparison
ms.assetid: 992565b3-90c5-4f44-7cae-ba0533529127
ms.date: 06/08/2017
---


# ODSOFilter.Comparison Property (Office)

Gets or sets an  **MsoFilterComparison** constant that represents how to compare the **Column** and **CompareTo** properties. Read/write.


## Syntax

 _expression_. **Comparison**

 _expression_ A variable that represents an **ODSOFilter** object.


## Example

The following example changes an existing filter to remove from the mail merge all records that do not have a  **Region** field equal to "WA".


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


## See also


#### Concepts


[ODSOFilter Object](odsofilter-object-office.md)
#### Other resources


[ODSOFilter Object Members](odsofilter-members-office.md)

