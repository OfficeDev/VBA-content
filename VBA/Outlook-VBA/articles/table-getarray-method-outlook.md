---
title: Table.GetArray Method (Outlook)
keywords: vbaol11.chm2230
f1_keywords:
- vbaol11.chm2230
ms.prod: outlook
api_name:
- Outlook.Table.GetArray
ms.assetid: 2594bb2e-290f-8e88-52d1-cd2b2191bbe3
ms.date: 06/08/2017
---


# Table.GetArray Method (Outlook)

Obtains a two-dimensional array that contains a set of row and column values from the  **[Table](table-object-outlook.md)** .


## Syntax

 _expression_ . **GetArray**( **_MaxRows_** )

 _expression_ A variable that represents a **Table** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _MaxRows_|Required| **Long**|Specifies the maximum number of rows to return from the  **Table** .|

### Return Value

A  **Variant** value that is a two-dimensional array representing a set of row and column values from the **Table** . The array is zero-based; an array index (i, j) indexes into the i-th column and j-th row in the array. Columns in the array correspond to columns in the **Table** , and rows in the array correspond to rows in the **Table** . The number of rows in the returned array is the lesser value of _MaxRows_ and the actual number of rows in the **Table**.


## Remarks

The  **GetArray** method offers a conceptually simple means to get values from a **Table** by copying all or part of the data in the **Table** (based on the current row) to an array and indexing into the array.

 **GetArray** always starts at the current row of the **Table** . It returns an array with _MaxRows_ number of rows if there are at least _MaxRows_ number of rows in the **Table** starting at the current position. If _MaxRows_ is not larger than the total number of rows in the **Table** , and there are fewer than _MaxRows_ number of elements in the **Table** starting at the current row, it will return an array that contains only the remaining rows in the **Table** . If **GetArray** is called and there are no remaining rows, then **GetArray** returns an empty array with zero elements.

After obtaining the appropriate rows from the  **Table** and before it returns, **GetArray** always repositions the current row to the next row in the **Table** , if there exists a next row. `GetArray(n)` operates as if **[Table.GetNextRow](table-getnextrow-method-outlook.md)** is called n times.

The values in the columns map to columns in the  **Table** , and are therefore determined by the format of the property name used for the column. For more information, see[Factors Affecting Property Value Representation in the Table and View Classes](http://msdn.microsoft.com/library/13cf9945-a9e0-bb32-a2cb-74366a365ae1%28Office.15%29.aspx).


## Example

The following code sample obtains a  **Table** by filtering on all items in the Inbox that contain "Office" in the subject. It then uses the **Table.GetArray** method to copy the data from the **Table** to an array, and prints the property value of each item returned.

For more information on specifying property names in a filter by namespace reference, see [Referencing Properties by Namespace](http://msdn.microsoft.com/library/c1c7bfa9-64d7-81d2-84e7-f0a4c57780b3%28Office.15%29.aspx).




```vb
Sub DemoTableUsingGetArray() 
 'Declarations 
 Dim Filter As String 
 Dim i, ubRows As Long 
 Dim j, ubCols As Integer 
 Dim varArray 
 Dim oTable As Outlook.Table 
 Dim oFolder As Outlook.Folder 
 Const SchemaPropTag As String = _ 
 "http://schemas.microsoft.com/mapi/proptag/" 
 
 On Error Resume Next 
 'Get a Folder object for the Inbox 
 Set oFolder = Application.Session.GetDefaultFolder(olFolderInbox) 
 'Filter on the subject containing "Office" 
 Filter = "@SQL=" &; Chr(34) &; SchemaPropTag &; "0x0037001E" _ 
 &; Chr(34) &; " like '%Office%'" 
 'Get all items in Inbox that meet the filter 
 Set oTable = oFolder.GetTable(Filter) 
 
 On Error GoTo Err_Trap 
 varArray = oTable.GetArray(oTable.GetRowCount) 
 
 'Number of rows is the second dimension of the array 
 ubRows = UBound(varArray, 2) 
 'Number of columns is the first dimension of the array 
 ubCols = UBound(varArray) 
 
 'Array is zero-based 
 'Rows corrspond to items in the table, so for each item... 
 For j = 0 To ubRows 
 'Columns correspond to properties in the table, print the value of each property 
 For i = 0 To ubCols 
 Debug.Print varArray(i, j) 
 Next 
 Next 
 Exit Sub 
 
Err_Trap: 
 Debug.Print "Error#:" &; Err.Number &; " Desc: " &; Err.Description 
 Resume Next 
End Sub
```


## See also


#### Concepts


[Table Object](table-object-outlook.md)

