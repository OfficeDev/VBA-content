---
title: FROM Clause (Microsoft Access SQL)
ms.prod: access
ms.assetid: f3c5931e-2768-198e-d69c-095a01c23bb5
ms.date: 06/08/2017
---


# FROM Clause (Microsoft Access SQL)

**Applies to:** Access 2013 | Access 2016

 **In this article**
[Syntax](#syntax)
[Remarks](#remarks)
[Example](#example)

Specifies the tables or queries that contain the fields listed in the [SELECT](http://msdn.microsoft.com/library/a5c9da94-5f9e-0fc0-767a-4117f38a5ef3%28Office.15%29.aspx) statement.

## Syntax
SELECT  _fieldlist_ FROM _tableexpression_ [IN _externaldatabase_ ]

A SELECT statement containing a FROM clause has these parts:



|**Part**|**Description**|
|:-----|:-----|
| _fieldlist_|The name of the field or fields to be retrieved along with any field-name aliases, [SQL aggregate functions](http://msdn.microsoft.com/library/8866cd71-0216-25b4-6a6a-02cb7acad9a2%28Office.15%29.aspx), selection predicates ([ALL, DISTINCT, DISTINCTROW, or TOP](all-distinct-distinctrow-top-predicates-microsoft-access-sql.md)), or other SELECT statement options.|
| _tableexpression_|An expression that identifies one or more tables from which data is retrieved. The expression can be a single table name, a saved query name, or a compound resulting from an [INNER JOIN](http://msdn.microsoft.com/library/8d16c74c-02c6-12b7-b180-3e7744ef65f3%28Office.15%29.aspx), [LEFT JOIN,](http://msdn.microsoft.com/library/9c10525f-98b1-fd4f-8b40-07a32c5c6502%28Office.15%29.aspx) or [RIGHT JOIN](http://msdn.microsoft.com/library/9c10525f-98b1-fd4f-8b40-07a32c5c6502%28Office.15%29.aspx).|
| _externaldatabase_|The full path of an external database containing all the tables in  _tableexpression._|

## Remarks

FROM is required and follows any SELECT statement.

The order of the table names in  _tableexpression_ is not important.

For improved performance and ease of use, it is recommended that you use a linked table instead of an IN clause to retrieve data from an external database.

The following example shows how you can retrieve data from the Employees table:


```sql
SELECT LastName, FirstName 
FROM Employees;
```


## Example
Some of the following examples assume the existence of a hypothetical Salary field in an Employees table. Note that this field does not actually exist in the Northwind database Employees table.

This example creates a dynaset-type  **Recordset** based on an SQL statement that selects the LastName and FirstName fields of all records in the Employees table. It calls the EnumFields procedure, which prints the contents of a **Recordset** object to the **Debug** window.



```vb
Sub SelectX1() 
 
    Dim dbs As Database, rst As Recordset 
 
    ' Modify this line to include the path to Northwind 
    ' on your computer. 
    Set dbs = OpenDatabase("Northwind.mdb") 
 
    ' Select the last name and first name values of all  
    ' records in the Employees table. 
    Set rst = dbs.OpenRecordset("SELECT LastName, " _ 
        &; "FirstName FROM Employees;") 
 
    ' Populate the recordset. 
    rst.MoveLast 
 
    ' Call EnumFields to print the contents of the 
    ' Recordset. 
    EnumFields rst,12 
 
    dbs.Close 
 
End Sub
```

This example counts the number of records that have an entry in the PostalCode field and names the returned field Tally.



```vb
Sub SelectX2() 
 
    Dim dbs As Database, rst As Recordset 
 
    ' Modify this line to include the path to Northwind 
    ' on your computer. 
    Set dbs = OpenDatabase("Northwind.mdb") 
 
    ' Count the number of records with a PostalCode  
    ' value and return the total in the Tally field. 
    Set rst = dbs.OpenRecordset("SELECT Count " _ 
        &; "(PostalCode) AS Tally FROM Customers;") 
 
    ' Populate the Recordset. 
    rst.MoveLast 
 
    ' Call EnumFields to print the contents of  
    ' the Recordset. Specify field width = 12. 
    EnumFields rst, 12 
 
    dbs.Close 
 
End Sub
```

This example shows the number of employees and the average and maximum salaries.

```vb
Sub SelectX3() 
 
    Dim dbs As Database, rst As Recordset 
 
    ' Modify this line to include the path to Northwind 
    ' on your computer. 
    Set dbs = OpenDatabase("Northwind.mdb") 
 
    ' Count the number of employees, calculate the  
    ' average salary, and return the highest salary. 
    Set rst = dbs.OpenRecordset("SELECT Count (*) " _ 
        &; "AS TotalEmployees, Avg(Salary) " _ 
        &; "AS AverageSalary, Max(Salary) " _ 
        &; "AS MaximumSalary FROM Employees;") 
 
    ' Populate the Recordset. 
    rst.MoveLast 
 
    ' Call EnumFields to print the contents of 
    ' the Recordset. Pass the Recordset object and 
    ' desired field width. 
    EnumFields rst, 17 
 
    dbs.Close 
 
End Sub
```

The **Sub** procedure EnumFields is passed a **Recordset** object from the calling procedure. The procedure then formats and prints the fields of the **Recordset** to the **Debug** window. The variable is the desired printed field width. Some fields may be truncated.



```vb
Sub EnumFields(rst As Recordset, intFldLen As Integer) 
 
    Dim lngRecords As Long, lngFields As Long 
    Dim lngRecCount As Long, lngFldCount As Long 
    Dim strTitle As String, strTemp As String 
 
    ' Set the lngRecords variable to the number of 
    ' records in the Recordset. 
    lngRecords = rst.RecordCount 
 
    ' Set the lngFields variable to the number of 
    ' fields in the Recordset. 
    lngFields = rst.Fields.Count 
 
    Debug.Print "There are " &; lngRecords _ 
        &; " records containing " &; lngFields _ 
        &; " fields in the recordset." 
    Debug.Print 
 
    ' Form a string to print the column heading. 
    strTitle = "Record  " 
    For lngFldCount = 0 To lngFields - 1 
        strTitle = strTitle _ 
        &; Left(rst.Fields(lngFldCount).Name _ 
        &; Space(intFldLen), intFldLen) 
    Next lngFldCount     
 
    ' Print the column heading. 
    Debug.Print strTitle 
    Debug.Print 
 
    ' Loop through the Recordset; print the record 
    ' number and field values. 
    rst.MoveFirst 
 
    For lngRecCount = 0 To lngRecords - 1 
        Debug.Print Right(Space(6) &; _ 
            Str(lngRecCount), 6) &; "  "; 
 
        For lngFldCount = 0 To lngFields - 1 
            ' Check for Null values. 
            If IsNull(rst.Fields(lngFldCount)) Then 
                strTemp = "<null>" 
            Else 
                ' Set strTemp to the field contents.  
                Select Case _ 
                    rst.Fields(lngFldCount).Type 
                    Case 11 
                        strTemp = "" 
                    Case dbText, dbMemo 
                        strTemp = _ 
                            rst.Fields(lngFldCount) 
                    Case Else 
                        strTemp = _ 
                            str(rst.Fields(lngFldCount)) 
                End Select 
            End If 
 
            Debug.Print Left(strTemp _  
                &; Space(intFldLen), intFldLen); 
        Next lngFldCount 
 
        Debug.Print 
 
        rst.MoveNext 
 
    Next lngRecCount 
 
End Sub
```

## See also
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)
[Access help on support.office.com](https://support.office.com/search/results?query=Access)
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)
[Search for specific Access error codes on Bing](http://www.bing.com/)
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)
