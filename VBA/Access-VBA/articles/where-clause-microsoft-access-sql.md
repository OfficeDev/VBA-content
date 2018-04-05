---
title: WHERE Clause (Microsoft Access SQL)
keywords: jetsql40.chm5277568
f1_keywords:
- jetsql40.chm5277568
ms.prod: access
ms.assetid: 67e4caed-6512-e8bd-39d0-6dca18114b18
ms.date: 06/08/2017
---


# WHERE Clause (Microsoft Access SQL)

  

**Applies to:** Access 2013 | Access 2016

 **In this article**
[Syntax](#syntax)
[Remarks](#remarks)
[Example](#example)


Specifies which records from the tables listed in the [FROM](from-clause-microsoft-access-sql.md) clause are affected by a [SELECT](http://msdn.microsoft.com/library/a5c9da94-5f9e-0fc0-767a-4117f38a5ef3%28Office.15%29.aspx), [UPDATE](http://msdn.microsoft.com/library/08f9c3d6-c020-ecf1-5748-43b93a76dfbb%28Office.15%29.aspx), or [DELETE](http://msdn.microsoft.com/library/64c235bc-5b1a-0a33-714a-9933ba7a81e5%28Office.15%29.aspx) statement.

## Syntax
SELECT  _fieldlist_ FROM _tableexpression_ WHERE _criteria_

A SELECT statement containing a WHERE clause has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _fieldlist_|The name of the field or fields to be retrieved along with any field-name aliases, selection predicates ([ALL, DISTINCT, DISTINCTROW, or TOP](all-distinct-distinctrow-top-predicates-microsoft-access-sql.md)), or other SELECT statement options.|
| _tableexpression_|The name of the table or tables from which data is retrieved.|
| _criteria_|An expression that records must satisfy to be included in the query results.|

## Remarks

The Access database engine selects the records that meet the conditions listed in the WHERE clause. If you do not specify a WHERE clause, your query returns all rows from the table. If you specify more than one table in your query and you have not included a WHERE clause or a JOIN clause, your query generates a Cartesian product of the tables.

WHERE is optional, but when included, follows FROM. For example, you can select all employees in the sales department ( `WHERE Dept = 'Sales'`) or all customers between the ages of 18 and 30 () or all customers between the ages of 18 and 30 ( `WHERE Age Between 18 And 30`).

If you do not use a JOIN clause to perform SQL join operations on multiple tables, the resulting **Recordset** object will not be updatable.

WHERE is similar to [HAVING](having-clause-microsoft-access-sql.md). WHERE determines which records are selected. Similarly, once records are grouped with [GROUP BY](group-by-clause-microsoft-access-sql.md), HAVING determines which records are displayed.

Use the WHERE clause to eliminate records you do not want grouped by a GROUP BY clause.

Use various expressions to determine which records the SQL statement returns. For example, the following SQL statement selects all employees whose salaries are more than $21,000:



```sql
SELECT LastName, Salary 
FROM Employees 
WHERE Salary > 21000;
```

A WHERE clause can contain up to 40 expressions linked by logical operators, such as **And** and **Or**.

When you enter a field name that contains a space or punctuation, surround the name with brackets ([ ]). For example, a customer information table might include information about specific customers:



```
SELECT [Customer's Favorite Restaurant]
```

When you specify the  _criteria_ argument, date literals must be in U.S. format, even if you are not using the U.S. version of the Microsoft® Jet database engine. For example, May 10, 1996, is written 10/5/96 in the United Kingdom and 5/10/96 in the United States. Be sure to enclose your date literals with the number sign (#) as shown in the following examples.

To find records dated May 10, 1996 in a United Kingdom database, you must use the following SQL statement:



```sql
SELECT * 
FROM Orders 
WHERE ShippedDate = #5/10/96#;
```

You can also use the  **DateValue** function which is aware of the international settings established by Microsoft Windows®. For example, use this code for the United States:



```sql
SELECT * 
FROM Orders 
WHERE ShippedDate = DateValue('5/10/96');
```

And use this code for the United Kingdom:




```sql
SELECT * 
FROM Orders 
WHERE ShippedDate = DateValue('10/5/96');
```


>**Note**  If the column referenced in the criteria string is of type GUID, the criteria expression uses a slightly different syntax:


```
WHERE ReplicaID = {GUID {12345678-90AB-CDEF-1234-567890ABCDEF}}
```

Be sure to include the nested braces and hyphens as shown.


## Example

The following example assumes the existence of a hypothetical Salary field in an Employees table. Note that this field does not actually exist in the Northwind database Employees table.

This example selects the LastName and FirstName fields of each record in which the last name is King.

This example calls the EnumFields procedure, which you can find in the SELECT statement example.



```vb
Sub WhereX() 
 
    Dim dbs As Database, rst As Recordset 
 
    ' Modify this line to include the path to Northwind 
    ' on your computer. 
    Set dbs = OpenDatabase("Northwind.mdb") 
 
    ' Select records from the Employees table where the 
    ' last name is King. 
    Set rst = dbs.OpenRecordset("SELECT LastName, " _ 
        &; "FirstName FROM Employees " _ 
        &; "WHERE LastName = 'King';") 
     
    ' Populate the Recordset. 
    rst.MoveLast 
     
    ' Call EnumFields to print the contents of the 
    ' Recordset. 
    EnumFields rst, 12 
 
    dbs.Close 
 
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