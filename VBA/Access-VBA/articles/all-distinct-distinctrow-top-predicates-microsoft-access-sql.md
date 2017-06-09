---
title: ALL, DISTINCT, DISTINCTROW, TOP Predicates (Microsoft Access SQL)
keywords: jetsql40.chm5277572
f1_keywords:
- jetsql40.chm5277572
ms.prod: access
ms.assetid: 6ff5c418-897b-7d65-8551-5a0ace3c587f
ms.date: 06/08/2017
---


# ALL, DISTINCT, DISTINCTROW, TOP Predicates (Microsoft Access SQL)

  

**Applies to:** Access 2013 | Access 2016

Specifies records selected with SQL queries.


## Syntax

SELECT [ALL | DISTINCT | DISTINCTROW | [TOP  _n_ [PERCENT]]] FROM _table_

A SELECT statement containing these predicates has the following parts:



|**Part**|**Description**|
|:-----|:-----|
|ALL|Assumed if you do not include one of the predicates. The Microsoft Access database engine selects all of the records that meet the conditions in the SQL statement. The following two examples are equivalent and return all records from the Employees table:
```sql
SELECT ALL * 
FROM Employees 
ORDER BY EmployeeID; 

```


```sql
SELECT * 
FROM Employees 
ORDER BY EmployeeID;
```

|
|DISTINCT|Omits records that contain duplicate data in the selected fields. To be included in the results of the query, the values for each field listed in the SELECT statement must be unique. For example, several employees listed in an Employees table may have the same last name. If two records contain Smith in the LastName field, the following SQL statement returns only one record that contains Smith:
```sql
SELECT DISTINCT 
LastName 
FROM Employees;
```

If you omit DISTINCT, this query returns both Smith records.If the SELECT clause contains more than one field, the combination of values from all fields must be unique for a given record to be included in the results.The output of a query that uses DISTINCT is not updatable and does not reflect subsequent changes made by other users.|
|DISTINCTROW|Omits data based on entire duplicate records, not just duplicate fields. For example, you could create a query that joins the Customers and Orders tables on the CustomerID field. The Customers table contains no duplicate CustomerID fields, but the Orders table does because each customer can have many orders. The following SQL statement shows how you can use DISTINCTROW to produce a list of companies that have at least one order but without any details about those orders:
```sql
SELECT DISTINCTROW CompanyName 
FROM Customers INNER JOIN Orders 
ON Customers.CustomerID = Orders.CustomerID 
ORDER BY CompanyName;
```

If you omit DISTINCTROW, this query produces multiple rows for each company that has more than one order.DISTINCTROW has an effect only when you select fields from some, but not all, of the tables used in the query. DISTINCTROW is ignored if your query includes only one table, or if you output fields from all tables.|
|TOP  _n_ [PERCENT]|Returns a certain number of records that fall at the top or the bottom of a range specified by an ORDER BY clause. Suppose you want the names of the top 25 students from the class of 1994:
```sql
SELECT TOP 25 
FirstName, LastName 
FROM Students 
WHERE GraduationYear = 1994 
ORDER BY GradePointAverage DESC;
```

If you do not include the ORDER BY clause, the query will return an arbitrary set of 25 records from the Students table that satisfy the WHERE clause.The TOP predicate does not choose between equal values. In the preceding example, if the twenty-fifth and twenty-sixth highest grade point averages are the same, the query will return 26 records.You can also use the PERCENT reserved word to return a certain percentage of records that fall at the top or the bottom of a range specified by an ORDER BY clause. Suppose that, instead of the top 25 students, you want the bottom 10 percent of the class:


```sql
SELECT TOP 10 PERCENT 
FirstName, LastName 
FROM Students 
WHERE GraduationYear = 1994 
ORDER BY GradePointAverage ASC;
```

The ASC predicate specifies a return of bottom values. The value that follows TOP must be an unsigned  **Integer**.TOP does not affect whether or not the query is updatable.|
| _table_|The name of the table from which records are retrieved.|

## Example

This example creates a query that joins the Customers and Orders tables on the CustomerID field. The Customers table contains no duplicate CustomerID fields, but the Orders table does because each customer can have many orders. Using DISTINCTROW produces a list of companies that have at least one order but without any details about those orders.


```vb
Sub AllDistinctX() 
    Dim dbs As Database, rst As Recordset 
    ' Modify this line to include the path to Northwind 
    ' on your computer. 
    Set dbs = OpenDatabase("Northwind.mdb") 
   
    ' Join the Customers and Orders tables on the  
    ' CustomerID field. Select a list of companies  
    ' that have at least one order. 
 
    Set rst = dbs.OpenRecordset("SELECT DISTINCTROW " _ 
        &; "CompanyName FROM Customers " _ 
        &; "INNER JOIN Orders " _ 
        &; "ON Customers.CustomerID = " _ 
        &; "Orders.CustomerID " _ 
        &; "ORDER BY CompanyName;") 
 
    ' Populate the Recordset. 
    rst.MoveLast 
     
    ' Call EnumFields to print the contents of the  
    ' Recordset. Pass the Recordset object and desired 
    ' field width. 
    EnumFields rst, 25 
    dbs.Close 
End Sub 

```

 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

