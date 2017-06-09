---
title: Retrieve Records Using Access SQL
ms.prod: access
ms.assetid: b613a24a-2fc4-ac18-501f-c44b5cc2a45d
ms.date: 06/08/2017
---


# Retrieve Records Using Access SQL

The most basic and most often used SQL statement is the  **[SELECT](http://msdn.microsoft.com/library/A5C9DA94-5F9E-0FC0-767A-4117F38A5EF3%28Office.15%29.aspx)** statement. **SELECT** statements are the workhorses of all SQL statements, and they are commonly referred to as select queries. You use the **SELECT** statement to retrieve data from the database tables, and the results are usually returned in a set of records (or rows) made up of any number of fields (or columns). You must use the **[FROM](from-clause-microsoft-access-sql.md)** clause to designate which table or tables to select from. The basic structure of a **SELECT** statement is:


```sql
SELECT field list  
    FROM table list
```


To select all fields from a table, use an asterisk (*). For example, the following statement selects all the fields and all the records from the Customers table.




```sql
SELECT * 
    FROM tblCustomers 

```

To limit the fields retrieved by the query, simply use the field names instead. For example.



```sql
SELECT [Last Name], Phone 
    FROM tblCustomers 

```

To designate a different name for a field in the result set, use the  **[AS](as-microsoft-access-sql-reserved-word.md)** keyword to establish an alias for that field.



```sql
SELECT CustomerID AS [Customer Number] 
    FROM tblCustomers 

```


## Restricting the Result Set

More often than not, you will not want to retrieve all records from a table. You will want only a subset of those records based on some qualifying criteria. To qualify a  **SELECT** statement, you must use a **[WHERE](where-clause-microsoft-access-sql.md)** clause, which will allow you to specify exactly which records you want to retrieve.


```sql
SELECT * 
    FROM tblInvoices 
    WHERE CustomerID = 1 

```

Be aware of the  `CustomerID = 1` portion of the **WHERE** clause. A **WHERE** clause can contain up to 40 such expressions, and they can be joined with the And or Or logical operators. Using more than one expression allows you to further filter out records in the result set.




```sql
SELECT * 
    FROM tblInvoices 
    WHERE CustomerID = 1 AND InvoiceDate > #01/01/98# 

```

Be aware that the date string is enclosed in number signs (#). If you are using a regular string in an expression, you must enclose the string in single quotation marks ('). For example.




```sql
SELECT * 
    FROM tblCustomers 
    WHERE [Last Name] = 'White' 

```

If you do not know the whole string value, you can use wildcard characters with the  **[Like](like-operator-microsoft-access-sql.md)** operator.




```sql
SELECT * 
    FROM tblCustomers 
    WHERE [Last Name] LIKE 'W*' 

```

There are a number of wildcard characters to choose from, and the following table lists what they are and what they can be used for.


|||
|:-----|:-----|
|**Wildcard character**|**Description**|
|*|Zero or more characters|
|?|Any single character|
|#|Any single digit (0-9)|
|[ _charlist_ ]|Any single character in  _charlist_|
|[! _charlist_ ]|Any single character not in  _charlist_|

## Sorting the Result Set

To specify a particular sort order on one or more fields in the result set, use the optional  **[ORDER BY](order-by-clause-microsoft-access-sql.md)** clause. Records can be sorted in either ascending ( **ASC** ) or descending ( **DESC** ) order; ascending is the default.

Fields referenced in the  **ORDER BY** clause do not have to be part of the **SELECT** statement's field list, and sorting can be applied to string, numeric, and date/time values. Always place the **ORDER BY** clause at the end of the **SELECT** statement.




```sql
SELECT * 
    FROM tblCustomers 
    ORDER BY [Last Name], [First Name] DESC 

```

You can also use the field numbers (or positions) instead of field names in the  **ORDER BY** clause.




```sql
SELECT * 
    FROM tblCustomers 
    ORDER BY 2, 3 DESC 

```


