---
title: Perform Joins Using Access SQL
ms.prod: access
ms.assetid: 1a19bc56-afd3-3917-b503-44b77078483d
ms.date: 06/08/2017
---


# Perform Joins Using Access SQL

In a relational database system like Access, you often need to extract information from more than one table at a time. This can be accomplished by using an SQL  **[JOIN](join-microsoft-access-sql-reserved-word.md)** statement, which enables you to retrieve records from tables that have defined relationships, whether they are one-to-one, one-to-many, or many-to-many.


## INNER JOINs

The  **[INNER JOIN](http://msdn.microsoft.com/library/8D16C74C-02C6-12B7-B180-3E7744EF65F3%28Office.15%29.aspx)**, also known as an equi-join, is the most commonly used type of join. This join is used to retrieve rows from two or more tables by matching a field value that is common between the tables. The fields you join on must have similar data types, and you cannot join on MEMO or OLEOBJECT data types. To build an **INNER JOIN** statement, use the **INNER JOIN** keywords in the **[FROM](from-clause-microsoft-access-sql.md)** clause of a **[SELECT](http://msdn.microsoft.com/library/A5C9DA94-5F9E-0FC0-767A-4117F38A5EF3%28Office.15%29.aspx)** statement. This example uses the **INNER JOIN** to build a result set of all customers who have invoices, in addition to the dates and amounts of those invoices.


```sql
SELECT [Last Name], InvoiceDate, Amount 
   FROM tblCustomers INNER JOIN tblInvoices 
   ON tblCustomers.CustomerID=tblInvoices.CustomerID 
   ORDER BY InvoiceDate 
```

Be aware that the table names are divided by the  **INNER JOIN** keywords and that the relational comparison is after the **ON** keyword. For the relational comparisons, you can also use the <, >, <=, >=, or <> operators, and you can also use the **BETWEEN** keyword. Also note that the ID fields from both tables are used only in the relational comparison; they are not part of the final result set.

To further qualify the  **SELECT** statement, you can use a **[WHERE](where-clause-microsoft-access-sql.md)** clause after the join comparison in the **ON** clause. The following example narrows the result set to include only invoices dated after January 1, 1998.




```sql
SELECT [Last Name], InvoiceDate, Amount 
   FROM tblCustomers INNER JOIN tblInvoices 
   ON tblCustomers.CustomerID=tblInvoices.CustomerID 
   WHERE tblInvoices.InvoiceDate > #01/01/1998# 
   ORDER BY InvoiceDate 
```

When you must join more than one table, you can nest the  **INNER JOIN** clauses. The following example builds on a previous **SELECT** statement to create the result set, but also includes the city and state of each customer by adding the **INNER JOIN** for the tblShipping table.




```sql
SELECT [Last Name], InvoiceDate, Amount, City, State 
   FROM (tblCustomers INNER JOIN tblInvoices 
   ON tblCustomers.CustomerID=tblInvoices.CustomerID) 
      INNER JOIN tblShipping 
      ON tblCustomers.CustomerID=tblShipping.CustomerID 
   ORDER BY InvoiceDate 
```

Be aware that the first  **JOIN** clause is enclosed in parentheses to keep it logically separated from the second **JOIN** clause. It is also possible to join a table to itself by using an alias for the second table name in the **FROM** clause. Suppose that you want to find all customer records that have duplicate last names. You can do this by creating the alias "A" for the second table and checking for first names that are different.




```sql
SELECT tblCustomers.[Last Name], 
   tblCustomers.[First Name] 
   FROM tblCustomers INNER JOIN tblCustomers AS A 
   ON tblCustomers.[Last Name]=A.[Last Name] 
   WHERE tblCustomers.[First Name]<>A.[First Name] 
   ORDER BY tblCustomers.[Last Name] 
```


## OUTER JOINs

An  **[OUTER JOIN](http://msdn.microsoft.com/library/9c10525f-98b1-fd4f-8b40-07a32c5c6502%28Office.15%29.aspx)** is used to retrieve records from multiple tables while preserving records from one of the tables, even if there is no matching record in the other table. There are two types of **OUTER JOINs** that the Access database engine supports: **LEFT OUTER JOINs** and **RIGHT OUTER JOINs**. Think of two tables that are beside each other, a table on the left and a table on the right. The ** LEFT OUTER JOIN** selects all rows in the right table that match the relational comparison criteria, and also selects all rows from the left table, even if no match exists in the right table. The **RIGHT OUTER JOIN** is simply the reverse of the **LEFT OUTER JOIN**; all rows in the right table are preserved instead.

As an example, suppose that you want to determine the total amount invoiced to each customer, but if a customer has no invoices, you want to show it by displaying the word "NONE."




```sql
SELECT [Last Name] &; ', ' &;  [First Name] AS Name, 
   IIF(Sum(Amount) IS NULL,'NONE',Sum(Amount)) AS Total 
   FROM tblCustomers LEFT OUTER JOIN tblInvoices 
   ON tblCustomers.CustomerID=tblInvoices.CustomerID 
   GROUP BY [Last Name] &; ', ' &;  [First Name] 
```

Several things occur in the previous SQL statement. The first is the use of the string concatenation operator "&;". This operator allows you to join two or more fields together as one string. The second is the immediate if ( **IIf** ) statement, which checks to see if the total is null. If it is, the statement returns the word "NONE." If the total is not null, the value is returned. The final thing is the **OUTER JOIN** clause. Using the **LEFT OUTER JOIN** preserves the rows in the left table so that you see all customers, even those who do not have invoices.

 **OUTER JOINs** can be nested inside **INNER JOINs** in a multi-table join, but **INNER JOINs** cannot be nested inside **OUTER JOINs**.


## The Cartesian product

A term that often comes up when discussing joins is the Cartesian product. A Cartesian product is defined as "all possible combinations of all rows in all tables." For example, if you were to join two tables without any kind of qualification or join type, you would get a Cartesian product.


```sql
SELECT * 
   FROM tblCustomers, tblInvoices 
```

This is not a good thing, especially with tables that contain hundreds or thousands of rows. You should avoid creating Cartesian products by always qualifying your joins.


## The UNION operator

Although the  **[UNION](http://msdn.microsoft.com/library/A5139921-51E5-7D96-74E3-11C3FD5F7EAA%28Office.15%29.aspx)** operator, also known as a union query, is not technically a join, it is included here because it does involve combining data from multiple sources of data into one result set, which is similar to some types of joins. The **UNION** operator is used to splice together data from tables, **SELECT** statements, or queries, while leaving out any duplicate rows. Both data sources must have the same number of fields, but the fields do not have to be the same data type. Suppose that you have an Employees table that has the same structure as the Customers table, and you want to build a list of names and e-mail addresses by combining both tables.


```sql
SELECT [Last Name], [First Name], Email 
   FROM tblCustomers 
UNION 
SELECT [Last Name], [First Name], Email 
   FROM tblEmployees 
```

To retrieve all fields from both tables, you could use the  **[TABLE](table-microsoft-access-sql-reserved-word.md)** keyword, like this.




```sql
TABLE tblCustomers 
UNION 
TABLE tblEmployees 
```

The  **UNION** operator will not display any records that are exact duplicates in both tables, but this can be overridden by using the **[ALL](all-microsoft-access-sql-reserved-word.md)** predicate after the **UNION** keyword, like this:




```sql
SELECT [Last Name], [First Name], Email 
   FROM tblCustomers 
UNION ALL 
SELECT [Last Name], [First Name], Email 
   FROM tblEmployees 
```


## The TRANSFORM statement

Although the  **[TRANSFORM](http://msdn.microsoft.com/library/419770B1-C833-959D-A84D-56C68764799F%28Office.15%29.aspx)** statement, also known as a crosstab query, is also not technically considered a join, it is included here because it does involve combining data from multiple sources of data into one result set, which is similar to some types of joins.

A  **TRANSFORM** statement is used to calculate a sum, average, count, or other type of aggregate total on records. It then displays the information in a grid or spreadsheet format with data grouped both vertically (rows) and horizontally (columns). The general form for a **TRANSFORM** statement is the following.




```sql
   TRANSFORM aggregating function 
   SELECT statement 
   PIVOT column heading field 
```

An example scenario could be if you want to build a datasheet that displays the invoice totals for each customer on a year-by-year basis. The vertical headings will be the customer names, and the horizontal headings will be the years. You can modify a previous SQL statement to fit the transform statement.




```sql
TRANSFORM 
IIF(Sum([Amount]) IS NULL,'NONE',Sum([Amount])) 
   AS Total 
SELECT [Last Name] &; ', ' &; [First Name] AS Name 
      FROM tblCustomers LEFT JOIN tblInvoices 
      ON tblCustomers.CustomerID=tblInvoices.CustomerID 
      GROUP BY [Last Name] &; ', ' &; [First Name] 
PIVOT Format(InvoiceDate, 'yyyy') 
   IN ('1996','1997','1998','1999','2000') 
```

Be aware that the aggregating function is the  **[Sum](sum-function-microsoft-access-sql.md)** function, the vertical headings are in the **[GROUP BY](group-by-clause-microsoft-access-sql.md)** clause of the **SELECT** statement, and the horizontal headings are determined by the field listed after the **PIVOT** keyword.


