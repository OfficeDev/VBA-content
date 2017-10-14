---
title: Group Records in a Result Set Using Access SQL
ms.prod: access
ms.assetid: 217e1a5a-cfe2-0859-1e16-a3d27450008c
ms.date: 06/08/2017
---


# Group Records in a Result Set Using Access SQL

Sometimes records in a table are logically related, as in the case of the invoices table. Because one customer can have many invoices, it could be useful to handle all invoices for one customer as a group, in order to find statistical and summary information about the group.

The key to grouping records is that one or more fields in each record must contain the same value for every record in the group. In the case of the invoices table, the CustomerID field value is the same for every invoice a particular customer has.

To create a group of records, use the  **[GROUP BY](group-by-clause-microsoft-access-sql.md)** clause with the name of the field or fields you want to group with.




```sql
SELECT CustomerID, Count(*) AS [Number of Invoices], 
    Avg(Amount) AS [Average Invoice Amount] 
    FROM tblInvoices 
    GROUP BY CustomerID 

```

Be aware that the statement will return one record that shows the customer ID, the number of invoices the customer has, and the average invoice amount, for every customer who has an invoice record in the invoices table. Because each customer invoice is handled as a group, you are able to count the number of invoices and then determine the average invoice amount.
You can specify a condition at the group level by using the HAVING clause, which is similar to the  **[WHERE](where-clause-microsoft-access-sql.md)** clause. For example, the following query returns only those records for each customer whose average invoice amount is less than 100.



```sql
SELECT CustomerID, Count(*) AS [Number of Invoices], 
    Avg(Amount) AS [Average Invoice Amount] 
    FROM tblInvoices 
    GROUP BY CustomerID 
    HAVING Avg(Amount) < 100 

```


