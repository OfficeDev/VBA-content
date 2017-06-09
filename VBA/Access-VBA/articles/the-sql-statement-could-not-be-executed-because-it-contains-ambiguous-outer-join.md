---
title: The SQL statement could not be executed because it contains ambiguous outer joins. To force one of the joins to be performed first, create a separate query that performs the first join and then include that query in your SQL statement. (Error 3258)
keywords: jeterr40.chm5003258
f1_keywords:
- jeterr40.chm5003258
ms.prod: access
ms.assetid: 17515e13-d6d8-8a1e-ee6c-ff2af543da0f
ms.date: 06/08/2017
---


# The SQL statement could not be executed because it contains ambiguous outer joins. To force one of the joins to be performed first, create a separate query that performs the first join and then include that query in your SQL statement. (Error 3258)

  

**Applies to:** Access 2013 | Access 2016

You tried to execute an SQL statement that contains multiple joins; the results of the query can differ depending on the order in which the joins are performed. For example, this error can occur if you execute the following SQL statement:




```sql
SELECT * FROM Customers, Orders, [Order Details],
Customers LEFT JOIN Orders 
ON Customers.CustomerID = Orders.CustomerID, 
Orders INNER JOIN [Order Details] 
ON Orders.OrderID = [Order Details].OrderID;

```

Executing this statement produces an error because the order of the joins is ambiguous. To force one of the joins to be performed first, create a separate query that performs the first join and then include that query in your SQL statement. The following queries illustrate how you might construct the preceding query so that the INNER JOIN operation is performed before the LEFT JOIN and RIGHT JOIN operation:
Query1



```sql
SELECT * FROM Orders, [Order Details],
Orders INNER JOIN [Order Details]
ON Orders. OrderID = [Order Details].OrderID;
```

Query2



```sql
SELECT * FROM Customers, Query1,
Customers LEFT JOIN Query1 
ON Customers.CustomerID = Orders.CustomerID;
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

