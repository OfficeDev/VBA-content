---
title: IN Clause (Microsoft Access SQL)
keywords: jetsql40.chm5277567
f1_keywords:
- jetsql40.chm5277567
ms.prod: access
ms.assetid: 5bca25c0-cd00-140f-79b8-80cd2d0c190b
ms.date: 06/08/2017
---


# IN Clause (Microsoft Access SQL)

  

**Applies to:** Access 2013 | Access 2016

 **In this article**
[Syntax](#sectionSection0)
[Remarks](#sectionSection1)
[ Example](#sectionSection2)


Identifies tables in any external database to which the Microsoft Access database engine can connect, such as a dBASE or Paradox database or an external Microsoft® Access database engine database.

## Syntax
<a name="sectionSection0"> </a>

To identify a destination table:

[SELECT | INSERT] INTO  _destination_ IN { _path_ | [" _path_ " " _type_ "] | ["" [ _type_; DATABASE = _path_ ]]}

To identify a source table:

FROM  _tableexpression_ IN { _path_ | [" _path_ " " _type_ "] | ["" [ _type_; DATABASE = _path_ ]]}

A SELECT statement containing an IN clause has these parts:



|**Part**|**Description**|
|:-----|:-----|
| _destination_|The name of the external table into which data is inserted.|
| _tableexpression_|The name of the table or tables from which data is retrieved. This argument can be a single table name, a saved query, or a compound resulting from an [INNER JOIN](http://msdn.microsoft.com/library/8d16c74c-02c6-12b7-b180-3e7744ef65f3%28Office.15%29.aspx), [LEFT JOIN](http://msdn.microsoft.com/library/9c10525f-98b1-fd4f-8b40-07a32c5c6502%28Office.15%29.aspx), or [RIGHT JOIN](http://msdn.microsoft.com/library/9c10525f-98b1-fd4f-8b40-07a32c5c6502%28Office.15%29.aspx).|
| _path_|The full path for the directory or file containing  _table._|
| _type_|The name of the database type used to create  _table_ if a database is not a Microsoft Access database engine database (for example, dBASE III, dBASE IV, Paradox 3.x, or Paradox 4.x).|

## Remarks
<a name="sectionSection1"> </a>

You can use IN to connect to only one external database at a time.

In some cases, the  _path_ argument refers to the directory containing the database files. For example, when working with dBASE, Microsoft FoxPro®, or Paradox database tables, the _path_ argument specifies the directory containing .dbf or .db files. The table file name is derived from the _destination_ or _tableexpression_ argument.

To specify a non-Microsoft Access database engine database, append a semicolon (;) to the name, and enclose it in single (' ') or double (" ") quotation marks. For example, either 'dBASE IV;' or "dBASE IV;" is acceptable.

You can also use the DATABASE reserved word to specify the external database. For example, the following lines specify the same table:




```sql
…FROM Table IN "" [dBASE IV; DATABASE=C:\DBASE\DATA\SALES;]; 

…FROM Table IN "C:\DBASE\DATA\SALES" "dBASE IV;"
```


 **Note**  


## Example
<a name="sectionSection2"> </a>

The following table shows how you can use the IN clause to retrieve data from an external database. In each example, assume the hypothetical Customers table is stored in an external database.





|**External database**|**SQL statement**|
|:-----|:-----|
|Microsoft® Access atabase engine database|
```
SELECT CustomerID
FROM Customers
IN OtherDB.mdb 
WHERE CustomerID Like "A*";
```

|
|dBASE III or IV. To retrieve data from a dBASE III table, substitute "dBASE III;" for "dBASE IV;".|
```vb
SELECT CustomerID
FROM Customer
IN "C:\DBASE\DATA\SALES" "dBASE IV;"
WHERE CustomerID Like "A*";

```

|
|dBASE III or IV using Database syntax.|
```
SELECT CustomerID
FROM Customer
IN "" [dBASE IV; Database=C:\DBASE\DATA\SALES;] 
WHERE CustomerID Like "A*";

```

|
|Paradox 3.x or 4.x. To retrieve data from a Paradox version 3.x table, substitute "Paradox 3.x;" for "Paradox 4.x;".|
```vb
SELECT CustomerID
FROM Customer
IN "C:\PARADOX\DATA\SALES" "Paradox 4.x;"
WHERE CustomerID Like "A*";

```

|
|Paradox 3.x or 4.x using Database syntax|
```
SELECT CustomerID
FROM Customer
IN "" [Paradox 4.x;Database=C:\PARADOX\DATA\SALES;] 
WHERE CustomerID Like "A*";

```

|
|A Microsoft Excel worksheet|
```sql
SELECT CustomerID, CompanyName
FROM [Customers$] 
IN "c:\documents\xldata.xls" "EXCEL 5.0;"
WHERE CustomerID Like "A*"
ORDER BY CustomerID;

```

|
|A named range in a worksheet|
```

SELECT CustomerID, CompanyName
FROM CustomersRange
IN "c:\documents\xldata.xls" "EXCEL 5.0;"
WHERE CustomerID Like "A*"
ORDER BY CustomerID;
```

|


 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

