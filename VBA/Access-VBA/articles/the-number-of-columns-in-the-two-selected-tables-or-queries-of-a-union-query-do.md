---
title: The number of columns in the two selected tables or queries of a union query do not match. (Error 3307)
keywords: jeterr40.chm5003307
f1_keywords:
- jeterr40.chm5003307
ms.prod: access
ms.assetid: fd745328-831b-c72e-b4b1-b80e34f5a838
ms.date: 06/08/2017
---


# The number of columns in the two selected tables or queries of a union query do not match. (Error 3307)

  

**Applies to:** Access 2013 | Access 2016

The two tables or queries joined by the UNION operation must generate the same number of columns. Remove columns from the SELECT statement that has too many columns or include more columns in the SELECT statement that has too few.


 **Note**  You can include constants instead of columns in the SELECT statement that has too few columns. For example, the following union query generates three columns from the first SELECT statement but one column and two constants in the second SELECT statement. The query returns all countries in the Employees and Regions tables. From the Employees table, the query also returns the first and last name of an employee. If the country value is from the Regions table, however, the query returns Null in the First Name and Last Name columns.




```sql
SELECT Country, FirstName, LastName FROM Employees 
UNION SELECT Country, NULL, NULL FROM Regions;
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

