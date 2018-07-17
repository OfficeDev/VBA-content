---
title: Parameter <name> specified where a database name is required. (Error 3217)
keywords: jeterr40.chm5003217
f1_keywords:
- jeterr40.chm5003217
ms.prod: access
ms.assetid: d6d700f2-5df5-5d26-a6ee-706ca4c1a12a
ms.date: 06/08/2017
---


# Parameter <name> specified where a database name is required. (Error 3217)

  

**Applies to:** Access 2013 | Access 2016

You created a parameter query that specifies an invalid parameter type. The following example produces this error:




```sql
PARAMETERS Param1 Text;

SELECT CustomerID
FROM Customers IN Param1;
```

 `Param1` is a text parameter, but the FROM clause requires a database parameter.
Change the parameter type from Text to Database, and then try the operation again.
 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

