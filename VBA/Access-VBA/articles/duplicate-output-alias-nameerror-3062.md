---
title: Duplicate output alias <name>. (Error 3062)
keywords: jeterr40.chm5003062
f1_keywords:
- jeterr40.chm5003062
ms.prod: access
ms.assetid: e0157e7c-d854-4a9a-b5ba-22afa0944cbc
ms.date: 06/08/2017
---


# Duplicate output alias <name>. (Error 3062)

  

**Applies to:** Access 2013 | Access 2016

You tried to execute an SQL statement that has more than one alias with the same name. The following statement, for example, would produce this error:




```sql
SELECT LastName AS Name, FirstName AS Name FROM Employees;

```

Rename one or more of the aliases, and then try the operation again.
 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

