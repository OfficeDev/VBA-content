---
title: Circular reference caused by <query reference>. (Error 3102)
keywords: jeterr40.chm5003102
f1_keywords:
- jeterr40.chm5003102
ms.prod: access
ms.assetid: f3ff3f1b-8a2f-7038-57f8-0abadde1c3cf
ms.date: 06/08/2017
---


# Circular reference caused by <query reference>. (Error 3102)

  

**Applies to:** Access 2013 | Access 2016

You tried to execute a query that depends on itself for data. For example, this error occurs if you execute either of the following queries:

Query1



```sql
SELECT * FROM Employees, Query2;

```

Query2



```sql
SELECT * FROM Query1;


```

Redesign the queries to eliminate the dependency.
 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

