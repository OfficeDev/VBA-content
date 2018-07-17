---
title: Cannot group on fields selected with '*' <table name>. (Error 3120)
keywords: jeterr40.chm5003120
f1_keywords:
- jeterr40.chm5003120
ms.prod: access
ms.assetid: 34cce8ec-dc95-7f1d-8537-9dd7dbbc442d
ms.date: 06/08/2017
---


# Cannot group on fields selected with '*' <table name>. (Error 3120)

  

**Applies to:** Access 2013 | Access 2016

You tried to execute a SELECT statement that groups or totals all fields in a single table, selected with an asterisk ( * ). This error occurs, for example, if you enter the following SQL statement:




```sql
SELECT Orders.* FROM Orders GROUP BY ShipVia;

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

