---
title: HAVING clause <name> without grouping or aggregation. (Error 3091)
ms.prod: access
ms.assetid: db7b0a94-4333-83fe-6a7c-d3e8d6311d81
ms.date: 06/08/2017
---


# HAVING clause <name> without grouping or aggregation. (Error 3091)

  

**Applies to:** Access 2013 | Access 2016

This error occurs when the query evaluates the identifiers in the SELECT statement. The error occurs because a GROUP BY clause is not specified before the HAVING clause or because the column referenced in the HAVING clause is not in an AGGREGATE function. For example this error would occur with SELECT col1 FROM table1 HAVING col1 > 20. If the statement was changed to SELECT col1 FROM table1 GROUP BY col1 HAVING col1 > 20, then the statement would be valid. Alternatively, the following would be valid SELECT col1, count(col2) FROM table1 GROUP BY col1 HAVING count(col1) > 20.

 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

