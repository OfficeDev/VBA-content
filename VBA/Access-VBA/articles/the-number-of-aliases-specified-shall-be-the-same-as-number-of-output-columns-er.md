---
title: The number of aliases specified shall be the same as number of output columns (Error 3731)
keywords: jeterr40.chm5003731
f1_keywords:
- jeterr40.chm5003731
ms.prod: access
ms.assetid: 884d1f65-60d7-66f3-f404-d7b0b996c46a
ms.date: 06/08/2017
---


# The number of aliases specified shall be the same as number of output columns (Error 3731)

  

**Applies to:** Access 2013 | Access 2016

This error occurs when trying to create a view through SQL DDL. The error occurs when a different number of correlation names or aliases are defined from what is in the SELECT statement. For example, the following syntax would generate this error: CREATE VIEW foo (col1, col2) AS SELECT col1 FROM table1.

 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

