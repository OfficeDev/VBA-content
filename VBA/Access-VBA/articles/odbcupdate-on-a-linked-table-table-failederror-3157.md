---
title: ODBC - update on a linked table <table> failed. (Error 3157)
keywords: jeterr40.chm5003157
f1_keywords:
- jeterr40.chm5003157
ms.prod: access
ms.assetid: f5d48f2d-7f12-9550-2a4e-27d7b01bf439
ms.date: 06/08/2017
---


# ODBC - update on a linked table <table> failed. (Error 3157)

  

**Applies to:** Access 2013 | Access 2016

Using an ODBC connection, you tried to update data in an ODBC database; that update could not be completed.

Possible causes:


- The update would have caused a rule violation.
    
- The ODBC database is read-only, or you do not have permission to update data in that database. Resolve the read-only condition, or see your system administrator or the person who created the database to obtain the necessary permissions.
    
- The ODBC database is on a network drive and the network is not connected. Make sure the network is available, and then try the operation again.
    

 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

