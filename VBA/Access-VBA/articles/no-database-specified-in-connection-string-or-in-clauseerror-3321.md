---
title: No database specified in connection string or IN clause. (Error 3321)
ms.prod: access
ms.assetid: dd1601c0-9998-4ae0-21a0-ec283dc8f0cd
ms.date: 06/08/2017
---


# No database specified in connection string or IN clause. (Error 3321)

  

**Applies to:** Access 2013 | Access 2016

The Microsoft Access database engine is unable to connect to an external ISAM database because you have not specified the name of the database to connect to. When connecting to an external source of data, you must specify a database name.

Possible causes:


- The connection string in the FROM clause of the SELECT statement is missing the parameter, DATABASE=.
    
- The IN clause of the SELECT statement includes a database type argument (indicating to select the data from an external database) but it is missing a database name argument.
    

 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

