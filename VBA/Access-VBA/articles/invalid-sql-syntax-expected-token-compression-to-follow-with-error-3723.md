---
title: "Invalid SQL Syntax: expected token: COMPRESSION to follow WITH (Error 3723)"
ms.prod: access
ms.assetid: 6d63cc77-dbcf-302d-6957-1439f18dceeb
ms.date: 06/08/2017
---


# Invalid SQL Syntax: expected token: COMPRESSION to follow WITH (Error 3723)

  

**Applies to:** Access 2013 | Access 2016

This error occurs when using CREATE TABLE or ALTER TABLE ALTER COLUMN syntax. It occurs when referencing one of the NATIONAL CHARACTER synonyms for a column type and not using the COMPRESSION keyword following the WITH keyword. The following is a valid SQL statement: CREATE TABLE foo (foo NCHAR WITH COMP);. The following SQL statement would return the error: CREATE TABLE foo (foo NCHAR WITH);.

 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

