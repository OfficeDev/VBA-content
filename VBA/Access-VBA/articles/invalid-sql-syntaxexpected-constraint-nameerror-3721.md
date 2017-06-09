---
title: Invalid SQL syntax - expected CONSTRAINT name. (Error 3721)
ms.prod: access
ms.assetid: 14da04b2-b7d0-3e23-20fe-20e42ef4b3d7
ms.date: 06/08/2017
---


# Invalid SQL syntax - expected CONSTRAINT name. (Error 3721)

  

**Applies to:** Access 2013 | Access 2016

When defining referential integrity from a SQL DDL statement it is necessary to name a constraint when using the CONSTRAINT keyword. If a constraint name is not desired, then do not use the CONSTRAINT keyword. An example of this error would be: CREATE TABLE Customers (CLstNm TEXT(50), CFrstNm TEXT(25), CONSTRAINT PRIMARY KEY (CFrstNm, CLstNm));.

To prevent the error, include a name after the CONSTRAINT keyword:
CREATE TABLE Customers (CLstNm TEXT(50), CFrstNm TEXT(25), CONSTRAINT pkCustomers PRIMARY KEY (CFrstNm, CLstNm));.
or do not use the CONSTRAINT keyword:
CREATE TABLE Customers (CLstNm TEXT(50), CFrstNm TEXT(25), PRIMARY KEY (CFrstNm, CLstNm));.
 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

