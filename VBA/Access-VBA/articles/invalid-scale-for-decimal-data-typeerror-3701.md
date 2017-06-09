---
title: Invalid scale for decimal data type. (Error 3701)
keywords: jeterr40.chm5003701
f1_keywords:
- jeterr40.chm5003701
ms.prod: access
ms.assetid: 2c839f39-d3ab-053a-d7b0-1bcde43232d4
ms.date: 06/08/2017
---


# Invalid scale for decimal data type. (Error 3701)

  

**Applies to:** Access 2013 | Access 2016

The scale of a DECIMAL data type must always be less or equal to the precision. For example, the following SQL statement would return this error: CREATE TABLE foo (foo DECIMAL(10,12));

 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

