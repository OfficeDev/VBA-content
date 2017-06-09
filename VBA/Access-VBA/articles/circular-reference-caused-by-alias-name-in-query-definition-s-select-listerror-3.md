---
title: Circular reference caused by alias <name> in query definition's SELECT list. (Error 3103)
keywords: jeterr40.chm5003103
f1_keywords:
- jeterr40.chm5003103
ms.prod: access
ms.assetid: b0f5d8a6-4735-367f-dd27-af3d97816430
ms.date: 06/08/2017
---


# Circular reference caused by alias <name> in query definition's SELECT list. (Error 3103)

  

**Applies to:** Access 2013 | Access 2016

The specified alias created a reference that cannot be resolved. This error can occur, for example, if you enter the following SQL statement, in which A is the circular reference:




```sql
SELECT A + B AS C, C + D AS E, E + F AS A

FROM MyTable;
```




```c#
SELECT week1 + week2 as hours, hours + overtime as gross, gross + ytdpay as week1FROM EmployeePay

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

