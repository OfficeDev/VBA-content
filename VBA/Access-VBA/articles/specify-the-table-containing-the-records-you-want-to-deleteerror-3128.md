---
title: Specify the table containing the records you want to delete. (Error 3128)
keywords: jeterr40.chm5003128
f1_keywords:
- jeterr40.chm5003128
ms.prod: access
ms.assetid: f6c49cba-5b9c-775c-625a-6d1e79c8adf0
ms.date: 06/08/2017
---


# Specify the table containing the records you want to delete. (Error 3128)

  

**Applies to:** Access 2013 | Access 2016

You tried to execute a delete query but the query does not specify the name of the table containing the records you want to delete.

Possible cause:


- You did not type an asterisk for each table in the ALL, DISTINCT, DISTINCTROW predicates. Instead, you typed field names (for example,  `Customers.Address` instead of `Customers.*`).
    

 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

