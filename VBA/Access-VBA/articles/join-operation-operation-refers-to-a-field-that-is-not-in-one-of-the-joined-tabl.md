---
title: JOIN operation <operation> refers to a field that is not in one of the joined tables. (Error 3082)
keywords: jeterr40.chm5003082
f1_keywords:
- jeterr40.chm5003082
ms.prod: access
ms.assetid: 13a1b996-709e-198a-fe68-9a23fd39f6a7
ms.date: 06/08/2017
---


# JOIN operation <operation> refers to a field that is not in one of the joined tables. (Error 3082)

  

**Applies to:** Access 2013 | Access 2016

You tried to execute an SQL statement that contains an invalid join. This error occurs when you try to create a join on a field that is not in one of the joined tables. The following example produces this error.




```sql
SELECT Authors.FirstName, Titles.ISBN 
FROM Authors, Titles, AuthorTitles, 
Authors INNER JOIN Titles ON Authors.ID = AuthorTitles.ISBN;
```

The error occurs because the join involves the Authors and Titles tables, but the joined fields are in the Authors and AuthorTitles tables.
 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

