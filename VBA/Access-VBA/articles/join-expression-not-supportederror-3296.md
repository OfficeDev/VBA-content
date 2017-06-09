---
title: Join expression not supported. (Error 3296)
ms.prod: access
ms.assetid: 42ae73b1-2543-1850-13a3-57ed42c54720
ms.date: 06/08/2017
---


# Join expression not supported. (Error 3296)

  

**Applies to:** Access 2013 | Access 2016

Possible causes:



- Your SQL statement contains multiple joins in which the results of the query can differ, depending on the order in which the joins are performed. You may want to create a separate query to perform the first join, and then include that query in your SQL statement.
    
- The ON statement in your JOIN operation is incomplete or contains too many tables. You may want to put your ON expression in a WHERE clause.
    

 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

