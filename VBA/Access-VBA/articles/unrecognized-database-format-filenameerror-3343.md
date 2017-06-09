---
title: Unrecognized database format <filename>. (Error 3343)
keywords: jeterr40.chm5003343
f1_keywords:
- jeterr40.chm5003343
ms.prod: access
ms.assetid: d917be92-c946-1764-9409-9368d011390a
ms.date: 06/08/2017
---


# Unrecognized database format <filename>. (Error 3343)

  

**Applies to:** Access 2013 | Access 2016

Possible causes:



- The specified file name is not a Microsoft Access database engine database.
    
- The specified file name is a device name, for example, a printer or a console.
    
- The database file has invalid header information or an unknown sort order.
    
- A commit is pending from another user but the lock file cannot be found.
    
- During a commit, you are attempting to write a Long value larger than the 2K maximum page size.
    
- The database is damaged. Compact the database and then try opening it again.
    

 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

