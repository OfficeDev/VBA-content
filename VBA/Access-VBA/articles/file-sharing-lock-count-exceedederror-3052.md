---
title: File sharing lock count exceeded. (Error 3052)
keywords: jeterr40.chm5003052
f1_keywords:
- jeterr40.chm5003052
ms.prod: access
ms.assetid: 682df85c-6e2e-26d4-d035-d787de5672ae
ms.date: 06/08/2017
---


# File sharing lock count exceeded. (Error 3052)

  

**Applies to:** Access 2013 | Access 2016

You have exceeded the maximum number of locks allowed on a recordset. This limit is specified by the MaxLocksPerFile setting in your system registry. The default value is 9500, and can be changed either by editing the registry with Regedit.exe or with the  **SetOption** method.

Some other factors that may cause an application to reach this threshold include the following:


- amount of available memory
    
- size of rows in the recordset
    
- network operating system restrictions
    

 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

