---
title: The object cannot be replicated because the database is not replicated. (Error 3485)
keywords: jeterr40.chm5003485
f1_keywords:
- jeterr40.chm5003485
ms.prod: access
ms.assetid: ca11f046-2fa6-6da3-89ba-eacab953a992
ms.date: 06/08/2017
---


# The object cannot be replicated because the database is not replicated. (Error 3485)

  

**Applies to:** Access 2013 | Access 2016

You cannot replicate an object in a database unless you first replicate the database that contains it. You can replicate the database by:



- Dragging it into the Microsoft Windows Briefcase.
    
- Using DAO programming to set the  **Replicable** property to "T" or the **ReplicableBool** property to **True**.
    
- Using Microsoft Access.
    
- Using the Replication Manager.
    

All objects in the database are replicated when the database is replicated, unless the  **KeepLocal** property has been set on an object.
 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

