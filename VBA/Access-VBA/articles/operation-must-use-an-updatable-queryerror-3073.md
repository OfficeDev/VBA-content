---
title: Operation must use an updatable query. (Error 3073)
keywords: jeterr40.chm5003073
f1_keywords:
- jeterr40.chm5003073
ms.prod: access
ms.assetid: 4d304da6-ed0a-4819-8d1f-ba55bf9a41e9
ms.date: 06/08/2017
---


# Operation must use an updatable query. (Error 3073)

  

**Applies to:** Access 2013 | Access 2016

You tried to run, open, or modify a query that is not updatable.

Possible causes:


- You attempted to run a query that tried to update a field that cannot be updated. For example, you may have created the query in such a way that you tried to update a field on the one side of a one-to-many relationship.
    
- You tried to use the obsolete  **OpenQueryDef** method on a query that is in a database opened for read-only access.
    

The database is read-only for one of the following reasons:


- You used the  **OpenDatabase** method or the Visual Basic **Data** control, and opened the database for read-only access.
    
- The database file has been defined as read-only in your network operating system.
    
- In a network environment, you do not have write privileges for the database file.
    

Close the database, resolve the read-only condition, and then reopen it for read/write access.


- You do not have permission to make changes to the query. To change your permission assignments, see your system administrator or the query's creator.
    

 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

