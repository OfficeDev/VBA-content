---
title: Cannot update. Database or object is read-only. (Error 3027)
keywords: jeterr40.chm5003027
f1_keywords:
- jeterr40.chm5003027
ms.prod: access
ms.assetid: dc8387fe-aac4-46af-5c2f-bbbae7f7edb4
ms.date: 06/08/2017
---


# Cannot update. Database or object is read-only. (Error 3027)

  

**Applies to:** Access 2013 | Access 2016

You tried to save changes in a database that was opened for read-only access.

The database is read-only for one of these reasons:


- You used the  **OpenDatabase** method and opened the database for read-only access.
    
- In Microsoft Visual Basic, you are using the  **Data** control, and you set the **ReadOnly** property to **True**.
    
- The database file is defined as read-only in the operating system or by your network.
    
- The database file is stored on read-only media.
    
- In a network environment, you do not have write privileges for the database file.
    
- When working with a secured database, the database or one of its objects (such as a field or table) may be set to read-only. You may not have permission to access this data with your user name and password.
    

Close the database, resolve the read-only condition, and then reopen the file for read/write access.
 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

