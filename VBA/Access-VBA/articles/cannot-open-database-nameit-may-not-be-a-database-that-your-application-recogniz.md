---
title: Cannot open database <name>. It may not be a database that your application recognizes, or the file may be corrupt. (Error 3049)
keywords: jeterr40.chm5003049
f1_keywords:
- jeterr40.chm5003049
ms.prod: access
ms.assetid: 5441640a-c2e9-ac40-f7d7-1b1a216c9fd8
ms.date: 06/08/2017
---


# Cannot open database <name>. It may not be a database that your application recognizes, or the file may be corrupt. (Error 3049)

  

**Applies to:** Access 2013 | Access 2016

Possible causes:



- You tried to open a non-Microsoft Access database engine database, such as a Btrieve, dBASE, or Paradox database or table. Instead, link the database or table to an existing Microsoft Jet database.
    
- You tried to import or link a Paradox database or table, and the associated .px file could not be found. Make sure the .px file is the same directory as the Paradox database or table, and then try the operation again.
    
- If the specified database is a Microsoft Jet database, it is corrupted. You should attempt to repair the database. If the database cannot be repaired, restore the database from a backup copy, or create a new database.
    

 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

