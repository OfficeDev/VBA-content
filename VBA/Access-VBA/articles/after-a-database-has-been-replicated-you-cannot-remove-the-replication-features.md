---
title: After a database has been replicated, you cannot remove the replication features from the database. (Error 3459)
keywords: jeterr40.chm5003459
f1_keywords:
- jeterr40.chm5003459
ms.prod: access
ms.assetid: 06429f0d-7b78-d7e7-3c67-23142282b0ca
ms.date: 06/08/2017
---


# After a database has been replicated, you cannot remove the replication features from the database. (Error 3459)

  

**Applies to:** Access 2013 | Access 2016

You cannot remove the replication features from a database that has been replicated, either by setting its  **Replicable** property to "T" or by converting it with the Microsoft Windows Briefcase, Microsoft Access, or the Replication Manager. Using DAO to set the database's **Replicable** property to "F" or any other value has no effect on the replicated database. If you need to use a copy of the database that does not have the properties, fields, tables, and other characteristics associated with replication, open the backup copy of the database made before the database was first converted.

 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

