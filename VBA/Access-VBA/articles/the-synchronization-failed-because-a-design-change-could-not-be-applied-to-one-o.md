---
title: The synchronization failed because a design change could not be applied to one of the replicas. (Error 3492)
keywords: jeterr40.chm5003492
f1_keywords:
- jeterr40.chm5003492
ms.prod: access
ms.assetid: 08ba127a-7002-84ae-6f76-65f4aedeb052
ms.date: 06/08/2017
---


# The synchronization failed because a design change could not be applied to one of the replicas. (Error 3492)

  

**Applies to:** Access 2013 | Access 2016

The Microsoft Access database engine attempted to update the database design at one of the replicas. There are several possible reasons why the design could not be updated, including:



- The object you are trying to update is already open at the replica.
    
- You added an enforced relationship to a replica that has a foreign key that references a nonexistent primary key.
    

For additional information regarding the synchronization failure, look in the MSysSchemaProb table, either at the Design Master or the replica that was the target of the synchronization.
 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

