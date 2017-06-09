---
title: Replica has not been synchronized within the replica set retention period. (Error 3743)
keywords: jeterr40.chm5003743
f1_keywords:
- jeterr40.chm5003743
ms.prod: access
ms.assetid: 52fd5406-2664-8cbe-f1ac-f37c3cb7ad5c
ms.date: 06/08/2017
---


# Replica has not been synchronized within the replica set retention period. (Error 3743)

  

**Applies to:** Access 2013 | Access 2016

If the retention period expires for a replica, you cannot synchronize changes between the expired replica and the other replicas in the replica set. If a replica does not synchronize with another replica in the set within the retention period, the next time you attempt to synchronize the replica it gets removed from the replica set. The retention period is established when the database is initially made replicable. If you replicate the database by using Replication Manager, Data Access Objects (DAO), or ActiveX Data Objects (ADO), the default retention period is 60 days. If you replicate the database by using Microsoft Access or Briefcase, the default retention period is 1000 days.

 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

