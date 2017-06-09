---
title: Update/delete conflict - This updated record was deleted at another replica. (Error 3736)
keywords: jeterr40.chm5003736
f1_keywords:
- jeterr40.chm5003736
ms.prod: access
ms.assetid: d8e66115-9a71-72b1-137b-61305057fb00
ms.date: 06/08/2017
---


# Update/delete conflict - This updated record was deleted at another replica. (Error 3736)

  

**Applies to:** Access 2013 | Access 2016

When a record is deleted at one replica, but updated at another replica, the deleted record always wins in the conflict that occurs when the two replicas synchronize. The updated record is logged in the conflict table. To reverse the initial resolution of the conflict, reinsert the conflict record. To accept the current resolution, delete the conflict record.

 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

