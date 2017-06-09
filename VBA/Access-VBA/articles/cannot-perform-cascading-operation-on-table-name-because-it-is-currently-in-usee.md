---
title: Cannot perform cascading operation on table <name> because it is currently in use. (Error 3414)
keywords: jeterr40.chm5003414
f1_keywords:
- jeterr40.chm5003414
ms.prod: access
ms.assetid: 238227df-7dd2-6a72-7c3d-8b76a5bc7834
ms.date: 06/08/2017
---


# Cannot perform cascading operation on table <name> because it is currently in use. (Error 3414)

  

**Applies to:** Access 2013 | Access 2016

You are trying to save a value to a primary key field that is included in a relationship.

In Microsoft Access, the  **Cascade Update Related Fields** option is selected for the relationship, or in DAO code, the **dbRelationUpdateCascade** option is specified for the **Relation** object's **Attributes** property. Therefore, your application is attempting to update the matching field in the related table.
The matching field cannot be updated because you have it open or locked on your computer. To save the record, you must first close the related table.
 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

