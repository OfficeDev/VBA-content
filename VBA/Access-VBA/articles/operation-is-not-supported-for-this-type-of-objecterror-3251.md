---
title: Operation is not supported for this type of object. (Error 3251)
keywords: jeterr40.chm5003251
f1_keywords:
- jeterr40.chm5003251
ms.prod: access
ms.assetid: d6a18e82-02ab-4514-6e31-3960e972dd0b
ms.date: 06/08/2017
---


# Operation is not supported for this type of object. (Error 3251)

  

**Applies to:** Access 2013 | Access 2016

You were attempting to execute a method or assign a value to a property that is usually valid for the object, but is not supported in this specific instance. For example, the  **Edit** method is generally valid for **Recordset** objects, but not for a snapshot-type **Recordset**. This error could also occur in cases where the operation is not permitted due to the type or status of the object — as when trying to use the **MovePrevious** method on a forward-only-type **Recordset**. Some operations are also not supported, depending on if you are accessing a Microsoft Access database engine or an ODBC data source.

 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

