---
title: Using the Connection Properties in the Microsoft OLE DB Provider for Access
keywords: acmain11.chm1032169
f1_keywords:
- acmain11.chm1032169
ms.prod: access
ms.assetid: 7bf8c7d0-9185-d7b2-505a-6ddc449089b9
ms.date: 06/08/2017
---


# Using the Connection Properties in the Microsoft OLE DB Provider for Access

  

**Applies to:** Access 2013 | Access 2016

To customize the Microsoft® Windows® Registry settings, you can use the connection properties in the Microsoft OLE DB Provider for Access. This is accomplished by referencing a property in the connection object and changing its value. For example, assuming that your connection object is called ADOConnection, the following would yield the same results as going through ADO: 

ADOConnection.Properties("Jet OLEDB:Max Locks Per File") = 20000
The property names are different than the DAO constants and the registry settings. The property names are as follows:
Jet OLEDB:Max Locks Per File
Jet OLEDB:Implicit Commit Sync
Jet OLEDB:Flush Transaction Timeout
Jet OLEDB:Lock Delay
Jet OLEDB:Max Buffer Size
Jet OLEDB:User Commit Sync
Jet OLEDB:Lock Retry
Jet OLEDB:Exclusive Async Delay
Jet OLEDB:Shared Async Delay
Jet OLEDB:Page Timeout
Jet OLEDB:Recycle Long-Valued Pages
 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

