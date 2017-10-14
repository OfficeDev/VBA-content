---
title: ODBCTimeout Property
keywords: vbaac10.chm4443
f1_keywords:
- vbaac10.chm4443
ms.prod: access
api_name:
- Access.ODBCTimeout
ms.assetid: ebcac9df-87a9-481c-32cc-d28bb9f37717
ms.date: 06/08/2017
---


# ODBCTimeout Property

  

**Applies to:** Access 2013 | Access 2016

You can use the  **ODBCTimeout** property to specify the number of seconds Microsoft Access waits before a time-out error occurs when a query is run on an Open Database Connectivity (ODBC) database.


## Setting

The  **ODBCTimeout** property is an Integer value representing the number of seconds Microsoft Access waits. The default is 60 seconds. When this property is set to 0, no time-out error occurs.

You can set this property by using the query's property sheet or Data Access Objects (DAO) in Visual Basic code.


## Remarks

When you're using an ODBC database, such as Microsoft SQL Server, there may be delays due to network traffic or heavy use of the ODBC server. The  **ODBCTimeout** property lets you specify how long Microsoft Access waits for a network connection before a time-out error occurs.

 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

