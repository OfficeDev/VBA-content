---
title: SQL property [ADO]
ms.prod: access
ms.assetid: 210adcbb-5c89-150b-4c61-6a52dea9af56
ms.date: 06/08/2017
---


# SQL property [ADO]

  

**Applies to:** Access 2013 | Access 2016



Indicates the query string used to retrieve the [Recordset](http://msdn.microsoft.com/library/0f963bf8-f066-dc8a-b754-f427de712df1%28Office.15%29.aspx).
You can set the  **SQL** property at design time in the[RDS.DataControl](http://msdn.microsoft.com/library/ac430669-7628-696c-c036-b5d35405d788%28Office.15%29.aspx) object's OBJECT tags, or at run time in scripting code.

## Parameters


-  _QueryString_
    
- A  **String** value that contains a valid SQL data request.
    
-  _DataControl_
    
- An object variable that represents an  **RDS.DataControl** object.
    

## Remarks

In general, this is an SQL statement (using the dialect of the database server), such as . To ensure that records are matched and updated accurately, an updatable query must contain a field other than a Long Binary field or a computed field.

The  **SQL** property is optional if a custom server-side business objects retrieves the data for the client.

 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

