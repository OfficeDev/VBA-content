---
title: MaxRecords Property
keywords: vbaac10.chm4422
f1_keywords:
- vbaac10.chm4422
ms.prod: access
api_name:
- Access.MaxRecords
ms.assetid: 30ea62b8-9304-2cdf-ff2f-d8ed665b16b4
ms.date: 06/08/2017
---


# MaxRecords Property

  

**Applies to:** Access 2013 | Access 2016

Specifies the maximum number of records that will be returned by:


- A query that returns data from an ODBC database to an Microsoft Access database . 
    
- A view that returns data from a SQL database to an Access project (.adp). 
    

## Setting

The  **MaxRecords** property setting is a Long Integer value representing the number of records that will be returned.

In a Microsoft Access database, you can set this property by using the query's property sheet or Visual Basic .


## Remarks

When you set this property in Visual Basic you use the ADO  **MaxRecords** property.

Records are returned in the order specified by the query's ORDER BY clause.

You can use the  **MaxRecords** property in situations where limited system resources might prohibit a large number of returned records.

 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

