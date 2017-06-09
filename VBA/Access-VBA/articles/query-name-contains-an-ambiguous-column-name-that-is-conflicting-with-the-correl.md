---
title: Query <name> contains an ambiguous column name that is conflicting with the correlation (alias) name <name>. (Error 3807)
keywords: jeterr40.chm5003807
f1_keywords:
- jeterr40.chm5003807
ms.prod: access
ms.assetid: 40113ce0-9911-8bb5-ffbf-35b558ca26c8
ms.date: 06/08/2017
---


# Query <name> contains an ambiguous column name that is conflicting with the correlation (alias) name <name>. (Error 3807)

  

**Applies to:** Access 2013 | Access 2016

Either fully qualify the column name or change the correlation (alias) name. A SQL SELECT statement is using a correlation (alias) name that is used in one of the base tables in the FROM clause of the statement. Previous versions of Microsoft Jet returned an incorrect result set with this. To prevent the changing of results sets, this error message is being propagated.

This SQL SELECT statement will work when using the Microsoft OLE DB Provider for Jet by itself or through Active Data Objects (ADO). If this SQL SELECT statement is being used through any part of Microsoft Access outside of ADO using Microsoft OLE DB Provider for Jet, you will need to change the correlation (alias) name to something other than the column name in the base table.
An example of this would be the following: CREATE TABLE Orders (OrderDate DATE, Freight DOUBLE);. The following SQL SELECT statement would now return an error: SELECT OrderDate AS A1, Freight + Freight AS OrderDate. The workaround would be to change the correlation (alias) name OrderDate to some other name or to run this query through ADO using the Microsoft OLE DB Provider for Jet.
 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

