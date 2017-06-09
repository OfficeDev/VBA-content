---
title: Field <name> is based on an expression and cannot be edited. (Error 3327)
ms.prod: access
ms.assetid: 7d7c1e1f-645e-b111-60c3-666640d8bde1
ms.date: 06/08/2017
---


# Field <name> is based on an expression and cannot be edited. (Error 3327)

  

**Applies to:** Access 2013 | Access 2016

For example, if a stored query or view with a column made up of an expression was created, you would not be able to update that column. The following would return this error: CREATE VIEW VCustomer AS SELECT (FirstName &; LastName) AS Test FROM Customer followed by UPDATE Test FROM VCustomer

 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

