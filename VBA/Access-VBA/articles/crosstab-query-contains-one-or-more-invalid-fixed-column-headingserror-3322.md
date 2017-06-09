---
title: Crosstab query contains one or more invalid fixed column headings. (Error 3322)
ms.prod: access
ms.assetid: 863793f8-2c82-13b5-81cb-1ab3a766893a
ms.date: 06/08/2017
---


# Crosstab query contains one or more invalid fixed column headings. (Error 3322)

  

**Applies to:** Access 2013 | Access 2016

The list of entries for a fixed column heading of a crosstab query (the PIVOT clause in a TRANSFORM statement) is not valid.

Possible causes:


- There are no entries in the list. There must be at least one value in the parentheses following the IN reserved word in the PIVOT clause.
    
- A blank entry appears in the list. Two commas in a row (,,) create a blank entry.
    
- The list contains a field name of more than 64 characters.
    

Correct the PIVOT clause and execute the query again.
 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

