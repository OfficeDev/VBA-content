---
title: Could not parse field names from the first row of the import table. (Error 3245)
keywords: jeterr40.chm5003245
f1_keywords:
- jeterr40.chm5003245
ms.prod: access
ms.assetid: ac70f60f-e43b-30cc-fea4-969c132819df
ms.date: 06/08/2017
---


# Could not parse field names from the first row of the import table. (Error 3245)

  

**Applies to:** Access 2013 | Access 2016

The first row of data contains invalid field names, such as quoted and unquoted strings in the same field name. In the following example, the third and fourth field names cannot be parsed:




```
"Name", Date, "ID " Number, Phone" Number"

```

The first two fields are valid, but the third and fourth are not because they contain nonspace characters outside the quotation marks.
Check the import table for properly matched quotation marks, and then try the import operation again
 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

