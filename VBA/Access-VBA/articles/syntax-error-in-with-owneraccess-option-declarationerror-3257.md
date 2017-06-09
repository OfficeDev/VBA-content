---
title: Syntax error in WITH OWNERACCESS OPTION declaration. (Error 3257)
keywords: jeterr40.chm5003257
f1_keywords:
- jeterr40.chm5003257
ms.prod: access
ms.assetid: a1b4ae18-4efa-d79a-ffec-4ec705a0236b
ms.date: 06/08/2017
---


# Syntax error in WITH OWNERACCESS OPTION declaration. (Error 3257)

  

**Applies to:** Access 2013 | Access 2016

Possible causes:



- The WITH OWNERACCESS OPTION declaration is incomplete or includes a space between OWNER and ACCESS.
    
- The declaration appears in an unexpected and disallowed position in the SQL statement. For example:
    
```sql
  SELECT * WITH OWNERACCESS OPTION FROM [My Table]; 

```


    The WITH OWNERACCESS OPTION declaration should appear at the end of the SQL statement, usually after the ORDER BY clause, if present:
    


```sql
  SELECT * FROM [My Table] WITH OWNERACCESS OPTION;
```

 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

