---
title: Object is invalid or no longer set. (Error 3420)
keywords: jeterr40.chm5003420
f1_keywords:
- jeterr40.chm5003420
ms.prod: access
ms.assetid: 5744c5e1-1cf7-52eb-6ac3-a35044f2f6d6
ms.date: 06/08/2017
---


# Object is invalid or no longer set. (Error 3420)

  

**Applies to:** Access 2013 | Access 2016

You are attempting to reference an object that is no longer valid or has not been set.

Possible causes:


- The object has been closed.
    
- The object has been orphaned (the parent object has been closed or deleted).
    
- The object is out of scope.
    
- The object library is not registered in the Microsoft Windows Registry.
    
- You are trying to reference a method or property of the collection, but you have not assigned it to a variable first. For example, to reference the  **Name** property, use the following:
    
```vb
  Dim dbsPublish As Database 
Set dbsPublish = OpenDatabase("BIBLIO.mdb")
dbname = dbsPublish.Name

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

