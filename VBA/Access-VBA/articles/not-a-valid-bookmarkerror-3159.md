---
title: Not a valid bookmark. (Error 3159)
keywords: jeterr40.chm5003159
f1_keywords:
- jeterr40.chm5003159
ms.prod: access
ms.assetid: 99e8083c-d098-916f-3160-d9787e354216
ms.date: 06/08/2017
---


# Not a valid bookmark. (Error 3159)

  

**Applies to:** Access 2013 | Access 2016

You tried to set a bookmark to an invalid string.

This error can occur if you set the  **Bookmark** property to a string that is invalid or was not saved from previously reading a **Bookmark** property. For example, the following code produces this error:



```VB.net
Sub SetBookmark() 
    Dim dbs As Database 
    Dim rstEmployees As Recordset 
    Dim strPlaceholder As String 

    Set dbs = OpenDatabase("Northwind.mdb") 

    Set rstEmployees = _ 
        dbs.OpenRecordset _
        ("Employees", dbOpenDynaset) 

    strPlaceholder = "1" 

    rstEmployees.Bookmark = strPlaceholder    ' Not a valid bookmark. 
End Sub
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

