---
title: Invalid NetworkAccess setting in the Windows Registry. (Error 3248)
ms.prod: access
ms.assetid: c61fb87b-47eb-9644-f4af-5c5676e8c2b8
ms.date: 06/08/2017
---


# Invalid NetworkAccess setting in the Windows Registry. (Error 3248)

  

**Applies to:** Access 2013 | Access 2016

There is an invalid  **NetworkAccess** setting in the Microsoft Windows Registry.

 To complete this operation


1. Exit your application.
    
2. Start the Registry Editor, and navigate to the  **NetworkAccess** value. Depending on which installable ISAM you are trying to use, the invalid entry is in the **HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\14.0\Access Connectivity Engine\Engines\Xbase** key.
    
3. On the  **Edit** menu, click **Modify**.
    
4. Specify either  **On** or **Off** in the **Value data** box.
    
5. Restart your application, and then try the operation again.
    

 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

