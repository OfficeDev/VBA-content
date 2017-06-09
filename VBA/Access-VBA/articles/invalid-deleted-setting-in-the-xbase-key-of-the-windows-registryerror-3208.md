---
title: Invalid Deleted setting in the Xbase key of the Windows Registry. (Error 3208)
ms.prod: access
ms.assetid: f28f55fc-2332-d008-8d8f-013175c22f5b
ms.date: 06/08/2017
---


# Invalid Deleted setting in the Xbase key of the Windows Registry. (Error 3208)

  

**Applies to:** Access 2013 | Access 2016

There is an invalid  **Deleted** setting in the **Xbase** key of the Microsoft Windows Registry.

 To fix the Deleted setting


1. Exit your application.
    
2. Start the Registry Editor, navigate to the  **HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\14.0\Access Connectivity Engine\Engines\Xbase** key, and select the **Deleted** value.
    
3. On the  **Edit** menu, click **Modify**.
    
4. Correct the  **Deleted** data in the **Value data** box.
    
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

