---
title: Could not open Paradox.net. (Error 3172)
keywords: jeterr40.chm5003172
f1_keywords:
- jeterr40.chm5003172
ms.prod: access
ms.assetid: f42fa1fd-fb7c-3c88-f44a-c77925cc520b
ms.date: 06/08/2017
---


# Could not open Paradox.net. (Error 3172)

  

**Applies to:** Access 2013 | Access 2016

Possible causes:



- The  **ParadoxNetPath** value in the \ **HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\14.0\Access Connectivity Engine\Engines\Paradox** Registry key is not valid. This error occurs if you are using the Paradox external database, and the **ParadoxNetPath** value points to a nonexistent directory. Exit the application, correct the invalid entry, and try the operation again.
    
- The  **ParadoxNetPath** Registry value points to a network drive, and you are not connected to that network drive. Make sure the network drive is available, and then try the operation again.
    

 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

