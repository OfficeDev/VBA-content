---
title: Invalid DataCodePage option in initialization setting. (Error 3337)
ms.prod: access
ms.assetid: 51df967e-82dd-38c3-e413-dfbf728d065d
ms.date: 06/08/2017
---


# Invalid DataCodePage option in initialization setting. (Error 3337)

  

**Applies to:** Access 2013 | Access 2016

The  **DataCodePage** setting for the external data source you are attempting to use is not valid. This setting is in the corresponding **HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\14.0\Access Connectivity Engine\Engines** _&lt;external data source ISAM>_ in the Microsoft Windows Registry.

Valid settings are:


-  **OEM** — Data is stored as OEM data; OemToAnsi and AnsiToOem conversions are done.
    
-  **ANSI** — Data is stored as ANSI data; OemToAnsi and AnsiToOem conversions are not done.
    

 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

