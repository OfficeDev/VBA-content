---
title: Selected collating sequence not supported by the operating system. (Error 3228)
keywords: jeterr40.chm5003228
f1_keywords:
- jeterr40.chm5003228
ms.prod: access
ms.assetid: 39eae021-6584-478d-cb75-ce5131341dba
ms.date: 06/08/2017
---


# Selected collating sequence not supported by the operating system. (Error 3228)

  

**Applies to:** Access 2013 | Access 2016

There is an invalid  **CollatingSequence** setting in the **Paradox** key of the Microsoft Windows Registry.

 To fix the CollatingSequence setting


1. Exit your application.
    
2. Start the Registry Editor, navigate to the  **HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\14.0\Access Connectivity Engine\Engines\Paradox** key, and select the **CollatingSequence** value.
    
3. On the  **Edit** menu, click **Modify**.
    
4. Specify either  **ascii** or **International** in the **Value data** box.
    
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

