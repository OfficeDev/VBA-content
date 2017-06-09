---
title: Out of MS-DOS file handles. (Error 3042)
keywords: jeterr40.chm5003042
f1_keywords:
- jeterr40.chm5003042
ms.prod: access
ms.assetid: 03f53859-f944-b5c4-5b3d-e39e240c3120
ms.date: 06/08/2017
---


# Out of MS-DOS file handles. (Error 3042)

  

**Applies to:** Access 2013 | Access 2016

Either close some files, close other applications, or increase the number of available MS-DOS file handles.

To increase the number of file handles


1. Edit your CONFIG.sys file using Microsoft Windows Notepad or another text editor. The CONFIG.sys file is usually located in the root directory of your boot drive. If you do not have a CONFIG.sys file on your system, you must create one that contains the line listed in step 2.
    
2. Look for the line that reads FILES =  _x_, where _x_ is some number. Increase the number specified by the FILES command; the exact number of handles you enter depends on the applications you run and the number of files that are open at any given time. If other applications open multiple files, you may need to specify more file handles. As you increase the number of file handles, remember that each handle consumes more memory. For additional information, refer to your operating system manual.
    
3. Exit Microsoft Windows.
    
4. Reboot your system, and then try the operation again.
    

 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

