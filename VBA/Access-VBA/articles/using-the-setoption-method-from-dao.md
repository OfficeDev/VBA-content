---
title: Using the SetOption Method from DAO
keywords: acmain11.chm1032171
f1_keywords:
- acmain11.chm1032171
ms.prod: access
ms.assetid: 5a722d52-f41f-07a6-8197-7b73841a2fad
ms.date: 06/08/2017
---


# Using the SetOption Method from DAO

  

**Applies to:** Access 2013 | Access 2016

 Microsoft® Windows® Registry settings can be modified at run time with the **SetOption** method. To customize the Windows Registry settings, you can use the **SetOption** method from DAO. With this option, your application obtains the maximum flexibility and control. This approach allows you to create applications that are easier to maintain and that are tuned for maximum performance.

The syntax for doing this is dbEngine.SetOption, constant, NewValueSetting. For example, the following syntax, dbEngine.SetOption dbMaxLocksPerfFile, 20000, would allow the Microsoft Access database engine to track 20,000 locks at one time. The names of defined constants are the same as the registry name with db added as a prefix.
This is the recommended way to fine tune registry settings for your application. This method is the most flexible approach and provides you with the most control over how the registry is changed. With The  **SetOption** method you can specify new settings for any of the following default settings:


- PageTimeout key
    
- SharedAsyncDelay key
    
- ExclusiveAsyncDelay key
    
- LockRetry key
    
- UserCommitSync key
    
- ImplicitCommitSync key
    
- MaxBufferSize key
    
- MaxLocksPerFile key
    
- LockDelay key
    
- RecycleLVs
    
- FlushTransactionTimeout key
    

 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

