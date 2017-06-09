---
title: 5 Database Engine Driver (1 of 2)
keywords: acmain11.chm1032160
f1_keywords:
- acmain11.chm1032160
ms.prod: access
ms.assetid: dcc90f49-3674-8f95-ab25-82513f7e2cfa
ms.date: 06/08/2017
---


# 5 Database Engine Driver (1 of 2)

 
**Applies to:** Access 2013 | Access 2016

When you install the Microsoft® Jet version 3.5 Engine database driver, the Setup program writes a set of default values to the Microsoft Windows® Registry in the Engines and ISAM Formats subkeys. You must use the Registry Editor to add, remove, or change these settings. The following sections describe initialization and ISAM Format settings for the Microsoft Jet Engine database driver.


## Microsoft Jet Engine Initialization Settings

The  **Access Connectivity Engine\Engines\Jet 3.x** folder includes initialization settings for the Acer3x.dll driver, used for access to Microsoft Access 97 worksheets. Typical initialization settings for the entries in this folder are shown in the following example.


```
win32=<path>\ Acer3x.dll

FlushTransactionTimeout=500 

LockDelay=100 

LockRetry=20 

MaxBufferSize= 0 

MaxLocksPerFile= 9500 

PageTimeout=5000 

Threads=3 

UserCommitSync=Yes 

ImplicitCommitSync=No 

ExclusiveAsyncDelay=2000 

SharedAsyncDelay=0 

RecycleLVs=0 

SortMemorySource=0
```

The Microsoft Access database engine uses the following entries.



|**Entry**|**Description**|
|:-----|:-----|
|win32|Location of the database engine driver (.dll). The path is determined at the time of installation. Values are of type REG_SZ.|
|PageTimeout|The length of time between the time when data that is not read-locked is placed in an internal cache and when it is invalidated, expressed in milliseconds. The default is 5000 milliseconds or 5 seconds. Values are of type REG_DWORD.|
|FlushTransactionTimeout|This entry disables both the ExclusiveAsyncDelay and SharedAsyncDelay registry entries. To enable those entries, a value of zero must be entered. FlushTransactionTimeout changes the Microsoft Jet database engine's method for doing asynchronous writes to a database file. Previously, the Microsoft Jet database engine would use either the ExclusiveAsyncDelay or SharedAsyncDelay to determine how long it would wait before forcing asynchronous writes. FlushTransactionTimeout changes that behavior by having a value that will start asynchronous writes only after the specified amount of time has expired and no pages have been added to the cache. The only exception to this is if the cache exceeds the MaxBufferSize, at which point the cache will start asynchronous writing regardless if the time has expired. Microsoft Jet 3.5 database engine will wait 500 milliseconds of non-activity or until the cache size is exceeded before starting asynchronous writes.|
|LockDelay|This setting works in conjunction with the LockRetry setting in that it causes each LockRetry to wait 100 milliseconds before issuing another lock request. The LockDelay setting was added to prevent "bursting" that would occur with certain networking operating systems.|
|MaxLocksPerFile|This setting prevents transactions in Microsoft Jet from exceeding the specified value. If the locks in a transaction attempts to exceed this value, then the transaction is split into two or more parts and partially committed. This setting was added to prevent Netware 3.1 server crashes when the specified Netware lock limit was exceeded and to improve performance with both Netware and NT.|
|LockRetry|The number of times to repeat attempts to access a locked page before returning a lock conflict message. The default is 20. Values are of type REG_DWORD.|
|RecycleLVs|This setting, when enabled, will cause Microsoft Jet to recycle long value (LV) pages (Memo, Long Binary [OLE object], and Binary data types). Microsoft Jet 3.0 would not recycle those types of pages until the last user closed the database. If the RecyleLVs setting is enabled, Microsoft Jet 3.5 will start to recycle most LV pages when the database is expanded (that is, when groups of pages are added).
 **Note**  By enabling this feature, users will notice a performance degradation when manipulating long value data types. Microsoft Access 97 automatically enables and disables this feature when manipulating modules, forms, and reports, thus eliminating the need to turn it on when modifying those objects. The default value is 0. Values are of type REG_DWORD.

|
|MaxBufferSize|The size of the database engine internal cache, measured in kilobytes (K). MaxBufferSize must be an integer value greater than or equal to 512. The default is based on the following formula: `((TotalRAM in MB - 12 MB) / 4) + 512 KB` For example, on a system with 32 MB of RAM, the default buffer size is ((32 MB - 12 MB) / 4) + 512 KB or 5632 KB. To set the value to the default, set the registry key to `MaxBufferSize=` Values are of type REG_DWORD.|
|Threads|The number of background threads available to the Microsoft Jet database engine. The default is 3. Values are of type REG_DWORD.|
|UserCommitSync|Specifies whether the system waits for a commit to finish. A value of Yes instructs the system to wait; a value of No instructs the system to perform the commit asynchronously. The default is Yes. Values are of type REG_DWORD.|
|ImplicitCommitSync|Specifies whether the system waits for a commit to finish. A value of No instructs the system to proceed without waiting for the commit to finish; a value of Yes instructs the system to wait for the commit to finish. The default is No. Values are of type REG_DWORD.|
|ExclusiveAsyncDelay|Specifies the length of time, in milliseconds, to defer an asynchronous flush of an exclusive database. The default value is 2000 or 2 seconds. Values are of type REG_DWORD.|
|SharedAsyncDelay|Specifies the length of time, in milliseconds, to defer an asynchronous flush of a shared database. The default value is 0. Values are of type REG_DWORD.|
|SortMemorySource|Specifies how Microsoft Jet obtains the memory that is used for sort keys. A value of 0 indicates that memory should be taken from the heap. A value of 1 indicates that memory should be taken from global memory using the malloc function call.|

## Microsoft Jet Engine ISAM Formats

The  **Jet\4.0\ISAM Formats\Jet 3.x** folder contains the following entries.



|**Entry name**|** Type**|**Value**|
|:-----|:-----|:-----|
|Engine|REG_SZ|Jet 3.x|
|OneTablePerFile|REG_BINARY|00|
|IndexDialog|REG_BINARY|00|
|CreateDBOnExport|REG_BINARY|00|
|IsamType|REG_DWORD|0|

 **Note**  When you change Windows Registry settings, you must exit and then restart the database engine for the new settings to take effect.

 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

