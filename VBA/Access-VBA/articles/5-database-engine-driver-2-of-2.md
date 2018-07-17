---
title: 5 Database Engine Driver (2 of 2)
keywords: acmain11.chm1032162
f1_keywords:
- acmain11.chm1032162
ms.prod: access
ms.assetid: 4d72622b-b956-4dd2-64cf-c0b17da0196e
ms.date: 06/08/2017
---


# 5 Database Engine Driver (2 of 2)

 
**Applies to:** Access 2013 | Access 2016

When you install the Microsoft® Jet version 2.5 Engine database driver, the Setup program writes a set of default values to the Microsoft Windows® Registry in the Engines and ISAM Formats subkeys. You must use the Registry Editor to add, remove, or change these settings. The following sections describe initialization and ISAM Format settings for the Microsoft Jet Engine database driver.


## Microsoft Jet Engine Initialization Settings

The  **Access Connectivity Engine\Engines\Jet 2.x** folder includes initialization settings for the Acer2x.dll driver, used for access to Microsoft Access 2.0 worksheets. Typical initialization settings for the entries in this folder are shown in the following example.


```
win32=<path>\ACER2X.DLL 

PageTimeout=5 

LockedPageTimeout=5 

CursorTimeout=5 

LockRetry=20 

CommitLockRetry=20 

MaxBufferSize=512 

ReadAheadPages=16 

IdleFrequency=10 

ForceOsFlush = 0
```

The following entries are used to configure the Microsoft Access database engine.



|**Entry**|**Description**|
|:-----|:-----|
|win32|Location of the database engine driver (.dll). The path is determined at the time of installation. Values are of type REG_SZ.|
|PageTimeout|The length of time between when data that is not read-locked is placed in an internal cache and when it is invalidated, expressed in 100 millisecond units. The default is 5 units (or 0.5 seconds). Values are of type REG_DWORD.|
|LockedPageTimeout|The length of time between when data that is read-locked is placed in an internal cache and when it is invalidated, expressed in 100 millisecond units. The default is 5 units (or 0.5 seconds). Values are of type REG_DWORD.|
|CursorTimeout|The length of time a reference to a page will remain on that page, expressed in 100 millisecond units. The default is 5 units (or 0.5 seconds). This setting applies only to databases created with version 1.x of the Microsoft Jet database engine. Values are of type REG_DWORD.|
|LockRetry|The number of times to repeat attempts to access a locked page before returning a lock conflict message. The default is 20 times; LockRetry is related to CommitLockRetry. Values are of type REG_DWORD.|
|CommitLockRetry|The number of times the Microsoft Jet database engine attempts to acquire a lock on data to commit changes to that data. If the Microsoft Jet database engine cannot acquire a commit lock, changes to the data will be unsuccessful. The number of attempts the Microsoft Jet database engine makes to get a commit lock is directly related to the LockRetry value. For each attempt made to acquire a commit lock, the Microsoft Jet database engine will make as many attempts as specified by the LockRetry value to acquire a lock. For example, if CommitLockRetry is set to 20 and LockRetry is set to 20, the Microsoft Access database engine will try to acquire a commit lock as many as 20 times; for each of those attempts, the Microsoft Access database engine will try to acquire a lock as many as 20 times, for a total of 400 attempts. The default value for CommitLockRetry is 20. Values are of type REG_DWORD.|
|MaxBufferSize|The size of the database engine internal cache, measured in kilobytes (K). MaxBufferSize must be a whole number value between 9 and 4096, inclusive. The default is 512. Values are of type REG_DWORD.|
|ReadAheadPages|The number of pages to read ahead when performing sequential scans. The default is 16. Values are of type REG_DWORD.|
|ForceOSFlush|Any setting other than 0 means a commit or a write will force flushing the OS cache to disk. A setting of 0 (the default setting) means no force flush occurs. Values are of type REG_DWORD.|
|IdleFrequency|The amount of time, in 100 millisecond units, that Microsoft Jet will wait before releasing a read lock. The default is 10 units or one second. Values are of type REG_DWORD.|

## Microsoft Jet Engine ISAM Formats

The  **Access Connectivity Engine\ISAM Formats\Jet 2.x** folder contains the following entries.



|**Entry name**|**Type**|**Value**|
|:-----|:-----|:-----|
|Engine|REG_SZ|Jet 2.x|
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

