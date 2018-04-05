---
title: 0 Database Engine Driver
keywords: acmain11.chm1032161
f1_keywords:
- acmain11.chm1032161
ms.prod: access
ms.assetid: cff53f53-5848-72f7-82b0-e600e82bd3de
ms.date: 06/08/2017
---


# 0 Database Engine Driver


**Applies to:** Access 2013 | Access 2016

When you install the Microsoft® Access database engine database driver, the Setup program writes a set of default values to the Microsoft Windows® Registry in the Engines and ISAM Formats subkeys. You must use the Registry Editor to add, remove, or change these settings. The following sections describe initialization and ISAM Format settings for the Microsoft Access Database Engine database driver.


## Microsoft Jet Engine Initialization Settings

The  **Access Connectivity Engine\Engines** folder includes initialization settings for the msjet40.dll database engine, used for access to Microsoft Access databases. Typical initialization settings for the entries in this folder are shown in the following example.


```
SystemDB = <path>\System.mdb 

CompactBYPkey = 1 

PrevFormatCompactWithUNICODECompression=1
```

The Microsoft Access database engine uses the following entries.



|**Entry**|**Description**|
|:-----|:-----|
|SystemDB|Specifies the full path and file name of the workgroup information file. The default is the appropriate path followed by the file name System.mdb. Values are of type REG_SZ.|
|CompactByPKey|Specifies that when you compact tables they are copied in primary-key order, if a primary key exists on the table. If no primary key exists on a table, the tables are copied in base-table order. A value of 0 indicates that tables should be compacted in base-table order; a non-zero value indicates that tables should be compacted in primary-key order, if a primary key exists. The default value is non-zero. Values are of type REG_DWORD.|
|PrevFormatCompactWithUNICODECompression|Microsoft Access database engine databases use the Unicode character set to store textual data. Compressing the Unicode data can significantly improve the performance of the database because of the reduced number of page read/write operations that are needed afterwards. This key determines if databases created by the Microsoft Jet database engine version 3.x or earlier should be created with compressed Unicode or un-compressed Unicode. **Note**  This setting does not apply to compacting Microsoft Access database engine databases databases. Microsoft Access database engine databases databases will default to keep the compression settings with which they were created.|

The **Access Connectivity Engine\Engines\ACE** folder includes initialization settings for the Ace.dll database engine, used for access to Microsoft Access databases. Typical initialization settings for the entries in this folder are shown in the following example.




```
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

PagesLockedToTableLock=0
```

The Microsoft Access database engine uses the following entries.



|**Entry**|**Description**|
|:-----|:-----|
|PageTimeout|The length of time between the time when data that is not read-locked is placed in an internal cache and the time when it is invalidated, expressed in milliseconds. The default is 5000 milliseconds or 5 seconds. Values are of type REG_DWORD.|
|FlushTransactionTimeout|This entry disables both the ExclusiveAsyncDelay and SharedAsyncDelay registry entries. To enable those entries, a value of zero must be entered. FlushTransactionTimeout changes the Microsoft Access database engine's method for doing asynchronous writes to a database file. |
|LockDelay|This setting works in conjunction with the LockRetry setting in that it causes each LockRetry to wait 100 milliseconds before issuing another lock request. The LockDelay setting was added to prevent bursting that would occur with certain networking operating systems.|
|MaxLocksPerFile|This setting prevents transactions in the Microsoft Access database engine from exceeding the specified value. If the locks in a transaction attempt to exceed this value, then the transaction is split into two or more parts and partially committed. This setting was added to prevent Netware 3.1 server crashes when the specified Netware lock limit was exceeded, and to improve performance with both Netware and NT.|
|LockRetry|The number of times to repeat attempts to access a locked page before returning a lock conflict message. The default is 20. Values are of type REG_DWORD.|
|RecycleLVs|This setting, when enabled, will cause the Microsoft Access database engine to recycle long value (LV) pages (Memo, Long Binary [OLE object], and Binary data types). Values are of type REG_DWORD.|
|MaxBufferSize|The size of the database engine internal cache, measured in kilobytes (K). MaxBufferSize must be an integer value greater than or equal to 512. The default is based on the following formula: `((TotalRAM in MB - 12 MB) / 4) + 512 KB` For example, on a system with 32 MB of RAM, the default buffer size is ((32 MB - 12 MB) / 4) + 512 KB or 5632 KB. To set the value to the default, set the registry key to `MaxBufferSize=` Values are of type REG_DWORD.|
|Threads|The number of background threads available to the Microsoft Access database engine. The default is 3. Values are of type REG_DWORD.|
|UserCommitSync|Specifies whether the system waits for a commit to finish. A value of Yes instructs the system to wait; a value of No instructs the system to perform the commit asynchronously. The default is Yes. Values are of type REG_SZ.|
|ImplicitCommitSync|Specifies whether the system waits for a commit to finish. A value of No instructs the system to proceed without waiting for the commit to finish; a value of Yes instructs the system to wait for the commit to finish. The default is No. Values are of type REG_SZ.|
|ExclusiveAsyncDelay|Specifies the length of time, in milliseconds, to defer an asynchronous flush of an exclusive database. The default value is 2000, or 2 seconds. Values are of type REG_DWORD.|
|SharedAsyncDelay|Specifies the length of time, in milliseconds, to defer an asynchronous flush of a shared database. The default value is 0. Values are of type REG_DWORD.|
|PagesLockedToTableLock|During bulk operations it is often more efficient to lock a whole table, instead of obtaining locks for each individual page of the table as you try to access it. This setting specifies the number of pages that the Microsoft Access database engine will allow to be locked in any particular transaction before the Access database engine attempts to escalate to an exclusive table lock The default value of 0 indicates that the Access database engine will never automatically change from page locking to table locking.|



> **Note** This setting should be used carefully. If a database is needed for multi-user access, then locking a whole table could cause locking conflicts for other users. This would be especially severe if a small number was used for this setting. Even when a larger number was used, such as 25 or 50, the operation for other users might become unpredictable.

> **Note** When you change Windows Registry settings, you must exit and then restart the database engine for the new settings to take effect.

## Additional resources
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

