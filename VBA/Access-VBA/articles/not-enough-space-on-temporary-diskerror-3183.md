---
title: Not enough space on temporary disk. (Error 3183)
keywords: jeterr40.chm5003183
f1_keywords:
- jeterr40.chm5003183
ms.prod: access
ms.assetid: ba122b0f-2445-705c-f24d-810ebc9ddeb9
ms.date: 06/08/2017
---


# Not enough space on temporary disk. (Error 3183)

  

**Applies to:** Access 2013 | Access 2016

You tried to perform an operation that requires more space than is available on the temporary disk. Your temporary disk location is based on the TEMP DOS environment variable, which was set when your system started.

For example, you may be trying to create a query that creates temporary files larger than the temporary disk. Reduce the size of the temporary files by accessing smaller amounts of data at one time or increase the size of the temporary disk.
You can increase the amount of available temporary disk space in several ways:


- Select fewer records. Dynaset-type, forward-only — type, and snapshot-type  **Recordset** objects record keys and data to the temporary disk.
    
- Select a different drive for your temporary disk.
    
- If the temporary disk is a RAM disk, increase the amount of available RAM and the space allocated to the RAM disk, or move it to a fixed disk.
    
- Free some space by deleting data or by removing unneeded tables, queries, forms, macros, and modules from your database.
    
- Free some space by compressing deleted records out of your database.
    
- If you still need additional space, consider removing other unused files from your disk.
    

 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

