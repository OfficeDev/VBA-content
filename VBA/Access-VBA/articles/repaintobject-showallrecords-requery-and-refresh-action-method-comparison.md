---
title: RepaintObject, ShowAllRecords, Requery, and Refresh Action/Method Comparison
keywords: vbaac10.chm5257635
f1_keywords:
- vbaac10.chm5257635
ms.prod: access
ms.assetid: ef1eec86-54d1-5b86-323f-48fb4f7d3897
ms.date: 06/08/2017
---


# RepaintObject, ShowAllRecords, Requery, and Refresh Action/Method Comparison

  

**Applies to:** Access 2013 | Access 2016

The following table provides a brief comparison of the RepaintObject action,  **RepaintObject** method, **Repaint** method, ShowAllRecords action, **ShowAllRecords** method, Requery action, **DoCmd.Requery** method, **Refresh** method, and **Requery** method.



|**Action or Method**|**Description**|
|:-----|:-----|
|RepaintObject action
 **DoCmd**.RepaintObject, **Repaint** method|Use the RepaintObject action,  **RepaintObject** method or **Repaint** method to repaint controls in the specified object. They don't requery the database or display new records.|
|ShowAllRecords action
 **ShowAllRecords** method|Use the ShowAllRecords action to requery and display the most recent records and remove any applied filters, which the Requery action doesn't do.|
|Requery action
 **DoCmd.Requery** method|Use the Requery action or method to requery the source of the object or one of its controls. The Requery action or method does one of the following: Reruns the query on which the control or object is based. Displays any new or changed records, and removes any deleted records from the table on which the control or object is based.|
|**Refresh** method|Use the  **Refresh** method to immediately update the records in the underlying record source for a specified form or datasheet to reflect changes made to the data by you and other users in a multiuser environment. The **Refresh** method shows only changes that have been made to the current set of records; it doesn't reflect new records or deleted records in the record source.|
|**Requery** method|
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>Use the <b>Requery</b>  method to update the data underlying a form or control to reflect records that are new to or have been deleted from the record source since it was last requeried. 



If you want to requery a control that isn't on the active object, you must use this method, not the Requery action or its corresponding <b>DoCmd.Requery</b>  method.

						</p></li></ul>|
 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

