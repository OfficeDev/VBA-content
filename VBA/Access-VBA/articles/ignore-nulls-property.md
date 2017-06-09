---
title: Ignore Nulls Property
keywords: acmain11.chm7025
f1_keywords:
- acmain11.chm7025
ms.prod: access
ms.assetid: 87d95ca8-ea29-f0ca-366a-56527c500f13
ms.date: 06/08/2017
---


# Ignore Nulls Property

  

**Applies to:** Access 2013 | Access 2016



You can use the IgnoreNulls property to specify that records with Null values in the indexed fields not be included in the index.

## Settings



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|Yes|**True**|Records that contain Null values in the indexed fields aren't included in the index.|
|No|**False**|(Default) Records that contain Null values in the indexed fields are included in the index.|
You can set this property by using the Indexes window of table Design view or Visual Basic.

To access the  **Ignore Nulls** property of an index by using Visual Basic, use the DAO **IgnoreNulls** property.

You can define an index for a field to facilitate faster searches for records indexed on that field. If you allow  **Null** entries in the indexed field and expect to have many of them, set the **Ignore Nulls** property for the index to Yes to reduce the amount of storage space that the index uses.

 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

