---
title: Unique Property
keywords: acmain11.chm6173
f1_keywords:
- acmain11.chm6173
ms.prod: access
api_name:
- Access.Unique
ms.assetid: 283e5d33-b281-150f-9766-6ecc0da6a09a
ms.date: 06/08/2017
---


# Unique Property

  

**Applies to:** Access 2013 | Access 2016



You can use the Unique property to specify that an index enforces uniqueness of the data in the table's key index.

## Setting

The  **Unique** property uses the following settings.



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|Yes|**True** (-1)|The index is a key (unique) index.|
|No|**False** (0)|The index is a non-key index.|
You can set this property in the Indexes window of table Design view or by using Visual Basic.


 **Note**  To access the  **Unique** property of an index by using Visual Basic, use the ADO **Unique** property.


## Remarks

A key index optimizes finding records. It consists of one or more fields that uniquely arrange all records in a table in a predefined order. If the index consists of one field, values in that field must be unique. If the index consists of more than one field, duplicate values can occur in each field, but each combination of values from all the indexed fields must be unique. A non-key index has fields with values that are not necessarily unique. 

An index is the primary index for a table if its  **Primary** property is set to Yes. Each table can have only one primary index.

 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

