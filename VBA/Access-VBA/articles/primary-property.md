---
title: Primary Property
keywords: vbaac10.chm4470
f1_keywords:
- vbaac10.chm4470
ms.prod: access
api_name:
- Access.Primary
ms.assetid: 782e3341-f47a-2054-9884-1feb2e68496c
ms.date: 06/08/2017
---


# Primary Property

  

**Applies to:** Access 2013 | Access 2016



You can use the Primary property to specify the primary key field for a table. A primary key field holds data that uniquely identifies each record in a table.

## Setting

The  **Primary** property uses the following settings.



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|Yes|**True**|The selected index is the primary key.|
|No|**False**|The selected index isn't the primary key.|
You can set the  **Primary** property in three ways:




- In table Design view, select the field or fields in the order you want for the primary key and then click  **Primary Key** on the toolbar.
    
- In the Indexes window, select or enter the name of an index in the  **Index Name** column and set the **Primary** property to Yes in the Index Properties section.
    
- In Visual Basic, to access the  **Primary** property of an index, use the DAO **PrimaryKey** property.
    



## Remarks

Microsoft Access automatically creates an index on the primary key field of a table and uses it to find records and to create joins between tables. The primary key index requires an entry in each primary key field and allows no duplicates. The order of the fields in a multiple-field primary key determines the default sort order for the table. 

If there is no primary key when you save a table's design, Microsoft Access will display a dialog box asking whether you want a primary key to be created. If you click  **Yes**, an AutoNumber data type field will be added to the table (with its **NewValues** property set to Increment) and set as the primary key. If you click **No**, no primary key will be created.

A table with no primary key can't be used in a relationship and can be slower to sort and search. 

 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

