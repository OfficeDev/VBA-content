---
title: Required Property (Access)
keywords: vbaac10.chm7020
f1_keywords:
- vbaac10.chm7020
ms.prod: access
api_name:
- Access.Required
ms.assetid: 4aa6d0bb-4a07-9efc-4d2e-414bcc11e42e
ms.date: 06/08/2017
---


# Required Property (Access)

  

**Applies to:** Access 2013 | Access 2016

You can use the  **Required** property to specify whether a value is required in a field. If this property is set to Yes, when you enter data in a record, you must enter a value in the field or in any control bound to the field, and the value cannot be **Null**. For example, you might want to be sure that a LastName control has a value for each record. When you want to permit **Null** values in a field, you must not only set the **Required** property to No but, if there is a **ValidationRule** property setting, it must also explicitly state " _validationrule_ Or Is Null".


 **Note**  The  **Required** property doesn't apply to AutoNumber fields.


## Setting

The  **Required** property uses the following settings.



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|Yes|**True** (-1)|The field requires a value.|
|No|**False** (0)|(Default) The field doesn't require a value.|
You can set this property for all table fields (except AutoNumber data type fields) by using the table's property sheet or Visual Basic .


 **Note**  To access a field's  **Required** property in Visual Basic, use the DAO **Required** property.


## Remarks

The  **Required** property is enforced at the table level by the Microsoft Jet database engine. If you set this property to Yes, the field must receive or already contain a value when it has the focus — when a user enters data in a table (or in a form or datasheet based on the table), when a macro or Visual Basic sets the value of the field, or when data is imported into the table.

You can use the  **Required** and **AllowZeroLength** properties to differentiate between information that doesn't exist (stored as a zero-length string (" ") in the field) and information that may exist but is unknown (stored as a **Null** value in the field). If you set the **AllowZeroLength** property to Yes, a zero-length string will be a valid entry in the field regardless of the **Required** property setting. If you set **Required** to Yes and **AllowZeroLength** to No, you must enter a value in the field, and a zero-length string won't be a valid entry.


 **Tip**  You can use an input mask when data is entered in a field to distinguish between the display of a  **Null** value and a zero-length string. For example, the string "None" could be displayed when a zero-length string is entered.

The following table shows the results you can expect when you combine the settings of the  **Required** and **AllowZeroLength** properties.



|**Required**|**AllowZeroLength**|**User's action**|**Value stored**|
|:-----|:-----|:-----|:-----|
|No|No|Presses ENTER Presses SPACEBAR Enters a zero-length string|**Null** **Null** (not allowed)|
|No|Yes|Presses ENTER Presses SPACEBAR Enters a zero-length string|**Null** **Null** Zero-length string|
|Yes|No|Presses ENTER Presses SPACEBAR Enters a zero-length string|(not allowed) (not allowed) (not allowed)|
|Yes|Yes|Presses ENTER Presses SPACEBAR Enters a zero-length string|(not allowed) Zero-length string Zero-length string|
If you set the  **Required** property to Yes for a field in a table that already contains data, Microsoft Access gives you the option of checking whether the field has a value in all existing records. However, you can require that a value be entered in this field in all new records even if there are existing records with **Null** values in the field.


 **Note**  To enforce a relationship between related tables that don't allow  **Null** values, set the **Required** property of the foreign key field in the related table to Yes. The Jet database engine then ensures that you have a related record in the parent table before you can create a record in the child table. If the foreign key field is part of the primary key of the child table, this is unnecessary, because a primary key field can't contain a **Null** value.

 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

