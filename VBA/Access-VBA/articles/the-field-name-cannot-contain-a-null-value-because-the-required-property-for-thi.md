---
title: The field <name> cannot contain a Null value because the Required property for this field is set to True. Enter a value in this field. (Error 3314)
keywords: jeterr40.chm5003314
f1_keywords:
- jeterr40.chm5003314
ms.prod: access
ms.assetid: 451a9b22-e0ec-cb43-92d4-2f010086802c
ms.date: 06/08/2017
---


# The field <name> cannot contain a Null value because the Required property for this field is set to True. Enter a value in this field. (Error 3314)

  

**Applies to:** Access 2013 | Access 2016

 **In this article**
[What Is a Null?](#sectionSection0)
[What Is the Required Property?](#sectionSection1)
[Solution](#sectionSection2)
[To remove the required property](#sectionSection3)
[Setting the AllowZeroLength Property](#sectionSection4)


The  **Required** property for this field is set to **Yes**, prohibiting the entry of a **Null** or zero-length string (" ") in the field. Enter a value in the field.

## What Is a Null?
<a name="sectionSection0"> </a>

A  **Null** is a value you can enter in a field or use in expressions or queries to indicate missing or unknown data. In Microsoft Visual Basic, the **Null** keyword indicates a **Null** value. Some fields, such as primary key fields, cannot contain **Null**.


## What Is the Required Property?
<a name="sectionSection1"> </a>

You can use the  **Required** property to specify whether a value is required in a field. If this property is set to **Yes**, when you enter data in a record, you must enter a value in the field or in any control bound to the field, and the value cannot be **Null**. For example, you might want to be sure that a LastName control has a value for each record.


## Solution
<a name="sectionSection2"> </a>

This problem has the following possible solutions:


- Enter a value in the field named in the error message.
    
- Remove the  **Required** property setting from the field.
    
- Use the  **AllowZeroLength** property to allow zero-length strings (" ") to be stored in the field.
    

## To remove the required property
<a name="sectionSection3"> </a>


1. In the Navigation Pane, right-click the name of the table that contains the required field, and then click  **Design View**.
    
2. Click the  **Field Name** for the field named in the error message.
    
3. Under  **Field Properties**, click the  **General** tab.
    
4. In the  **Required** property, click **No**.
    
5. To save your changes, click  **Save** on the **Quick Access Toolbar**, or press CTRL+S.
    

## Setting the AllowZeroLength Property
<a name="sectionSection4"> </a>

You can use the  **Required** and **AllowZeroLength** properties to differentiate between information that does not exist (stored as a zero-length string (" ") in the field) and information that may exist but is unknown (stored as a **Null** value in the field). If you set the **AllowZeroLength** property to Yes, a zero-length string will be a valid entry in the field regardless of the **Required** property setting. If you set **Required** to **Yes** and **AllowZeroLength** to **No**, you must enter a value in the field, and a zero-length string will not be a valid entry.

The following table shows the results you can expect when you combine the settings of the  **Required** and **AllowZeroLength** properties.



|**Required**|**AllowZeroLength**|**User's action**|**Value stored**|
|:-----|:-----|:-----|:-----|
|No|No|Presses ENTER Presses SPACEBAR Enters a zero-length string|**Null** **Null** (not allowed)|
|No|Yes|Presses ENTER Presses SPACEBAR Enters a zero-length string|**Null** **Null** Zero-length string|
|Yes|No|Presses ENTER Presses SPACEBAR Enters a zero-length string|(not allowed) (not allowed) (not allowed)|
|Yes|Yes|Presses ENTER Presses SPACEBAR Enters a zero-length string|(not allowed) Zero-length string Zero-length string|
 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

