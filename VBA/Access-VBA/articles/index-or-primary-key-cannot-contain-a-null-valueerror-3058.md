---
title: Index or primary key cannot contain a Null value. (Error 3058)
ms.prod: access
ms.assetid: ec435ace-b33a-b14d-54ce-fd918666ee53
ms.date: 06/08/2017
---


# Index or primary key cannot contain a Null value. (Error 3058)

  

**Applies to:** Access 2013 | Access 2016

 **In this article**
[What Is a Primary Key?](#sectionSection0)
[What Is a Null?](#sectionSection1)
[Solution](#sectionSection2)


Possible causes:


- You tried to add a new record but did not enter a value in the field that contains the primary key.
    
- You tried to add a  **Null** value to a primary key field.
    
- You executed a query that tried to put a  **Null** value in a primary key field.
    


## What Is a Primary Key?
<a name="sectionSection0"> </a>

A primary key is a field or set of fields in your table that provide Microsoft Access with a unique identifier for every row. In a relational database, such as an Access database, you divide your information into separate, subject-based tables. You then use table relationships and primary keys to tell Access how to bring the information back together again. Access uses primary key fields to quickly associate data from multiple tables and combine that data in a meaningful way.

Often, a unique identification number, such as an ID number or a serial number or code, serves as a primary key in a table. For example, you might have a Customers table where each customer has a unique customer ID number. The customer ID field is the primary key.

An example of a poor choice for a primary key would be a name or address. Both contain information that might change over time.

Access ensures that every record has a value in the primary key field, and that the value is always unique.


## What Is a Null?
<a name="sectionSection1"> </a>

A  **Null** is a value you can enter in a field or use in expressions or queries to indicate missing or unknown data. In Microsoft Visual Basic, the **Null** keyword indicates a **Null** value. Some fields, such as primary key fields, cannot contain **Null**.


## Solution
<a name="sectionSection2"> </a>

To solve this problem, you must enter a value in the primary key field before moving to another record.

 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

