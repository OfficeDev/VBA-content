---
title: UniqueRecords Property
keywords: vbaac10.chm4530
f1_keywords:
- vbaac10.chm4530
ms.prod: access
api_name:
- Access.UniqueRecords
ms.assetid: 05ab9d9b-d23f-2da3-9ba4-fa917e013dec
ms.date: 06/08/2017
---


# UniqueRecords Property

  

**Applies to:** Access 2013 | Access 2016

 **In this article**
[Setting](#sectionSection0)
[Remarks](#sectionSection1)
[Example](#sectionSection2)
[Customers table](#sectionSection3)
[Orders table](#sectionSection4)


You can use the  **UniqueRecords** property to specify whether to return only unique records based on all fields in the underlying data source, not just those fields present in the query itself.

 **Note**  The  **UniqueRecords** property applies only to append and make-table action queries and select queries.


## Setting
<a name="sectionSection0"> </a>

The  **UniqueRecords** property uses the following settings.



|**Setting**|**Description**|
|:-----|:-----|
|Yes|Doesn't return duplicate records.|
|No|(Default) Returns duplicate records.|
You can set the  **UniqueRecords** property in the query's property sheet or in SQL view of the Query window.


 **Note**  You set this property when you create a new query by using an SQL statement. The DISTINCTROW predicate corresponds to the  **UniqueRecords** property setting. The DISTINCT predicate corresponds to the **UniqueValues** property setting.


## Remarks
<a name="sectionSection1"> </a>

You can use the  **UniqueRecords** property when you want to omit data based on entire duplicate records, not just duplicate fields. Microsoft Access considers a record to be unique as long as the value in one field in the record differs from the value in the same field in another record.

The  **UniqueRecords** property has an effect only when you use more than one table in the query and select fields from the tables used in the query. The **UniqueRecords** property is ignored if the query includes only one table.

The  **UniqueRecords** and **UniqueValues** properties are related in that only one of them can be set to Yes at a time. When you set **UniqueRecords** to Yes, for example, Microsoft Access automatically sets **UniqueValues** to No. You can, however, set both of them to No. When both properties are set to No, all records are returned.


## Example
<a name="sectionSection2"> </a>

The query in this example returns a list of customers from the Customers table who have at least one order in the Orders table.


## Customers table
<a name="sectionSection3"> </a>



|**Company name**|**Customer ID**|
|:-----|:-----|
|Ernst Handel|ERNSH|
|Familia Arquibaldo|FAMIA|
|FISSA Fabrica Inter. Salchichas S.A.|FISSA|
|Folies gourmandes|FOLIG|

## Orders table
<a name="sectionSection4"> </a>



|**Customer ID**|**Order ID**|
|:-----|:-----|
|ERNSH|10698|
|FAMIA|10512|
|FAMIA|10725|
|FOLIG|10763|
|FOLIG|10408|
The following SQL statement returns the customer names in the following table:


```sql
SELECT DISTINCTROW Customers.CompanyName, Customers.CustomerID 
FROM Customers INNER JOIN Orders 
ON Customers.CustomerID = Orders.CustomerID; 
 
```



|**Customers returned**|**Customer ID**|
|:-----|:-----|
|Ernst Handel|ERNSH|
|Familia Arquibaldo|FAMIA|
|Folies gourmandes|FOLIG|
 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

