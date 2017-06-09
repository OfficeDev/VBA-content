---
title: TopValues Property
keywords: vbaac10.chm4525
f1_keywords:
- vbaac10.chm4525
ms.prod: access
api_name:
- Access.TopValues
ms.assetid: 86198e46-2061-f39f-b6cf-58b90ef063b7
ms.date: 06/08/2017
---


# TopValues Property

  

**Applies to:** Access 2013 | Access 2016

 **In this article**
[Setting](#sectionSection0)
[Remarks](#sectionSection1)
[Example](#sectionSection2)


You can use the  **TopValues** property to return a specified number of records or a percentage of records that meet the criteria you specify. For example, you might want to return the top 10 values or the top 25 percent of all values in a field.

 **Note**  The  **TopValues** property applies only to append, make-table, and select queries.


## Setting
<a name="sectionSection0"> </a>

The  **TopValues** property setting is an Integer value that represents the exact number of values to return or a number followed by a percent sign (%) that represents the percent of records to return. For example, to return the top 10 values, set the **TopValues** property to 10; to return the top 10 percent of values, set the **TopValues** property to 10%.

You cannot set this property in code directly. It is set in SQL view of the Query window by using a TOP n or TOP n PERCENT clause in the SQL statement.

You can also set the  **TopValues** property by using the query property sheet or the **Return** box in the **Query Setup** group on the ribbon.


 **Note**  The  **TopValues** property in the query property sheet **Return** box in the **Query Setup** group on the ribbon is a combo box that contains a list of values and percentage values. You can select one of these values or you can type any valid setting in the text box portion of this control.


## Remarks
<a name="sectionSection1"> </a>

Typically, you use the  **TopValues** property setting together with sorted fields. The field you want to display top values for should be the leftmost field that has the Sort box selected in the query design grid. An ascending sort returns the bottom-most records, and a descending sort returns the topmost records. If you specify that a specific number of records be returned, all records with values that match the value in the last record are also returned.

For example, suppose a set of employees has the following sales totals.



|**Sales**|**Salesperson**|
|:-----|:-----|
|90,000|Leverling|
|80,000|Peacock|
|70,000|Davilio|
|70,000|King|
|60,000|Suyama|
|50,000|Buchanan|
If you set the  **TopValues** property to 3 with a descending sort on the Sales field, Microsoft Access returns the following four records.



|**Sales**|**Salesperson**|
|:-----|:-----|
|90,000|Leverling|
|80,000|Peacock|
|70,000|Davilio|
|70,000|King|

 **Note**  To return the topmost or bottommost values without displaying duplicate values, set the  **UniqueValues** property in the query property sheet to Yes.


## Example
<a name="sectionSection2"> </a>

The following code example assigns an SQL string that returns the top 10 most expensive products to the  **RecordSource** property for a form that will display the ten most expensive products.


```sql
Dim strGetSQL As String 
strGetSQL = "SELECT TOP 10 Products.[ProductName] " _ 
    &; "AS TenMostExpensiveProducts, Products.UnitPrice FROM Products " _ 
    &; "ORDER BY Products.[UnitPrice] DESC;" 
Me.RecordSource = strGetSQL  

```

 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

