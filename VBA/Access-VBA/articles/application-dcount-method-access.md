---
title: Application.DCount Method (Access)
keywords: vbaac10.chm12536
f1_keywords:
- vbaac10.chm12536
ms.prod: access
api_name:
- Access.Application.DCount
ms.assetid: 257f0b2a-e23d-2728-afd2-7700b59e5456
ms.date: 06/08/2017
---


# Application.DCount Method (Access)

You can use the  **DCount** function to determine the number of records that are in a specified set of records (a domain).


## Syntax

 _expression_. **DCount**( ** _Expr_**, ** _Domain_**, ** _Criteria_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Expr_|Required|**String**|An expression that identifies the field for which you want to count records. It can be a string expression identifying a field in a table or query, or it can be an expression that performs a calculation on data in that field. In  _expr_, you can include the name of a field in a table, a control on a form, a constant, or a function. If  _expr_ includes a function, it can be either built-in or user-defined, but not another domain aggregate or SQL aggregate function.|
| _Domain_|Required|**String**|A string expression identifying the set of records that constitutes the domain. It can be a table name or a query name for a query that does not require a parameter.|
| _Criteria_|Optional|**Variant**|An optional string expression used to restrict the range of data on which the  **DCount** function is performed. For example, _criteria_ is often equivalent to the WHERE clause in an SQL expression, without the word WHERE. If _criteria_ is omitted, the **DCount** function evaluates _expr_ against the entire domain. Any field that is included in _criteria_ must also be a field in _domain_; otherwise the  **DCount** function returns a **Null**.|

### Return Value

Variant


## Remarks

For example, you could use the  **DCount** function in a module to return the number of records in an Orders table that correspond to orders placed on a particular date.

Use the  **DCount** function to count the number of records in a domain when you don't need to know their particular values. Although the _expr_ argument can perform a calculation on a field, the **DCount** function simply tallies the number of records. The value of any calculation performed by _expr_ is unavailable.

Use the  **DCount** function in a calculated control when you need to specify criteria to restrict the range of data on which the function is performed. For example, to display the number of orders to be shipped to California, set the **ControlSource** property of a text box to the following expression:




```
=DCount("[OrderID]", "Orders", "[ShipRegion] = 'CA'")
```

If you simply want to count all records in  _domain_ without specifying any restrictions, use the **Count** function.

 The **Count** function has been optimized to speed counting of records in queries. Use the **Count** function in a query expression instead of the **DCount** function, and set optional criteria to enforce any restrictions on the results. Use the **DCount** function when you must count records in a domain from within a code module or macro, or in a calculated control.

You can use the  **DCount** function to count the number of records containing a particular field that isn't in the record source on which your form or report is based. For example, you could display the number of orders in the Orders table in a calculated control on a form based on the Products table.

The  **DCount** function doesn't count records that contain **Null** values in the field referenced by _expr_ unless _expr_ is the asterisk (*) wildcard character. If you use an asterisk, the **DCount** function calculates the total number of records, including those that contain **Null** fields. The following example calculates the number of records in an Orders table.




```
intX = DCount("*", "Orders")
```

If  _domain_ is a table with a primary key, you can also count the total number of records by setting _expr_ to the primary key field, since there will never be a **Null** in the primary key field.

If  _expr_ identifies multiple fields, separate the field names with a concatenation operator, either an ampersand (&;) or the addition operator (+). If you use an ampersand to separate the fields, the **DCount** function returns the number of records containing data in any of the listed fields. If you use the addition operator, the **DCount** function returns only the number of records containing data in all of the listed fields. The following example demonstrates the effects of each operator when used with a field that contains data in all records (ShipName) and a field that contains no data (ShipRegion).




```
intW = DCount("[ShipName]", "Orders") 
intX = DCount("[ShipRegion]", "Orders") 
intY = DCount("[ShipName] + [ShipRegion]", "Orders") 
intZ = DCount("[ShipName] &; [ShipRegion]", "Orders")
```


 **Note**   The ampersand is the preferred operator for performing string concatenation. You should avoid using the addition operator for anything other than numeric addition, unless you specifically wish to propagate **Nulls** through an expression.


## Example

The following function returns the number of orders shipped to a specified country or region after a specified ship date. The domain is an Orders table.


```vb
Public Function OrdersCount(ByVal strCountry As String, _ 
                            ByVal dteShipDate As Date) As Integer 
 
    OrdersCount = DCount("[ShippedDate]", "Orders", _ 
                  "[ShipCountry] = '" &; strCountry &; _ 
                  "' AND [ShippedDate] > #" &; dteShipDate &; "#") 
End Function
```



The following examples show how to use various types of criteria with the  **DCount** function.

 **Sample code provided by:**
![Community Member Icon](images/8b9774c4-6c97-470e-b3a2-56d8f786444c.png) The[UtterAccess](http://www.utteraccess.com) community




```js
    ' ***************************
    ' Typical Use
    ' Numerical values. Replace "number" with the number to use.
    variable = DCount("[FieldName]", "TableName", "[Criteria] = number")

    ' Strings.
    ' Numerical values. Replace "string" with the string to use.
    variable = DCount("[FieldName]", "TableName", "[Criteria]= 'string'")

    ' Dates. Replace "date" with the string to use.
    variable = DCount("[FieldName]", "TableName", "[Criteria]= #date#")
    ' ***************************

    ' ***************************
    ' Referring to a control on a form
    ' Numerical values
    variable = DCount("[FieldName]", "TableName", "[Criteria] = " &; Forms!FormName!ControlName)

    ' Strings
    variable = DCount("[FieldName]", "TableName", "[Criteria] = '" &; Forms!FormName!ControlName &; "'")

    ' Dates
    variable = DCount("[FieldName]", "TableName", "[Criteria] = #" &; Forms!FormName!ControlName &; "#")
    ' ***************************

    ' ***************************
    ' Combinations
    ' Multiple types of criteria
    variable = DCount("[FieldName]", "TableName", "[Criteria1] = " &; Forms![FormName]![Control1] _
             &; " AND [Criteria2] = '" &; Forms![FormName]![Control2] &; "'" _
            &; " AND [Criteria3] =#" &; Forms![FormName]![Control3] &; "#")
    
    ' Use two fields from a single record.
    variable = DCount("[LastName] &; ', ' &; [FirstName]", "tblPeople", "[PrimaryKey] = 7")
            
    ' Expressions
    variable = DCount("[Field1] + [Field2]", "tableName", "[PrimaryKey] = 7")
    
    ' Control Structures
    variable = DCount("IIf([LastName] Like 'Smith', 'True', 'False')", "tableName", "[PrimaryKey] = 7")
    ' ***************************
```


## About the Contributors
<a name="AboutContributors"> </a>

UtterAccess is the premier Microsoft Access wiki and help forum. Click here to join. 


## See also
<a name="AboutContributors"> </a>


#### Concepts


[Application Object](application-object-access.md)

