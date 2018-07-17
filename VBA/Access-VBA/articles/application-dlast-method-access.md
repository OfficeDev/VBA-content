---
title: Application.DLast Method (Access)
keywords: vbaac10.chm12530
f1_keywords:
- vbaac10.chm12530
ms.prod: access
api_name:
- Access.Application.DLast
ms.assetid: 0a04cbcc-0dbc-4cfc-e5a3-deb9b0f343be
ms.date: 06/08/2017
---


# Application.DLast Method (Access)

You can use the  **DLast** function to return a random record from a particular field in a table or query when you simply need any value from that field. .


## Syntax

 _expression_. **DLast**( ** _Expr_**, ** _Domain_**, ** _Criteria_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Expr_|Required|**String**|An expression that identifies the field from which you want to find the first or last value. It can be either a string expression identifying a field in a table or query, or an expression that performs a [calculation on data in that field](calculate-fields-in-domain-aggregate-functions.md). In  _expr_, you can include the name of a field in a table, a control on a form, a constant, or a function. If  _expr_ includes a function, it can be either built-in or user-defined, but not another domain aggregate or SQL aggregate function.|
| _Domain_|Required|**String**|A string expression identifying the set of records that constitutes the domain.|
| _Criteria_|Optional|**Variant**|An optional string expression used to restrict the range of data on which the  **DLast** function is performed. For example, _criteria_ is often equivalent to the WHERE clause in an SQL expression, without the wrd WHERE. If _criteria_ is omitted, the **DLast** function evaluates _expr_ against the entire domain. Any field that is included in _criteria_ must also be a field in _domain_; otherwise, the  **DLast** function returns a **Null**.|

### Return Value

Variant


## Remarks




 **Note**   If you want to return the first or last record in a set of records (a domain), you should create a query sorted as either ascending or descending and set the **TopValues** property to 1. From Visual Basic, you can also create an ADO **Recordset** object and use the **MoveFirst** or **MoveLast** method to return the first or last record in a set of records.


## Example



The following examples show how to use various types of criteria with the  **DLast** function.

 **Sample code provided by:**
![Community Member Icon](images/8b9774c4-6c97-470e-b3a2-56d8f786444c.png) The[UtterAccess](http://www.utteraccess.com) community




```vb
    ' ***************************
    ' Typical Use
    ' Numerical values. Replace "number" with the number to use.
    variable = DLast("[FieldName]", "TableName", "[Criteria] = number")

    ' Strings.
    ' Numerical values. Replace "string" with the string to use.
    variable = DLast("[FieldName]", "TableName", "[Criteria]= 'string'")

    ' Dates. Replace "date" with the string to use.
    variable = DLast("[FieldName]", "TableName", "[Criteria]= #date#")
    ' ***************************

    ' ***************************
    ' Referring to a control on a form
    ' Numerical values
    variable = DLast("[FieldName]", "TableName", "[Criteria] = " &; Forms!FormName!ControlName)

    ' Strings
    variable = DLast("[FieldName]", "TableName", "[Criteria] = '" &; Forms!FormName!ControlName &; "'")

    ' Dates
    variable = DLast("[FieldName]", "TableName", "[Criteria] = #" &; Forms!FormName!ControlName &; "#")
    ' ***************************

    ' ***************************
    ' Combinations
    ' Multiple types of criteria
    variable = DLast("[FieldName]", "TableName", "[Criteria1] = " &; Forms![FormName]![Control1] _
             &; " AND [Criteria2] = '" &; Forms![FormName]![Control2] &; "'" _
            &; " AND [Criteria3] =#" &; Forms![FormName]![Control3] &; "#")
    
    ' Use two fields from a single record.
    variable = DLast("[LastName] &; ', ' &; [FirstName]", "tblPeople", "[PrimaryKey] = 7")
            
    ' Expressions
    variable = DLast("[Field1] + [Field2]", "tableName", "[PrimaryKey] = 7")
    
    ' Control Structures
    variable = DLast("IIf([LastName] Like 'Smith', 'True', 'False')", "tableName", "[PrimaryKey] = 7")
    ' ***************************
```


## About the Contributors
<a name="AboutContributors"> </a>

UtterAccess is the premier Microsoft Access wiki and help forum. Click here to join. 


## See also
<a name="AboutContributors"> </a>


#### Concepts


[Application Object](application-object-access.md)

