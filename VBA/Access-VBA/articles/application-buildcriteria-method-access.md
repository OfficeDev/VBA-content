---
title: Application.BuildCriteria Method (Access)
keywords: vbaac10.chm12548
f1_keywords:
- vbaac10.chm12548
ms.prod: access
api_name:
- Access.Application.BuildCriteria
ms.assetid: 098e9aca-3dc1-ad21-4374-5d8ae7c80c56
ms.date: 06/08/2017
---


# Application.BuildCriteria Method (Access)

The  **BuildCriteria** method returns a parsed criteria string as it would appear in the query design grid, in Filter By Form or Server Filter By Form mode. For example, you may want to set a form's **Filter** or **[ServerFilter](form-serverfilter-property-access.md)** property based on varying criteria from the user. You can use the **BuildCriteria** method to construct the string expression argument for the **Filter** or **ServerFilter** property. **String**.


## Syntax

 _expression_. **BuildCriteria**( ** _Field_**, ** _FieldType_**, ** _Expression_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Field_|Required|**String**|The field for which you wish to define criteria.|
| _FieldType_|Required|**Integer**|An intrinsic constant denoting the data type of the field. Can be set to one of the DAO  **DataTypeEnum** values.|
| _Expression_|Required|**String**|A string expression identifying the criteria to be parsed.|

### Return Value

String


## Remarks

The  **BuildCriteria** method enables you to easily construct criteria for a filter based on user input. It parses the _expression_ argument in the same way that the expression would be parsed had it been entered in the query design grid, in Filter By Form or Server Filter By Form mode.

For example, a user creating a query on an Orders table might restrict the result set to orders placed after January 1, 1995, by setting criteria on an OrderDate field. The user might enter an expression such as the following one in the  **Criteria** row beneath the OrderDate field:

>1-1-95

Microsoft Access automatically parses this expression and returns the following expression:

>#1/1/95#

The  **BuildCriteria** method provides the same parsing from Visual Basic code. For example, to return the preceding correctly parsed string, you can supply the following arguments to the **BuildCriteria** method:




```vb
Dim strCriteria As String 
strCriteria = BuildCriteria("OrderDate", dbDate, ">1-1-95")
```

Since you need to supply criteria for the  **Filter** property in correctly parsed form, you can use the **BuildCriteria** method to construct a correctly parsed string.

You can use the  **BuildCriteria** method to construct a string with multiple criteria if those criteria refer to the same field. For example, you can use the **BuildCriteria** method with the following arguments to construct a string with multiple criteria relating to the OrderDate field:




```vb
strCriteria = BuildCriteria("OrderDate", dbDate, ">1-1-95 and <5-1-95")
```

This example returns the following criteria string:

```text
OrderDate>#1/1/95# And OrderDate<#5/1/95#
```

However, if you wish to construct a criteria string that refers to multiple fields, you must create the strings and concatenate them yourself. For example, if you wish to construct criteria for a filter to show records for orders placed after 1-1-95 and for which freight is less than $50, you would need to use the  **BuildCriteria** method twice and concatenate the resulting strings.


## Example

The following example prompts the user to enter the first few letters of a product's name and then uses the  **BuildCriteria** method to construct a criteria string based on the user's input. Next, the procedure provides this string as an argument to the **Filter** property of a Products form. Finally, the **FilterOn** property is set to apply the filter.


```vb
Sub SetFilter() 
 Dim frm As Form, strMsg As String 
 Dim strInput As String, strFilter As String 
 
 ' Open Products form in Form view. 
 DoCmd.OpenForm "Products" 
 ' Return Form object variable pointing to Products form. 
 Set frm = Forms!Products 
 strMsg = "Enter one or more letters of product name " _ 
 &; "followed by an asterisk." 
 ' Prompt user for input. 
 strInput = InputBox(strMsg) 
 ' Build criteria string. 
 strFilter = BuildCriteria("ProductName", dbText, strInput) 
 ' Set Filter property to apply filter. 
 frm.Filter = strFilter 
 ' Set FilterOn property; form now shows filtered records. 
 frm.FilterOn = True 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-access.md)

