---
title: Application.Nz Method (Access)
keywords: vbaac10.chm12554
f1_keywords:
- vbaac10.chm12554
ms.prod: access
api_name:
- Access.Application.Nz
ms.assetid: 669fe962-3881-83bb-cc40-ec9b23b44116
ms.date: 06/08/2017
---


# Application.Nz Method (Access)

You can use the  **Nz** function to return zero, a zero-length string (" "), or another specified value when a **Variant** is **Null**. For example, you can use this function to convert a **Null** value to another value and prevent it from propagating through an expression.


## Syntax

 _expression_. **Nz**( ** _Value_**, ** _ValueIfNull_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Value_|Required|**Variant**|A variable of data type **Variant**.|
| _ValueIfNull_|Optional|**Variant**|Optional (unless used in a query). A  **Variant** that supplies a value to be returned if the variant argument is **Null**. This argument enables you to return a value other than zero or a zero-length string.<table><tr><th>**Note**</th></tr><tr><td>If you use the  **Nz** function in an expression in a query without using the _valueifnull_ argument, the results will be a zero-length string in the fields that contain null values.</td></tr></table>|

### Return Value

Variant


## Remarks

If the value of the variant argument is  **Null**, the **Nz** function returns the number zero or a zero-length string (always returns a zero-length string when used in a query expression), depending on whether the context indicates the value should be a number or a string. If the optional valueifnull argument is included, then the **Nz** function will return the value specified by that argument if the variant argument is **Null**. When used in a query expression, the **Nz** function should always include the valueifnull argument.

If the value of variant isn't  **Null**, then the **Nz** function returns the value of variant.

The  **Nz** function is useful for expressions that may include **Null** values. To force an expression to evaluate to a non- **Null** value even when it contains a **Null** value, use the **Nz** function to return zero, a zero-length string, or a custom return value.

For example, the expression  `2 + varX` will always return a **Null** value when the **Variant** `varX` is **Null**. However, `2 + Nz(varX)` returns 2.

You can often use the  **Nz** function as an alternative to the **IIf** function. For example, in the following code, two expressions including the **IIf** function are necessary to return the desired result. The first expression including the **IIf** function is used to check the value of a variable and convert it to zero if it is **Null**.




```
varTemp = IIf(IsNull(varFreight), 0, varFreight) 
varResult = IIf(varTemp > 50, "High", "Low")
```

In the next example, the  **Nz** function provides the same functionality as the first expression, and the desired result is achieved in one step rather than two.




```
varResult = IIf(Nz(varFreight) > 50, "High", "Low")
```

If you supply a value for the optional argument valueifnull, that value will be returned when variant is  **Null**. By including this optional argument, you may be able to avoid the use of an expression containing the **IIf** function. For example, the following expression uses the **IIf** function to return a string if the value of `varFreight` is **Null**.




```
varResult = IIf(IsNull(varFreight), "No Freight Charge", varFreight)
```

In the next example, the optional argument supplied to the  **Nz** function provides the string to be returned if `varFreight` is **Null**.




```
varResult = Nz(varFreight, "No Freight Charge")
```

 **Link provided by:**
![Community Member Icon](images/8b9774c4-6c97-470e-b3a2-56d8f786444c.png) The[UtterAccess](http://www.utteraccess.com) community


- [Nulls and Their Behavior](http://www.utteraccess.com/wiki/index.php/Nulls_And_Their_Behavior)
    

## Example

The following example evaluates a control on a form and returns one of two strings based on the control's value. If the value of the control is  **Null**, the procedure uses the **Nz** function to convert a **Null** value to a zero-length string.


```vb
Public Sub CheckValue() 
 
    Dim frm As Form 
    Dim ctl As Control 
    Dim varResult As Variant 
 
    ' Return Form object variable pointing to Orders form. 
    Set frm = Forms!Orders 
 
    ' Return Control object variable pointing to ShipRegion. 
    Set ctl = frm!ShipRegion 
 
    ' Choose result based on value of control. 
    varResult = IIf(Nz(ctl.Value) = vbNullString, _ 
        "No value.", "Value is " &; ctl.Value &; ".") 
 
    ' Display result. 
    MsgBox varResult, vbExclamation 
 
End Sub
```


## About the Contributors
<a name="AboutContributors"> </a>

UtterAccess is the premier Microsoft Access wiki and help forum. Click here to join. 


## See also
<a name="AboutContributors"> </a>


#### Concepts


[Application Object](application-object-access.md)

