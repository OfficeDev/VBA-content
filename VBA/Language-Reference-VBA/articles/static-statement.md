---
title: Static Statement
keywords: vblr6.chm1009031
f1_keywords:
- vblr6.chm1009031
ms.prod: office
ms.assetid: 56b817bc-7324-cc0b-10ec-7ffea364b91e
ms.date: 06/08/2017
---


# Static Statement

Used at [procedure level](vbe-glossary.md) to declare [variables](vbe-glossary.md) and allocate storage space. Variables declared with the **Static** statement retain their values as long as the code is running.

 **Syntax**

 **Static** _varname_ [ **(** [ _subscripts_ ] **)** ] [ **As** [ **New** ] _type_ ] [ **,** _varname_ [ **(** [ _subscripts_ ] **)** ] [ **As** [ **New** ] _type_ ]] **. . .**

The  **Static** statement syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _varname_|Required. Name of the variable; follows standard variable naming conventions.|
| _subscripts_|Optional. Dimensions of an [array](vbe-glossary.md) variable; up to 60 multiple dimensions may be declared. The _subscripts_ [argument](vbe-glossary.md) uses the following syntax: [ _lower_**To** ] _upper_ [ **,** [ _lower_**To** ] _upper_ ] **. . .** When not explicitly stated in _lower_, the lower bound of an array is controlled by the **Option** **Base** statement. The lower bound is zero if no **Option** **Base** statement is present.|
|**New**|Optional. [Keyword](vbe-glossary.md) that enables implicit creation of an object. If you use **New** when declaring the [object variable](vbe-glossary.md), a new instance of the object is created on first reference to it, so you don't have to use the  **Set** statement to assign the object reference. The **New** keyword can't be used to declare variables of any intrinsic [data type](vbe-glossary.md) and can't be used to declare instances of dependent objects.|
| _type_|Optional. Data type of the variable; may be [Byte](vbe-glossary.md), [Boolean](vbe-glossary.md), [Integer](vbe-glossary.md), [Long](vbe-glossary.md), [Currency](vbe-glossary.md), [Single](vbe-glossary.md), [Double](vbe-glossary.md), [Decimal](vbe-glossary.md) (not currently supported), [Date](vbe-glossary.md), [String](vbe-glossary.md), (for variable-length strings),  **String** * _length_ (for fixed-length strings), [Object](vbe-glossary.md), [Variant](vbe-glossary.md), a [user-defined type](vbe-glossary.md), or an [object type](vbe-glossary.md). Use a separate  **As** _type_ clause for each variable being defined.|
 **Remarks**
Once [module](vbe-glossary.md) code is running, variables declared with the **Static** [statement](vbe-glossary.md) retain their value until the module is reset or restarted. In [class modules](vbe-glossary.md), variables declared with the  **Static** statement retain their value in each class instance until that instance is destroyed. In [form modules](vbe-glossary.md), static variables retain their value until the form is closed. Use the  **Static** statement in nonstatic [procedure](vbe-glossary.md)s to explicitly declare variables that are visible only within the procedure, but whose lifetime is the same as the module in which the procedure is defined.
Use a  **Static** statement within a procedure to declare the data type of a variable that retains its value between procedure calls. For example, the following statement declares a fixed-size array of integers:



```
Static EmployeeNumber(200) As Integer 

```

The following statement declares a variable for a new instance of a worksheet:



```
Static X As New Worksheet 

```

If the  **New** keyword isn't used when declaring an object variable, the variable that refers to the object must be assigned an existing object using the **Set** statement before it can be used. Until it is assigned an object, the declared object variable has the special value **Nothing**, which indicates that it doesn't refer to any particular instance of an object. When you use the **New** keyword in the [declaration](vbe-glossary.md), an instance of the object is created on the first reference to the object.
If you don't specify a data type or object type, and there is no  **Def**_type_ statement in the module, the variable is **Variant** by default.

 **Note**  The  **Static** statement and the **Static** keyword are similar, but used for different effects. If you declare a procedure using the **Static** keyword (as in `Static Sub CountSales ()`), the storage space for all local variables within the procedure is allocated once, and the value of the variables is preserved for the entire time the program is running. For nonstatic procedures, storage space for variables is allocated each time the procedure is called and released when the procedure is exited. The  **Static** statement is used to declare specific variables within nonstatic procedures to preserve their value for as long as the program is running.

When variables are initialized, a numeric variable is initialized to 0, a variable-length string is initialized to a zero-length string (""), and a fixed-length string is filled with zeros.  **Variant** variables are initialized to [Empty](vbe-glossary.md). Each element of a user-defined type variable is initialized as if it were a separate variable.

 **Note**   When you use **Static** statements within a procedure, put them at the beginning of the procedure with other declarative statements such as **Dim**.


## Example

This example uses the  **Static** statement to retain the value of a variable for as long as module code is running.


```vb
' Function definition. 
Function KeepTotal(Number) 
    ' Only the variable Accumulate preserves its value between calls. 
    Static Accumulate 
    Accumulate = Accumulate + Number 
    KeepTotal = Accumulate 
End Function 
 
' Static function definition. 
Static Function MyFunction(Arg1, Arg2, Arg3) 
    ' All local variables preserve value between function calls. 
    Accumulate = Arg1 + Arg2 + Arg3 
    Half = Accumulate / 2 
    MyFunction = Half 
End Function
```


