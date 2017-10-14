---
title: Enum Statement
keywords: vblr6.chm1103514
f1_keywords:
- vblr6.chm1103514
ms.prod: office
ms.assetid: 22dbc78e-5ce7-f6ea-21dd-67d5db0d64d8
ms.date: 06/08/2017
---


# Enum Statement

Declares a type for an enumeration.

 **Syntax**

[ **Public** |**Private** ] **Enum**_name_

 _membername_ [= _constantexpression_ ]
 _membername_ [= _constantexpression_ ]
 **. . .**
 **End Enum**
The  **Enum** statement has these parts:


|**Part**|**Description**|
|:-----|:-----|
|**Public**|Optional. Specifies that the  **Enum** type is visible throughout the[project](vbe-glossary.md).  **Enum** types are **Public** by default.|
|**Private**|Optional. Specifies that the  **Enum** type is visible only within the[module](vbe-glossary.md) in which it appears.|
| _name_|Required. The name of the  **Enum** type. The _name_ must be a valid Visual Basic identifier and is specified as the type when declaring[variables](vbe-glossary.md) or[parameters](vbe-glossary.md) of the **Enum** type.|
| _membername_|Required. A valid Visual Basic identifier specifying the name by which a constituent element of the  **Enum** type will be known.|
| _constantexpression_|Optional. Value of the element (evaluates to a  **Long** ). If no _constantexpression_ is specified, the value assigned is either zero (if it is the first _membername_ ), or 1 greater than the value of the immediately preceding _membername_.|
 **Remarks**
Enumeration variables are variables declared with an  **Enum** type. Both variables and parameters can be declared with an **Enum** type. The elements of the **Enum** type are initialized to constant values within the **Enum** statement. The assigned values can't be modified at[run time](vbe-glossary.md) and can include both positive and negative numbers. For example:



```vb
Enum SecurityLevel 
 IllegalEntry = -1 
 SecurityLevel1 = 0 
 SecurityLevel2 = 1 
End Enum 

```

An  **Enum** statement can appear only at[module level](vbe-glossary.md). Once the  **Enum** type is defined, it can be used to declare variables, parameters, or[procedures](vbe-glossary.md) returning its type. You can't qualify an **Enum** type name with a module name. **Public** **Enum** types in a[class module](vbe-glossary.md) are not members of the class; however, they are written to the[type library](vbe-glossary.md).  **Enum** types defined in[standard modules](vbe-glossary.md) aren't written to type libraries. **Public Enum** types of the same name can't be defined in both standard modules and class modules, since they share the same name space. When two **Enum** types in different type libraries have the same name, but different elements, a reference to a variable of the type depends on which type library has higher priority in the **References**.
You can't use an  **Enum** type as the target in a **With** block.

## Example

The following example shows the  **Enum** statement used to define a collection of named constants. In this case, the constants are colors you might choose to design data entry forms for a database.


```vb
Public Enum InterfaceColors 
 icMistyRose = &;HE1E4FF&; 
 icSlateGray = &;H908070&; 
 icDodgerBlue = &;HFF901E&; 
 icDeepSkyBlue = &;HFFBF00&; 
 icSpringGreen = &;H7FFF00&; 
 icForestGreen = &;H228B22&; 
 icGoldenrod = &;H20A5DA&; 
 icFirebrick = &;H2222B2&; 
End Enum
```


