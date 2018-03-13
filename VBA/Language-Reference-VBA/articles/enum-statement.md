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


| <strong>Part</strong>       | <strong>Description</strong>                                                                                                                                                                                                                                                        |
|:----------------------------|:------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| <strong>Public</strong>     | Optional. Specifies that the  <strong>Enum</strong> type is visible throughout the[project](vbe-glossary.md).  <strong>Enum</strong> types are <strong>Public</strong> by default.                                                                                                  |
| <strong>Private</strong>    | Optional. Specifies that the  <strong>Enum</strong> type is visible only within the[module](vbe-glossary.md) in which it appears.                                                                                                                                                   |
| <em>name</em>               | Required. The name of the  <strong>Enum</strong> type. The <em>name</em> must be a valid Visual Basic identifier and is specified as the type when declaring[variables](vbe-glossary.md) or[parameters](vbe-glossary.md) of the <strong>Enum</strong> type.                         |
| <em>membername</em>         | Required. A valid Visual Basic identifier specifying the name by which a constituent element of the  <strong>Enum</strong> type will be known.                                                                                                                                      |
| <em>constantexpression</em> | Optional. Value of the element (evaluates to a  <strong>Long</strong> ). If no <em>constantexpression</em> is specified, the value assigned is either zero (if it is the first <em>membername</em> ), or 1 greater than the value of the immediately preceding <em>membername</em>. |

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


