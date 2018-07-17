---
title: Const Directive
keywords: vblr6.chm1014460
f1_keywords:
- vblr6.chm1014460
ms.prod: office
ms.assetid: c5d74b3a-75b1-1263-ab98-82a1a1087207
ms.date: 06/08/2017
---


# #Const Directive

Used to define [conditional compiler constants](vbe-glossary.md) for Visual Basic.

## Syntax

 **#Const** _constname_ = _expression_

The **#Const** compiler directive syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _constname_|Required;  **Variant** ( **String** ). Name of the [constant](vbe-glossary.md); follows standard [variable](vbe-glossary.md) naming conventions.|
| _expression_|Required. Literal, other conditional compiler constant, or any combination that includes any or all arithmetic or logical operators except  **Is**.|
 **Remarks**
Conditional compiler constants are always [Private](vbe-glossary.md) to the [module](vbe-glossary.md) in which they appear. It is not possible to create [Public](vbe-glossary.md) compiler constants using the **#Const** directive. **Public** compiler constants can only be created in the user interface.
Only conditional compiler constants and literals can be used in  _expression_. Using a standard constant defined with **Const**, or using a constant that is undefined, causes an error to occur. Conversely, constants defined using the **#Const** [keyword](vbe-glossary.md) can only be used for conditional compilation.
Conditional compiler constants are always evaluated at the [module level](vbe-glossary.md), regardless of their placement in code.

## Example

This example uses the  **#Const** directive to declare conditional compiler constants for use in **#If...#Else...#End If** constructs.


```
#Const DebugVersion = 1 ' Will evaluate true in #If block. 

```


