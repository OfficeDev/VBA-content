---
title: Friend <keyword>
keywords: vblr6.chm1103683
f1_keywords:
- vblr6.chm1103683
ms.prod: office
ms.assetid: 64745732-35c9-cce0-55bb-57133d322ef0
ms.date: 06/08/2017
---


# Friend <keyword>

Modifies the definition of a [procedure](vbe-glossary.md) in a form module or[class module](vbe-glossary.md) to make the procedure callable from modules that are outside the[class](vbe-glossary.md), but part of the [project](vbe-glossary.md) within which the class is defined. **Friend** procedures cannot be used in standard modules.

 **Syntax**

[ **Private** |**Friend** |**Public** ] [ **Static** ] [ **Sub** |**Function** |**Property** ] _procedurename_

The required  _procedurename_ is the name of the procedure to be made visible throughout the project, but not visible to controllers of the class.
 **Remarks**
 **Public** procedures in a class can be called from anywhere, even by controllers of instances of the class. Declaring a procedure **Private** prevents controllers of the object from calling the procedure, but also prevents the procedure from being called from within the project in which the class itself is defined. **Friend** makes the procedure visible throughout the project, but not to a controller of an instance of the object. **Friend** can appear only in form modules and class modules, and can only modify procedure names, not[variables](vbe-glossary.md) or types. Procedures in a class can access the **Friend** procedures of all other classes in a project. **Friend** procedures don't appear in the[type library](vbe-glossary.md) of their class. A **Friend** procedure can't be late bound.

## Friend Example

When placed in a class module, the following code makes the member variable dblBalance accessible to all users of the class within the project. Any user of the class can get the value; only code within the project can assign a value to that variable.


```vb
Private dblBalance As Double 
 
Public Property Get Balance() As Double 
 Balance = dblBalance 
End Property 
 
Friend Property Let Balance(dblNewBalance As Double) 
 dblBalance = dblNewBalance 
End Property 

```


