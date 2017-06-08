---
title: Executing code when setting properties
keywords: vbcn6.chm1101378
f1_keywords:
- vbcn6.chm1101378
ms.prod: office
ms.assetid: ff32c6d2-1857-102f-371e-2d0f6ab848dc
ms.date: 06/08/2017
---


# Executing code when setting properties

You can create  **Property Let**, **Property Set**, and **Property Get** procedures that share the same name. By doing this, you can create a group of related [procedures](vbe-glossary.md) that work together. Once a name is used for a **Property** procedure, that name can't be used to name a **Sub** or **Function** procedure, a [variable](vbe-glossary.md), or a [user-defined type](vbe-glossary.md).

The  **Property Let** statement allows you to create a procedure that sets the value of the [property](vbe-glossary.md). One example might be a  **Property** procedure that creates an inverted property for a bitmap on a form. This is the syntax used to call the **Property Let** procedure:



```
Form1.Inverted = True 

```

The actual work of inverting a bitmap on the form is done within the  **Property Let** procedure:



```vb
Private IsInverted As Boolean 
 
Property Let Inverted(X As Boolean) 
 IsInverted = X 
 If IsInverted Then 
 … 
 (statements) 
 Else 
 (statements) 
 End If 
End Property 

```

The form-level variable stores the setting of your property. By declaring it  **Private**, the user can only change it only using your **Property Let** procedure. Use a name that makes it easy to recognize that the variable is used for the property.
This  **Property Get** procedure is used to return the current state of the property:



```
Property Get Inverted() As Boolean 
 Inverted = IsInverted 
End Property 

```

[Property procedures](vbe-glossary.md) make it easy to execute code at the same time the value of a property is set. You can use property procedures to do the following processing:


- Before a property value is set to determine the value of the property.
    
- After a property value is set, based on the new value.
    


