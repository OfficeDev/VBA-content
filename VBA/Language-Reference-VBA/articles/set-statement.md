---
title: Set Statement
keywords: vblr6.chm1009016
f1_keywords:
- vblr6.chm1009016
ms.prod: office
ms.assetid: 59de2927-b338-0038-50b9-3379d7331935
ms.date: 06/08/2017
---


# Set Statement

Assigns an object reference to a [variable](vbe-glossary.md) or[property](vbe-glossary.md).

 **Syntax**

 **Set**_objectvar_**=** {[ **New** ] _objectexpression_ |**Nothing** }

The  **Set** statement syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _objectvar_|Required. Name of the variable or property; follows standard variable naming conventions.|
|**New**|Optional.  **New** is usually used during declaration to enable implicit object creation. When **New** is used with **Set**, it creates a new instance of the[class](vbe-glossary.md). If  _objectvar_ contained a reference to an object, that reference is released when the new one is assigned. The **New**[keyword](vbe-glossary.md) can't be used to create new instances of any intrinsic[data type](vbe-glossary.md) and can't be used to create dependent objects.|
| _objectexpression_|Required. [Expression](vbe-glossary.md) consisting of the name of an object, another declared variable of the same[object type](vbe-glossary.md), or a function or [method](vbe-glossary.md) that returns an object of the same object type.|
|**Nothing**|Optional. Discontinues association of  _objectvar_ with any specific object. Assigning **Nothing** to _objectvar_ releases all the system and memory resources associated with the previously referenced object when no other variable refers to it.|
 **Remarks**
To be valid,  _objectvar_ must be an object type consistent with the object being assigned to it.
The  **Dim**, **Private**, **Public**, **ReDim**, and **Static** statements only declare a variable that refers to an object. No actual object is referred to until you use the **Set** statement to assign a specific object.
The following example illustrates how  **Dim** is used to declare an[array](vbe-glossary.md) with the type `Form1`. No instance of  `Form1` actually exists. **Set** then assigns references to new instances of `Form1` to the . No instance of `Form1` actually exists. **Set** then assigns references to new instances of `Form1` to the `myChildForms` variable. Such code might be used to create child forms in an MDI application.



```vb
Dim myChildForms(1 to 4) As Form1 
Set myChildForms(1) = New Form1 
Set myChildForms(2) = New Form1 
Set myChildForms(3) = New Form1 
Set myChildForms(4) = New Form1 

```

Generally, when you use  **Set** to assign an object reference to a variable, no copy of the object is created for that variable. Instead, a reference to the object is created. More than one[object variable](vbe-glossary.md) can refer to the same object. Because such variables are references to the object rather than copies of the object, any change in the object is reflected in all variables that refer to it. However, when you use the **New** keyword in the **Set** statement, you are actually creating an instance of the object.

## Example

This example uses the  **Set** statement to assign object references to variables. `YourObject` is assumed to be a valid object with a **Text** property.


```vb
Dim YourObject, MyObject, MyStr 
Set MyObject = YourObject    ' Assign object reference. 
' MyObject and YourObject refer to the same object. 
YourObject.Text = "Hello World"    ' Initialize property. 
MyStr = MyObject.Text    ' Returns "Hello World". 
 
' Discontinue association. MyObject no longer refers to YourObject. 
Set MyObject = Nothing    ' Release the object. 

```


