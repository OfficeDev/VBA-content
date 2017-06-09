---
title: Invalid use of New keyword
keywords: vblr6.chm1040348
f1_keywords:
- vblr6.chm1040348
ms.prod: office
ms.assetid: 6bdc77a1-dde7-974e-4fee-b9279b4f3ae3
ms.date: 06/08/2017
---


# Invalid use of New keyword

The  **New** keyword can only be applied to a creatable object (an instance of a[class](vbe-glossary.md) or[Automation object](vbe-glossary.md)). This error has the following causes and solutions:



- You tried to instantiate something that can have only one instance. For example, you tried to create a new instance of a [module](vbe-glossary.md) by specifying `Module1` in a statement like the following:
    
```vb
Dim MyMod As New Module1 

  ```


    You can't create the new instance, since a module can have only one instance.
    
- You tried to instantiate an Automation object, but it was not a creatable object. For example, you tried to create a new instance of a list box by specifying  **ListBox** in a statement like the following:
    
  ```
  ' Valid syntax to create the variable. 
Dim MyListBox As ListBox     
Dim MyFormInst As Form 
' Invalid syntax to instantiate the object. 
Set MyFormInst = New Form 
Set MyListBox = New ListBox 

  ```


     **ListBox** and **Form** are class names, not specific object names. You can use them to specify that a[variable](vbe-glossary.md) will be a reference to a certain[object type](vbe-glossary.md), as with the valid  **Dim** statements above. But you can't use them to instantiate the objects themselves in a **Set** statement. You must specify a specific object, rather than the generic class name, in the **Set** statement:
    


  ```
  ' Valid syntax to create new instance of a form or list box. 
Set MyFormInst = New Form1 
Set MyListBox = New List1 

  ```


For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

