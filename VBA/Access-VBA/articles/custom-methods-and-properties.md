---
title: Custom Methods and Properties
keywords: vbaac10.chm5187125
f1_keywords:
- vbaac10.chm5187125
ms.prod: access
ms.assetid: 2915eacb-240f-6876-0afb-1db038c4ecba
ms.date: 06/08/2017
---


# Custom Methods and Properties

You can use a class module to create a definition for a new custom object. When you create a new instance of a class, you create a new object and return a reference to it.

Any public procedures defined within the class module become methods of the new object. The  **Sub** statement defines a method that doesn't return a value; the **Function** statement defines a method that may return a value of a specified type.

Any  **Property Let**, **Property Get** or **Property Set** procedures you define become properties of the new object. **Property Get** procedures retrieve the value of a property. **Property Let** procedures set the value of a nonobject property. **Property Set** procedures set the value of an object property.

For example, you can use a class module to create an interface layer between your application and a set of Windows application programming interface (API) functions that it calls. To do this, you create a set of simple procedures that call more complicated procedures in a DLL. When you create an instance of this class, the procedures you've created become methods of the new object. You can apply these methods as you would the methods of any object, and in doing so you also call the API functions.

