---
title: UserForm Object
keywords: vblr6.chm1105493
f1_keywords:
- vblr6.chm1105493
ms.prod: office
api_name:
- Office.UserForm
ms.assetid: c64c7a57-56c6-e999-a5fe-dc083bf8fe99
ms.date: 06/08/2017
---


# UserForm Object



A  **UserForm**[object](vbe-glossary.md) is a window or dialog box that makes up part of an application's user interface.
The  **UserForms**[collection](vbe-glossary.md) is a collection whose elements represent each loaded **UserForm** in an application. The **UserForms** collection has a **Count** property, an **Item** property, and an **Add** method. **Count** specifies the number of elements in the collection; **Item** (the default member) specifies a specific collection member; and **Add** places a new **UserForm** element in the collection.
 **Syntax**
 **UserForm**
 **UserForms** [ **.Item** ] **(**_index_**)**
The placeholder  _index_ represents an integer with a range from 0 to **UserForms.Count** - 1. **Item** is the default member of the **UserForms** collection and need not be specified.
 **Remarks**
You can use the  **UserForms** collection to iterate through all loaded user forms in an application. It identifies an intrinsic global[variable](vbe-glossary.md) named **UserForms**. You can pass **UserForms(**_index_**)** to a function whose[argument](vbe-glossary.md) is specified as a **UserForm** class.
User forms have [properties](vbe-glossary.md) that determine appearance such as position, size, and color; and aspects of their behavior.
User forms can also respond to events initiated by a user or triggered by the system. For example, you can write code in the  **Initialize** event procedure of the **UserForm** to initialize[module-level](vbe-glossary.md) variables before the **UserForm** is displayed.
In addition to properties and events, you can use methods to manipulate user forms using code. For example, you can use the  **Move** method to change the location and size of a **UserForm**.
When designing user forms, set the  **BorderStyle** property to define borders, and set the **Caption** property to put text in the title bar. In code, you can use the **Hide** and **Show** methods to make a **UserForm** invisible or visible at[run time](vbe-glossary.md).
 **UserForm** is an[Object data type](vbe-glossary.md). You can declare variables as type  **UserForm** before setting them to an instance of a type of **UserForm** declared at[design time](vbe-glossary.md). Similarly, you can pass an [argument](vbe-glossary.md) to a[procedure](vbe-glossary.md) as type **UserForm**. You can create multiple instances of user forms in code by using the **New** keyword in **Dim**, **Set**, and **Static** statements.
You can access the collection of [controls](vbe-glossary.md) on a **UserForm** using the **Controls** collection. For example, to hide all the controls on a **UserForm**, use code similar to the following:



```vb
For Each Control in UserForm1.Controls
    Control.Visible = False
Next Control

```


