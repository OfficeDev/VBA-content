---
title: Understanding Objects, Properties, Methods, and Events
keywords: vbcn6.chm1076676
f1_keywords:
- vbcn6.chm1076676
ms.prod: office
ms.assetid: 98f2fdcf-f4af-9b18-6164-7252c0a7c668
ms.date: 06/08/2017
---


# Understanding Objects, Properties, Methods, and Events

An object represents an element of an application, such as a worksheet, a cell, a chart, a form, or a report. In Visual Basic code, you must identify an object before you can apply one of the object's [methods](vbe-glossary.md) or change the value of one of its [properties](vbe-glossary.md).

A collection is an object that contains several other objects, usually, but not always, of the same type. In Microsoft Excel, for example, the  **Workbooks** object contains all the open **Workbook** objects. In Visual Basic, the **Forms** collection contains all the **Form** objects in an application.

Items in a collection can be identified by number or by name. For example, in the following [procedure](vbe-glossary.md), identifies the first open  **Workbook** object.




```vb
Sub CloseFirst() 
 Workbooks(1).Close 
End Sub
```

The following procedure uses a name specified as a string to identify a  **Form** object.



```vb
Sub CloseForm() 
 Forms("MyForm.frm").Close 
End Sub
```

You can also manipulate an entire collection of objects if the objects share common [methods](vbe-glossary.md). For example, the following procedure closes all open forms.



```vb
Sub CloseAll() 
 Forms.Close 
End Sub
```

A method is an action that an object can perform. For example,  **Add** is a method of the **ComboBox** object, because it adds a new entry to a combo box.
The following procedure uses the  **Add** method to add a new item to a **ComboBox**.



```vb
Sub AddEntry(newEntry as String) 
 Combo1.Add newEntry 
End Sub
```

A property is an attribute of an object that defines one of the object's characteristics, such as size, color, or screen location, or an aspect of its behavior, such as whether it is enabled or visible. To change the characteristics of an object, you change the values of its properties.
To set the value of a property, follow the reference to an object with a period, the property name, an equal sign ( **=** ), and the new property value. For example, the following procedure changes the caption of a Visual Basic form by setting the **Caption** property.



```vb
Sub ChangeName(newTitle) 
 myForm.Caption = newTitle 
End Sub
```

You can't set some properties. The Help topic for each property indicates whether you can set that property (read-write), only read the property (read-only), or only write the property (write-only).
You can retrieve information about an object by returning the value of one of its properties. The following procedure uses a message box to display the title that appears at the top of the currently active form.



```vb
Sub GetFormName() 
 formName = Screen.ActiveForm.Caption 
 MsgBox formName 
End Sub
```

An event is an action recognized by an object, such as clicking the mouse or pressing a key, and for which you can write code to respond. Events can occur as a result of a user action or program code, or they can be triggered by the system.

## Returning Objects

Every application has a way to return the objects it contains. However, they are not all the same, so you must refer to the Help topic for the object or collection you're using in the application to see how to return the object.


