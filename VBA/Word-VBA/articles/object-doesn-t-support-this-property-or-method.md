---
title: Object Doesn't Support this Property or Method
ms.prod: word
ms.assetid: 2595d458-e84f-2107-e27c-24e5a0131f9a
ms.date: 06/08/2017
---


# Object Doesn't Support this Property or Method

The "object doesn't support this property or method" error occurs when you try to use a method or property that the specified object does not support. For example, the following instruction results in an error.


```vb
ActiveDocument.Copy
```


The  **[ActiveDocument](application-activedocument-property-word.md)** property returns a **[Document](document-object-word.md)** object. There is no **Copy** method or property available for the **Document** object, therefore an error occurs. To determine what properties and methods are available for an object, do any of the following.


- Use the Object Browser to determine what members (properties and methods) are available for the selected class (object).
    
- Use the IntelliSense feature in the Visual Basic Editor. When you type a period (.) after a property or method in the Visual Basic Editor, a list of available properties and methods is displayed.
    
- Use Word Visual Basic for Applications Help to determine which properties and methods can be used with an object. Each object topic in Help includes a page that lists the properties and methods for the object. Press F1 while in the Object Browser or while in a module to display the appropriate Help topic.
    
- Use the  **TypeName** function to determine the type of object returned by an expression. The following example displays "Range" because the **[Content](document-content-property-word.md)** property returns a **[Range](range-object-word.md)** object.
    
```vb
  MsgBox TypeName(ActiveDocument.Content)
```


    
    

