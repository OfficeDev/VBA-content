---
title: Source Property (Visual Basic for Applications)
keywords: vblr6.chm1014188
f1_keywords:
- vblr6.chm1014188
ms.prod: office
ms.assetid: bbf51a29-682a-8fc5-52db-89647c184885
ms.date: 06/08/2017
---


# Source Property (Visual Basic for Applications)



Returns or sets a [string expression](vbe-glossary.md) specifying the name of the object or application that originally generated the error. Read/write.
 **Remarks**
The  **Source**[property](vbe-glossary.md) specifies a string expression representing the object that generated the error; the[expression](vbe-glossary.md) is usually the object's[class](vbe-glossary.md) name or programmatic ID. Use **Source** to provide information when your code is unable to handle an error generated in an accessed object. For example, if you access Microsoft Excel and it generates a `Division by zero` error, Microsoft Excel sets **Err.Number** to its error code for that error and sets **Source** to `Excel.Application`.
When generating an error from code,  **Source** is your application's programmatic ID. For[class modules](vbe-glossary.md),  **Source** should contain a name having the form _project.class_. When an unexpected error occurs in your code, the **Source** property is automatically filled in. For errors in a[standard module](vbe-glossary.md),  **Source** contains the[project](vbe-glossary.md) name. For errors in a class module, **Source** contains a name with the _project.class_ form.

## Example

This example assigns the Programmatic ID of an Automation object created in Visual Basic to the variable  `MyObjectID`, and then assigns that to the  **Source** property of the **Err** object when it generates an error with the **Raise** method. When handling errors, you should not use the **Source** property (or any **Err** properties other than **Number** ) programmatically. The only valid use of properties other than **Number** is for displaying rich information to an end user in cases where you can't handle an error. The example assumes that `App` and `MyClass` are valid references.


```vb
Dim MyClass, MyObjectID, MyHelpFile, MyHelpContext
' An object of type MyClass generates an error and fills all Err object
' properties, including Source, which receives MyObjectID, which is a 
' combination of the Title property of the App object and the Name
' property of the MyClass object.
MyObjectID = App.Title &; "." &; MyClass.Name
Err. Raise    Number := vbObjectError + 894, Source := MyObjectID, _
                Description := "Was not able to complete your task", _
                HelpFile := MyHelpFile, HelpContext := MyHelpContext 

```


