---
title: Conceptual Differences Between WordBasic and Visual Basic
keywords: vbawd10.chm5210304
f1_keywords:
- vbawd10.chm5210304
ms.prod: word
ms.assetid: 2ec0fa57-68c4-f4e9-000c-91a2b97ac9ac
ms.date: 06/08/2017
---


# Conceptual Differences Between WordBasic and Visual Basic

The primary difference between Visual Basic for Applications (VBA) and WordBasic is that, whereas the WordBasic language consists of a flat list of approximately 900 commands, Visual Basic consists of a hierarchy of objects, each of which exposes a specific set of methods and properties (similar to statements and functions in WordBasic). While most WordBasic commands can be run at any time, Visual Basic only exposes the methods and properties of the available objects at a given time.

Objects are the fundamental building blocks of Visual Basic; almost everything you do in Visual Basic involves modifying objects. Every element of Word—such as documents, paragraphs, fields, and bookmarks—can be represented by an object in Visual Basic. Unlike commands in a flat list, there are objects that can only be accessed from other objects. For example, the  **[Font](font-object-word.md)** object can be accessed from various objects, including the **[Style](style-object-word.md)**,  **[Selection](selection-object-word.md)**, and  **[Find](find-object-word.md)** objects.

The programming task of applying bold formatting demonstrates the differences between the two programming languages. The following WordBasic instruction applies bold formatting to the selection.




```
Bold 1
```

The following example is the Visual Basic equivalent for applying bold formatting to the selection.



```vb
Selection.Font.Bold = True
```

Visual Basic does not include a  **Bold** statement and function. Instead, there is a **Bold** property. (A property is usually an attribute of an object, such as its size, its color, or whether or not it is bold.) **Bold** is a property of the **Font** object. Likewise, **[Font](selection-font-property-word.md)** is a property of the **Selection** object that returns a **Font** object. Following the object hierarchy, you can build the instruction to apply bold formatting to the selection.
The  **Bold** property is a read/write Boolean property. This means that the **Bold** property can be set to **True** or **False** (on or off), or the current value can be returned. The following WordBasic instruction returns a value indicating whether bold formatting is applied to the selection.



```
x = Bold()
```

The following example is the Visual Basic equivalent for returning the bold formatting status from the selection.



```
x = Selection.Font.Bold
```


## The Visual Basic thought process

To perform a task in Visual Basic, you need to determine the appropriate object. For example, if you want to apply character formatting found in the  **Font** dialog box, use the **Font** object. Then you need to determine how to "drill down" through the Word object hierarchy from the **[Application](application-object-word.md)** object to the **Font** object, through the objects that contain the **Font** object you want to modify. After you have determined the path to your object (for example, ), use the Object Browser, Help, or the features such as Auto List Members in the Visual Basic Editor to determine what properties and methods can be applied to the object. For more information about drilling down to objects using properties and methods, see [Understanding Objects, Properties, and Methods](understanding-objects-properties-and-methods.md).

Properties and methods are often available to multiple objects in the Word object hierarchy. For example, the following instruction applies bold formatting to the entire document.




```vb
ActiveDocument.Content.Bold = True
```

Also, objects themselves often exist in more than one place in the object hierarchy.


## The Selection and Range objects

Most WordBasic commands modify the selection. For example, the  **Bold** command formats the selection with bold formatting. The **InsertField** command inserts a field at the insertion point. When you want to work with the selection in Visual Basic, you use the **[Selection](selection-childshaperange-property-word.md)** property to return the **Selection** object. The selection can be a block of text or just the insertion point.

The following Visual Basic example inserts text and a new paragraph after the selection.




```
Selection.InsertAfter Text:="Hello World" 
Selection.InsertParagraphAfter
```

In addition to working with the selection, you can define and work with various ranges of text in a document. A  **[Range](range-object-word.md)** object refers to a contiguous area in a document with a starting character position and an ending character position. Similar to the way bookmarks are used in a document, **Range** objects are used in Visual Basic to identify portions of a document. However, unlike a bookmark, a **Range** object is invisible to the user unless the **Range** has been selected using the **[Select](selection-boldrun-method-word.md)** method. For example, you can use Visual Basic to apply bold formatting anywhere in the document without changing the selection. The following example applies bold formatting to the first 10 characters in the active document.




```vb
ActiveDocument.Range(Start:=0, End:=10).Bold = True
```

The following example applies bold formatting to the first paragraph.




```vb
ActiveDocument.Paragraphs(1).Range.Bold = True
```

Both of these example change the formatting in the active document without changing the selection. For more information about the  **Range** object, see [Working with Range objects](working-with-range-objects.md).


