---
title: Understanding Objects, Properties, and Methods
ms.prod: word
ms.assetid: b79853f7-a497-78eb-4ab0-95b6d7a79912
ms.date: 06/08/2017
---


# Understanding Objects, Properties, and Methods

Objects are the fundamental building blocks of Microsoft Visual Basic; almost everything that you do in Visual Basic involves modifying objects. Every element of Microsoft Word—such as documents, tables, paragraphs, bookmarks, and fields—can be represented by an object in Visual Basic.


## What are objects and collections?

An object represents an element of Word, such as a document, a paragraph, a bookmark, or a single character. A collection is an object that contains several other objects, usually of the same type; for example, all the bookmark objects in a document are contained in a single collection object. By using properties and methods, you can modify a single object or a whole collection of objects.


## What is a property?

A property is an attribute of an object or an aspect of its behavior. For example, properties of a document include its name, its content, and its save status, and whether change tracking is turned on. To change the characteristics of an object, you change the values of its properties.

To set the value of a property, follow the reference to an object with a period, the property name, an equal sign, and the new property value. The following example turns on change tracking in the document named "MyDoc.doc".




```vb
Sub TrackChanges() 
    Documents("Sales.doc").TrackRevisions = True 
End Sub
```

In this example,  `Documents` refers to the collection of open documents, and the name "Sales.doc" identifies a single document in the collection. The **[TrackRevisions](document-trackrevisions-property-word.md)** property is set for that single document.

Some properties cannot be set. The Help topic for a property indicates whether that property can be set (read/write) or can only be read (read-only).

You can return information about an object by returning the value of one of its properties. The following example returns the name of the active document.




```vb
Sub GetDocumentName() 
    Dim strDocName As String 
    strDocName = ActiveDocument.Name 
    MsgBox strDocName 
End Sub
```

In this example,  `ActiveDocument` refers to the document in the active window in Word. The name of that document is assigned to the variable refers to the document in the active window in Word. The name of that document is assigned to the variable `strDocName`.

 **Remarks**

The Help topic for each property indicates whether you can set that property (read/write), only read the property (read-only), or only write the property (write-only). Also, the Object Browser in the Visual Basic Editor displays the read/write status at the bottom of the browser window when the property is selected.


## What is a method?

A method is an action that an object can perform. For example, just as a document can be printed, the  **[Document](document-object-word.md)** object has a **[PrintOut](document-printout-method-word.md)** method. Methods often have arguments that qualify how the action is performed. The following example prints the first three pages of the active document.


```vb
Sub PrintThreePages() 
    ActiveDocument.PrintOut Range:=wdPrintRangeOfPages, Pages:="1-3" 
End Sub
```

In most cases, methods are actions and properties are qualities. Using a method causes something to happen to an object, while using a property returns information about the object or causes a quality about the object to change.


## Returning an object

Most objects are returned by returning a single object from the collection. For example, the  **[Documents](document-object-word.md)** collection contains the open Word documents. You use the **[Documents](application-documents-property-word.md)** property of the **[Application](application-object-word.md)** object (the object at the top of the Word object hierarchy) to return the **Documents** collection.

After you access the collection, you can return a single object by using an index value in parentheses (this is similar to how you work with arrays). The index value is usually a number or a name. For more information, see  [Returning an Object from a Collection](returning-an-object-from-a-collection-word.md).

The following example uses the  **Documents** property to access the **Documents** collection. The index number is used to return the first document in the **Documents** collection. The **[Close](document-close-method-word.md)** method is then applied to the **Document** object to close the first document in the **Documents** collection.




```vb
Sub CloseDocument() 
    Documents(1).Close 
End Sub
```

The following example uses a name (specified as a string) to identify a  **Document** object within the **Documents** collection.




```vb
Sub CloseSalesDoc() 
    Documents("Sales.doc").Close 
End Sub
```

Collection objects often have methods and properties that you can use to modify the whole collection of objects. The  **Documents** object has a **[Save](documents-save-method-word.md)** method that saves all the documents in the collection. The following example saves the open documents by applying the **Save** method.




```vb
Sub SaveAllOpenDocuments() 
    Documents.Save 
End Sub
```

The  **Document** object also has a **Save** method that is available for saving a single document. The following example saves the document named Sales.doc.




```vb
Sub SaveSalesDoc() 
    Documents("Sales.doc").Save 
End Sub
```

To return an object that is further down in the Word object hierarchy, you must "drill down" to it by using properties and methods to return objects.

To see how this is done, open the Visual Basic Editor and click  **Object Browser** on the **View** menu. Click **Application** in the **Classes** list on the left. Then click **ActiveDocument** from the list of members on the right. The text at the bottom of the Object Browser indicates that **ActiveDocument** is a read-only property that returns a **Document** object. Click **Document** at the bottom of the Object Browser; the **Document** object is automatically selected in the **Classes** list, and the **Members** list displays the members of the **Document** object. Scroll through the list of members until you find **Close**. Click the  **Close** method. The text at the bottom of the Object Browser window shows the syntax for the method. For more information about the method, press F1 or click the **Help** button to jump to the **Close** method Help topic.

Given this information, you can write the following instruction to close the active document.




```vb
Sub CloseDocSaveChanges() 
    ActiveDocument.Close SaveChanges:=wdSaveChanges 
End Sub
```

The following example maximizes the active document window.




```vb
Sub MaximizeDocumentWindow() 
    ActiveDocument.ActiveWindow.WindowState = wdWindowStateMaximize 
End Sub
```

The  **ActiveWindow** property returns a **Window** object that represents the active window. The **WindowState** property is set to the maximize constant ( **wdWindowStateMaximize**).

The following example creates a document and displays the  **Save As** dialog box so that a name can be provided for the document.




```vb
Sub CreateSaveNewDocument() 
    Documents.Add.Save 
End Sub
```

The  **Documents** property returns the **Documents** collection. The **[Add](documents-add-method-word.md)** method creates a new document and returns a **Document** object. The **Save** method is then applied to the **Document** object.

As you can see, you use methods or properties to drill down to an object. That is, you return an object by applying a method or property to an object above it in the object hierarchy. After you return the object that you want, you can apply the methods and control the properties of that object.


## Getting Help on objects, methods, and properties

Until you become familiar with the Word object model, there are tools that you can use to help you drill down through the hierarchy.


-  **Microsoft IntelliSense**. When you type a period (.) after an object in the Visual Basic Editor, a list of available properties and methods is displayed. For example, if you type  `Application.`, a drop-down list of methods and properties of the  **Application** object is displayed.
    
-  **Help**. You can also use Help to find out which properties and methods can be used with an object. Each object topic in Help includes a See Also jump that displays a list of properties and methods for the object. Press  **F1** while in the Object Browser or in a module to jump to the appropriate Help topic.
    
-  **Object Browser**. The Object Browser in the Visual Basic Editor displays the members (properties and methods) of the Word objects.
    

