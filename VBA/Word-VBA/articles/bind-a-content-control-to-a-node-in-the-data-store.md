---
title: Bind a Content Control to a Node in the Data Store
ms.prod: word
ms.assetid: f76bcb03-1361-2235-b3ef-cdd078210698
ms.date: 06/08/2017
---


# Bind a Content Control to a Node in the Data Store

XML mapping is a feature of Word that enables you to create a link between a document and an XML file. This creates true data/view separation between the document formatting and the custom XML data.

XML mapping enables you to map an element in a custom XML part that is attached to the document. The data store provides access to all of the custom XML parts that are stored in an open file. You can refer to any node within any custom XML part inside the data store.

For more information about content controls, see  [Working with Content Controls](working-with-content-controls.md).

The objects used in this sample are:

-  **[ContentControl](contentcontrol-object-word.md)**
    
-  **[ContentControls](contentcontrols-object-word.md)**
    
-  **CustomXMLPart** (Microsoft Office core object model)
    
-  **CustomXMLParts** (Microsoft Office core object model)
    
-  **[XMLMapping](xmlmapping-object-word.md)**
    

## Sample

The following steps enable you to bind a content control to a node in the document's data store.


1.  **Create the content control to bind to a node in the data store.**Content controls are predefined pieces of content. There are several types of content controls, including text blocks, drop-down menus, combo boxes, calendar controls, and pictures. You can map these content controls to an element in an XML file. By using XML Path Language (XPath), you can programmatically map content in an XML file to a content control. This enables you to write a simple and short application to manipulate and modify data in a document. For more information about content controls, see  [Working with Content Controls](working-with-content-controls.md). The following code sample creates a plain-text content control and gives it a title of "MyTitle".
    
```vb
  Dim strTitle As String 
strTitle = "MyTitle" 
Dim oContentControl As Word.ContentControl 
Set oContentControl = ActiveDocument.ContentControls.Add(wdContentControlText) 
oContentControl.Title = strTitle
```

2.  **Set the XML mapping on the content control.**The data store in a document in the Word object model is contained in the  **[CustomXMLParts](document-customxmlparts-property-word.md)** property of the **[Document](document-object-word.md)** object. The **CustomXMLParts** property returns a **CustomXMLParts** collection that contains **CustomXMLPart** objects. It points to all the custom XML parts that are stored in a document. A **CustomXMLPart** object represents a single custom XML part in the data store. To load custom XML data, you must first add a new custom XML part to a **Document** object by using the Add method of the **CustomXMLParts** collection. This appends a new, empty custom XML part to the document. Because it is empty, there is no XML to map to. Next, you must load XML into the newly defined part by calling the **Load** method of the **CustomXMLPart** object, using a valid path to an XML file as the parameter, or by calling the **LoadXML** method of the **CustomXMLPart** and passing the XML directly. The default custom XML parts stored with a Word document contain the document's standard document properties; you cannot delete these parts. You can always view the contents of a custom XML part by calling the read-only XML property on it. If you call the XML property of a **CustomXMLPart** object, a string is returned, which contains the XML in that data store. Build a valid custom XML file and save it to your hard disk drive. Add a custom XML part to the document that contains the content control you want to map to custom XML data. Suppose the content control will be mapped to the following sample custom XML file.
    
```XML
  <?xml version="1.0" encoding="utf-8" ?>
<tree>
  <fruit>
    <fruitType>peach</fruitType>
    <fruitType>pear</fruitType>
    <fruitType>banana</fruitType>
  </fruit>
</tree>

```


    Now, suppose the content control is mapped to a <fruitType> node of the previous custom XML part. 
    
    The following sample code demonstrates how to attach an XML file to a document, so that it becomes an available data store item. 
    


```vb
  ActiveDocument.CustomXMLParts.Add
ActiveDocument.CustomXMLParts(ActiveDocument.CustomXMLParts.Count).Load ("c:\mySampleCustomXMLFile.xml")

```


    To create an XML mapping, you use an XPath expression to specify the node in the custom XML data part to which you want to map a content control. Setting an XML mapping on a content control then specifies the node in the added custom XML part, using that XPath expression. After you add a custom XML part to your document (and after the custom XML part contains XML), you are ready to map one of its nodes to a content control. To do this, pass a  **String** containing a valid XPath to a **ContentControl** object by using the **[SetMapping](xmlmapping-setmapping-method-word.md)** method of the **[XMLMapping](xmlmapping-object-word.md)** object (using the **[XMLMapping](contentcontrol-xmlmapping-property-word.md)** property of the **ContentControl** object). The following is an example of doing this with an XPath that refers to a data store node containing the value of the first fruitType element.
    


```vb
  Dim strXPath As String 
strXPath = "tree/fruit/fruitType[1]" 
ActiveDocument.ContentControls(1).XMLMapping.SetMapping strXPath 

```

If you omit the optional  **PrefixMappings** and **CustomXMLPart** arguments, Word searches each of the custom XML parts in order and maps the control to the first part that successfully retrieves a custom XML node using the specified XPath.


