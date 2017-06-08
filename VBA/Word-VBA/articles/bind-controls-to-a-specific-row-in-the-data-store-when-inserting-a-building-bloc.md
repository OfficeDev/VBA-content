---
title: Bind Controls to a Specific Row in the Data Store When Inserting a Building Block
ms.prod: word
ms.assetid: c701f613-c14e-267e-7a9b-ea1f193397c2
ms.date: 06/08/2017
---


# Bind Controls to a Specific Row in the Data Store When Inserting a Building Block

A document building block is a predesigned piece of content, such as a cover page, a header, a footer, or a custom-built clause in a contract. Custom building blocks make it easier for you to quickly create professional-looking Word documents.

You can use content controls within custom building blocks that have been mapped to XML that contains data. The contents of those content controls are then automatically linked to the appropriate custom XML part (if present) when the part is inserted. Alternatively, as the following sample shows, you can dynamically add the custom XML part and the XML mappings when the part is inserted. For example, to construct a cover page, you can place a picture content control that displays an image retrieved from an element in an attached  **CustomXMLPart** object. Similarly, you can create the project name using a text content control that you map to an element in a **CustomXMLPart** object containing the project name.

This makes it easier to update your data. To update one of these items, you can write a few lines of code to update every document stored on the server that uses this cover page building block. For example, you can replace an old logo with a new one. Or, if the project title changes, you can update the text in the XML element that you mapped to the text content control containing the project name, thereby automatically updating all the documents stored on the server.

The objects used in these samples are:

-  **[ContentControl](contentcontrol-object-word.md)**
    
-  **[ContentControls](contentcontrols-object-word.md)**
    
-  **CustomXMLPart** (Microsoft Office core object model)
    
-  **CustomXMLParts** (Microsoft Office core object model)
    
-  **[XMLMapping](xmlmapping-object-word.md)**
    

## Sample

Suppose the user has inserted your custom document building block into a document, and based on that action, you want to insert and map to the custom XML part.


```
<?xml version="1.0" encoding="utf-8" ?> 
<projects> 
  <project> 
    <title>Data-Driven Document Generation</title> 
    <manager>Frank Martinez</manager> 
    <customer>Northwind Traders</customer> 
  </project> 
</projects>
```

The following sample code loads the previous XML file and maps each content control to the appropriate XML node in that new custom XML part, when the document building block called "Company Report" is added.




```vb
Private Sub Document_BuildingBlockInsert(ByVal Range As Range, _ 
        ByVal Name As String, ByVal Category As String, _ 
        ByVal Type As String, ByVal Template As String) 
 
    Dim cc As ContentControl 
    Dim part As CustomXMLPart 
 
    If Name = "Company Report" Then 
        'add the custom XML 
        ActiveDocument.CustomXMLParts.Add 
        Set part = ActiveDocument.CustomXMLParts(ActiveDocument.CustomXMLParts.Count).Load("c:\myProjects.xml") 
 
        'map the controls 
        For Each cc In Range.ContentControls 
            cc.XMLMapping.SetMapping cc.XMLMapping.XPath, cc.XMLMapping.PrefixMappings, part 
        Next cc 
    End If 
 
End Sub
```


