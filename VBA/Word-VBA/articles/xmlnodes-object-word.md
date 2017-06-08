---
title: XMLNodes Object (Word)
ms.prod: word
api_name:
- Word.XMLNodes
ms.assetid: c29850f2-8db2-aef6-57ee-fed1b625616c
ms.date: 06/08/2017
---


# XMLNodes Object (Word)

A collection of  **XMLNode** objects that represents the nodes in the tree view of the **XML Structure** task pane, which indicates the elements that a user has applied to a document. Each node in the tree view is an instance of an **XMLNode** object. The hierarchy in the tree view indicates whether a node contains child nodes.Y


## Remarks

ou can return an  **XMLNodes** collection for a selection, a range, or the entire document. The order in which the **XMLNode** objects appear in the **XMLNodes** collection is the same order in which their start or end tags appear within the specified selection, range, or document.

Use the  **Item** method of the **XMLNodes** collection to return an individual **XMLNode** object. Use the **Validate** method to verify that an XML element is valid according to the applied schemas and that any required child elements exist and are in the required order. Once you run the **Validate** method, use the **ValidationStatus** property to verify whether an element is valid and the **ValidationErrorText** property to display a message to the user as to what the user needs to fix to make the XML in the document conform to the XML schema rules.

The following example validates each of the XML elements in the active document and, if the element or attribute is found to be invalid against the schema, returns a message to the user explaining why the element is invalid.




```vb
Dim objNode As XMLNode 
 
For Each objNode In ActiveDocument.XMLNodes 
 objNode.Validate 
 If objNode.ValidationStatus <> wdXMLValidationStatusOK Then 
 MsgBox objNode.ValidationErrorText(True) 
 End If 
Next
```

Use the  **Add** method to add an XML element to a selection, a range, or the document. The following example inserts the example element from the SimpleSample schema into the active document at the insertion point or surrounding the active selection.


 **Note**  Because XML is case sensitive, the XML element as typed in the Name parameter of the  **Add** method must be typed exactly as it appears in the schema referenced in the Namespace parameter.




```vb
Dim objNode As XMLNode 
Dim intResponse As Integer 
 
Set objNode = Selection.XMLNodes.Add("example", "SimpleSample") 
 
objNode.Validate 
 
If objNode.ValidationStatus < 0 Then 
 intResponse = MsgBox("This element is invalid. " &; _ 
 "Are you sure you want to add it?", vbYesNo) 
 If intResponse = vbNo Then objNode.Delete 
End If
```


## See also


#### Other resources



[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)

