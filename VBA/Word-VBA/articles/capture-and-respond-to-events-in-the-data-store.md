---
title: Capture and Respond to Events in the Data Store
ms.prod: word
ms.assetid: 5d9fd121-be59-0bcf-68d4-48bf62fc5003
ms.date: 06/08/2017
---


# Capture and Respond to Events in the Data Store

The client can listen for and respond to changes on a node or on a node and on all of its children. An add-in can respond to the following events.

On the  **CustomXMLParts** collection:

-  **StreamAfterAdd.** Allows a client to respond after a new store is added to the document.
    
-  **StreamBeforeDelete.** Allows a client to respond before a store is removed from the document.
    
-  **StreamAfterLoad.** Allows a client to respond after a store item is loaded with XML.
    
On the  **CustomXMLPart** object:

-  **NodeAfterInsert.** Allows a client to respond after a new node is added to a store. If the added node contains a subtree, the event fires only once , for the top-most node.
    
-  **NodeAfterDelete.** Allows a client to respond after a node is deleted. If the deleted node contains a subtree, the event fires only once, for the top-most node.
    
-  **NodeAfterReplace.** Allows a client to respond after an XML node is replaced in the store.
    

## Sample



Assume that there is an XML file, C:\test.xml, and two text content controls. The XML file looks like this:




```
<?xml version="1.0" standalone="no"?>  
<root xmlns="urn:test">  
  <a>NodeA</a>  
  <b>NodeB</b>  
</root>
```

One of the powerful things that you can accomplish with XML mapping is to have one mapped text content control update immediately when a user updates another text content control. This is accomplished using events. To do this, create a method with events and run it.




```vb
Dim WithEvents objStream As CustomXMLPart 
 
Sub Demo() 
  Set objStream = ThisDocument.CustomXMLParts(4) 
End Sub
```

Running the Demo subroutine sets up the  _objStream_ variable to listen to events.

Remember, from the previous scenario, that the document has two text content controls, one data mapped to the <a> node and the other data mapped to the "b" node. Suppose you want to set up events so that when the text in the <a> node is modified, the "b" node automatically does something. The following **objStream_NodeAfterReplace** event subroutine accomplishes this.




```vb
Private Sub objStream_NodeAfterReplace( _ 
        ByVal OldNode As Office.CustomXMLNode, _ 
        ByVal NewNode As Office.CustomXMLNode, _ 
        ByVal InUndoRedo As Boolean) 
 
    ' Check if NewNode, which is the node after the change, is 
    ' the "a" node by checking the BaseName of its ParentNode 
  If NewNode.ParentNode.BaseName = "a" Then 
    objStream.DocumentElement.LastChild.Text = "You changed a!" 
  End If 
 
End Sub
```

This routine is triggered after the user changes the text in the first text content control, mapped to element <a>. If the text in the <a> node changes, then the text of the last child in the custom XML part is updated. Because the stream has only two nodes, the last node is the "b" node. After the text of node is updated, the updated text of "You changed a!" automatically appears in the second text content control.

This example is very simple, but it shows what you can do with events, XML mapping, and content controls. You can use code such as this to update any text in a document when one text content control is changed. This is powerful because it assumes nothing about the document formatting, and it does not work with the document formatting. Instead, it works against the schema that you attach to the document.

 [Bind a Content Control to a Node in the Data Store](bind-a-content-control-to-a-node-in-the-data-store.md)
 [Bind a Content Control to a Node in the Data Store](bind-a-content-control-to-a-node-in-the-data-store.md)

