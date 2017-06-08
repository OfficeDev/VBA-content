---
title: Remove Data from the Data Store When Deleting a Content Control
ms.prod: word
ms.assetid: 9b7c7345-bd06-b8e2-d401-dea65ad75f92
ms.date: 06/08/2017
---


# Remove Data from the Data Store When Deleting a Content Control

You can delete a content control by calling the  **Delete** method of the **ContentControl** object. For example, the following code deletes the content control with the title "MyTitle".


```vb
ActiveDocument.ContentControls.Item("MyTitle").Delete
```


You can also delete a single node by calling the  **Delete** method of the **CustomDataXMLNode** object that you want to remove. You can delete an entire custom XML part by calling the **Delete** method of the **CustomXMLPart** object that you want to remove.

For more information about content controls, see  [Working with Content Controls](working-with-content-controls.md).
The objects used in these samples are:

-  **[ContentControl](contentcontrol-object-word.md)**
    
-  **[ContentControls](contentcontrols-object-word.md)**
    
-  **CustomXMLPart** (Microsoft Office system core object model)
    
-  **CustomXMLParts** (Microsoft Office system core object model)
    
-  **[XMLMapping](xmlmapping-object-word.md)**
    

## Sample 1

The first code sample creates a content control and sets an XML mapping on a content control.

Build a valid custom XML file, save it to your hard disk drive, and add a data store to the document that contains the information you want to map to.

Suppose the content control is mapped to the following sample custom XML file.




```
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




```vb
Sub AddContentControlAndCustomXMLPart() 
 
    Dim strTitle As String 
    strTitle = "MyTitle" 
    Dim oContentControl As Word.ContentControl 
 
    Set oContentControl = ActiveDocument.ContentControls.Add(wdContentControlText) 
    oContentControl.Title = strTitle 
 
    ActiveDocument.CustomXMLParts.Add 
    ActiveDocument.CustomXMLParts(4).Load ("c:\mySampleCustomXMLFile.xml") 
 
    Dim strXPath As String 
    strXPath = "tree/fruit/fruitType" 
    oContentControl.XMLMapping.SetMapping strXPath 
     
End Sub
```


## Sample 2

The second code sample removes the entire  **CustomXMLPart** object when the content control is deleted.


```vb
Private Sub Document_ContentControlBeforeDelete( _ 
        ByVal OldContentControl As ContentControl, _ 
        ByVal InUndoRedo As Boolean) 
 
    Dim objPart As CustomXMLPart 
     
    'Always void changing the Word document surface during undo! 
    If InUndoRedo Then 
        Return 
    End If 
 
    'Also delete the part with a root element called 'tree' 
    If OldContentControl.Title = "MyTitle" Then 
        For Each objPart In ActiveDocument.CustomXMLParts 
            If objPart.DocumentElement.BaseName = "tree" Then 
                objPart.Delete 
            End If 
        Next part 
    End If 
 
End Sub
```
## Additional Resources

 [Bind a Content Control to a Node in the Data Store](bind-a-content-control-to-a-node-in-the-data-store.md)

