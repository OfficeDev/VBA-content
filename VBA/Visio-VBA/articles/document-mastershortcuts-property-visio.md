---
title: Document.MasterShortcuts Property (Visio)
keywords: vis_sdr.chm10513885
f1_keywords:
- vis_sdr.chm10513885
ms.prod: visio
api_name:
- Visio.Document.MasterShortcuts
ms.assetid: 7d156dfe-ac70-355a-5927-eb7ebb28bb21
ms.date: 06/08/2017
---


# Document.MasterShortcuts Property (Visio)

Returns the  **MasterShortcuts** collection for a document stencil. Read-only.


## Syntax

 _expression_ . **MasterShortcuts**

 _expression_ A variable that represents a **Document** object.


### Return Value

MasterShortcuts


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **MasterShortcuts** property to get the collection of master shortcuts in a document stencil.



Before running this example, replace the reference to  _StencilWithShortucts.vss_ with a reference to a valid .vss file that contains master shortcuts.


### To create a stencil that contains master shortcuts:




1.  Open an existing stencil. (In the **Shapes** window, click **More Shapes**, click  **Open Stencil**, and then double-click a stencil.)
    
2.  Right-click a master in the stencil and click **Copy**.
    
3.  Create a new stencil. (In the **Shapes** window, click **More Shapes**, and then click  **New Stencil (US Units)** or **New Stencil (Metric)**.)
    
4.  Right-click the new stencil and click **Paste Shortcut**.
    
5.  Save the new stencil. (Right-click its title bar and click **Save**.) 
    


In the following code, replace  _StencilWithShortcuts.vss_ with the name of your new stencil.




```vb
 
Public Sub MasterShortcuts_Example() 
  
    Dim vsoMasterShortcuts As Visio.MasterShortcuts  
    Dim vsoMasterShortcut As Visio.MasterShortcut  
    Dim vsoStencil As Visio.Document  
 
    'Get a stencil that contains some shortcuts.  
    Set vsoStencil = Application.Documents ("StencilWithShortcuts.vss ")  
    Set vsoMasterShortcuts = vsoStencil.MasterShortcuts 
  
    For Each vsoMasterShortcut In vsoMasterShortcuts 
  
        'Print some of the more common properties of a  
        'master shortcut to the Immediate window.  
        With vsoMasterShortcut  
            Debug.Print .AlignName  
            Debug.Print .DropActions  
            Debug.Print .IconSize  
            Debug.Print .ID  
            Debug.Print .Index  
            Debug.Print .Name  
            Debug.Print .NameU  
            Debug.Print .ObjectType  
            Debug.Print .Prompt  
            Debug.Print .ShapeHelp  
            Debug.Print .Stat  
            Debug.Print .TargetDocumentName 
  
            'Original master where shortcut points  
            Debug.Print.TargetMasterName  
 
        End With          
 
    Next  
 
End Sub
```


