
# Viewer.OnDocumentLoaded Event (Visio Viewer)

 **Last modified:** June 07, 2012

 _**Applies to:** Visio 2013_

 **In this article**
 [Syntax](#sectionSection1)
 [Remarks](#sectionSection2)
 [Example](#sectionSection3)


Occurs after a document is loaded into Microsoft Visio Viewer.

## Syntax
<a name="sectionSection1"> </a>

 _expression_. **OnDocumentLoaded**

 _expression_An expression that returns a  **Viewer** object.


### Return Value

Nothing


## Remarks
<a name="sectionSection2"> </a>

You can load a document into Visio Viewer programmatically by using the  ** [Load](8d33e759-793c-2e3c-3731-131fd51b415a.md)** method.

To capture the  **OnDocumentLoaded** event when you are coding in Visual Basic 6.0, load the document in the **Form_Paint()** procedure. The event may not occur in response to calling the **Load** method within the **Form_Load()** procedure.


## Example
<a name="sectionSection3"> </a>

The following code shows how to display a message in the  **Immediate** window when a document is loaded in Visio Viewer, showing the name of the document.


```
Private Sub vsoViewer_OnDocumentLoaded()



        Debug.Print "Document loaded is "; vsoViewer.SRC

        

End Sub
```

