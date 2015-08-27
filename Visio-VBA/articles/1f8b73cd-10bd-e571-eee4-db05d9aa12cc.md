
# Application.Quit Method (Visio)

 **Last modified:** July 28, 2015

 _**Applies to:** Visio 2013 Preview_

Closes the indicated instance of Microsoft Visio.


## Syntax

 _expression_. **Quit**

 _expression_A variable that represents an  **Application** object.


### Return Value

Nothing


## Remarks

If the  **Quit** method is invoked when any open document has unsaved changes, a dialog box appears asking if you want to save the document. To quit the application without saving and seeing the dialog box, set the **Saved** property of the **Document** object representing the document to **True** immediately before quitting. Set the **Saved** property to **True** only if you are sure you want to close the document without saving changes, because you will lose any unsaved changes.

