
# Pages Object (Visio)

 **Last modified:** July 28, 2015

 _**Applies to:** Visio 2013 Preview_

Includes a  **Page** object for each drawing page in a document.


## Remarks

To retrieve a  **Pages** collection, use the **Pages** property of a **Document** object.

The default property of a  **Pages** collection is **Item**.

The order of items in a  **Pages** collection is significant: if there are _n_ foreground pages in a document, the first _n_ pages in its **Pages** collection are foreground pages and are in order. The remaining pages in the collection are the background pages of the document; these are in no particular order.

If your Visual Studio solution includes the  **Microsoft.Office.Interop.Visio** reference, this collection maps to the following types:


-  **Microsoft.Office.Interop.Visio.IVPages**
    
