
# Connects Object (Visio)

 **Last modified:** July 28, 2015

 _**Applies to:** Visio 2013_

 Includes a **Connect** object for each connection between two shapes in a drawing, such as a line and a box in an organization chart.


## Remarks

The default property of a  **Connects** collection is **Item**.

Use the  **Connects** property of a **Shape** object to retrieve a **Connects** collection with a **Connect** object for every **Shape** object to which the indicated **Shape** object is connected (glued).

Use the  **FromConnects** property of a **Shape** object to retrieve a **Connects** collection with a **Connect** object for every **Shape** object that is connected (glued) to the indicated **Shape** object.

Use the  **Connects** property of a **Page** object to retrieve a **Connects** collection with an entry for every connection on the **Page** object.

Use the  **Connects** property of a **Master** object to retrieve a **Connects** collection with an entry for every connection in the **Master** object.

If your Visual Studio solution includes the  **Microsoft.Office.Interop.Visio** reference, this collection maps to the following types:


-  **Microsoft.Office.Interop.Visio.IVConnects.GetEnumerator()** (to enumerate the **Connect** objects.)
    
