
# Window.ID Property (Visio)

 **Last modified:** July 28, 2015

 _**Applies to:** Visio 2013 Preview_

Gets the ID of an object. Read-only.


## Syntax

 _expression_. **ID**

 _expression_A variable that represents a  **Window** object.


### Return Value

Long


## Remarks

For  **Window** objects, the **ID** property can be used with the **ItemFromID** property of a **Windows** collection to retrieve a **Window** object from the collection without iterating through the collection. A **Window** object whose **Type** property is set to **visAnchorBarBuiltIn** returns an ID of **visWinIDCustProp**,  **visWinIDDrawingExplorer**,  **visWinIDFormulaTracing**,  **visWinIDMasterExplorer**,  **visWinIDPanZoom**,  **visWinIDSizePos**, or  **visWinIDStencilExplorer**. A  **Window** object whose **Type** property is set to **visAnchorBarAddon** returns an ID that is unique within its **Windows** collection for the lifetime of that collection. If a **Window** object has an ID of **visInvalWinID**, you cannot use the  **ItemFromID** property to retrieve the **Window** object from its collection.

