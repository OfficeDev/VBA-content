
# Page.Index Property (Visio)

 **Last modified:** July 28, 2015

 _**Applies to:** Visio 2013 Preview_

Gets or sets the ordinal position of a page in a  **Pages** collection. Read/write.


## Syntax

 _expression_. **Index**

 _expression_A variable that represents a  **Page** object.


### Return Value

Integer


## Remarks

In versions earlier than 2002, the  **Index** property of the **Page** object was read-only.

The  **Pages** collection is indexed starting with 1 rather than zero (0), so the index of the first element is 1, the index of the second element is 2, and so on. The index of the last element in a collection is the same as the value of that collection's **Count** property. You can iterate through a collection by using these index values. Adding objects to or deleting objects from a collection can change the index values of other objects in the collection.

You may only assign a new index to a foreground page. Background pages are unordered. Use the  **Background** property to determine if a given page is a background page.

Use the  **BackPage** property to assign a background page to a foreground page or to another background page.

