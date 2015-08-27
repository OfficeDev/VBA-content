
# Page.BackPage Property (Visio)

 **Last modified:** July 28, 2015

 _**Applies to:** Visio 2013 Preview_

Gets or sets the background page of a page. Read/write.


## Syntax

 _expression_. **BackPage**

 _expression_A variable that represents a  **Page** object.


### Return Value

Variant


## Remarks

If a page has no background, its  **BackPage** property returns an empty **Variant**. Otherwise the returned  **Variant** refers to a **Page** object, the background page of the indicated page.

To assign a background page to a page, set the page's  **BackPage** property to the name of that background page. To cause a page to have no background page, pass an empty string to the **BackPage** property.

Markup overlay pages cannot have background pages, so you cannot use the  **BackPage** property to assign a background page to a markup overlay page.


 **Note**  In earlier versions of Visio (through version 4.1), the  **BackPage** property returned an object (as opposed to a **Variant** of type **Object**) and it accepted a string (as opposed to a  **Variant** of type **String**). Because of changes in Automation support tools, the property has been modified so that it accepts and returns a  **Variant**.

