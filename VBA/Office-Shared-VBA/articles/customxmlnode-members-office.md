---
title: CustomXMLNode Members (Office)
ms.prod: office
ms.assetid: fbf957c8-40b8-2f75-fcc8-db0ed6e18438
ms.date: 06/08/2017
---


# CustomXMLNode Members (Office)
Represents an XML node in a tree in a document. The  **CustomXMLNode** object is a member of the **CustomXMLNodes** collection.

Represents an XML node in a tree in a document. The  **CustomXMLNode** object is a member of the **CustomXMLNodes** collection.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[AppendChildNode](customxmlnode-appendchildnode-method-office.md)|Appends a single node as the last child under the context element node in the tree. |
|[AppendChildSubtree](customxmlnode-appendchildsubtree-method-office.md)|Adds a subtree as the last child under the context element node in the tree.|
|[Delete](customxmlnode-delete-method-office.md)|Deletes the current node from the tree (including all of its children, if any exist).|
|[HasChildNodes](customxmlnode-haschildnodes-method-office.md)|Gets  **True** if the current element node has child element nodes.|
|[InsertNodeBefore](customxmlnode-insertnodebefore-method-office.md)|Inserts a new node just before the context node in the tree.|
|[InsertSubtreeBefore](customxmlnode-insertsubtreebefore-method-office.md)|Inserts the specified subtree into the location just before the context node. |
|[RemoveChild](customxmlnode-removechild-method-office.md)|Removes the specified child node from the tree.|
|[ReplaceChildNode](customxmlnode-replacechildnode-method-office.md)|Removes the specified child node (and its subtree) from the main tree, and replaces it with a different node in the same location.|
|[ReplaceChildSubtree](customxmlnode-replacechildsubtree-method-office.md)|Removes the specified node (and its subtree) from the main tree, and replaces it with a different subtree in the same location.|
|[SelectNodes](customxmlnode-selectnodes-method-office.md)|Selects a collection of nodes matching an XPath expression. This method differs from the  **CustomXMLPart**. **SelectNodes** method in that the XPath expression will be evaluated starting with the 'expression' node as the context node.|
|[SelectSingleNode](customxmlnode-selectsinglenode-method-office.md)|Selects a single node from a collection matching an XPath expression. This method differs from the  **CustomXMLPart**. **SelectSingleNode** method in that the XPath expression will be evaluated starting with the 'expression' node as the context node.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](customxmlnode-application-property-office.md)|Gets an  **Application** object that represents the container application for a **CustomXMLNode**. Read-only.|
|[Attributes](customxmlnode-attributes-property-office.md)|Gets a  **CustomXMLNodes** collection representing the attributes of the current element in the current node. Read-only.|
|[BaseName](customxmlnode-basename-property-office.md)|Gets the base name of the node without the namespace prefix, if one exists, in the Document Object Model (DOM). Read-only.|
|[ChildNodes](customxmlnode-childnodes-property-office.md)|Gets a  **CustomXMLNodes** collection containing all of the child elements of the current node. Read-only.|
|[Creator](customxmlnode-creator-property-office.md)|Gets a 32-bit integer that indicates the application in which the  **CustomXMLNode** object was created. Read-only.|
|[FirstChild](customxmlnode-firstchild-property-office.md)|Gets a  **CustomXMLNode** object corresponding to the first child element of the current node. If the node has no child elements (or if it isn't of type **msoCustomXMLNodeElement** ), returns **Nothing**. Read-only.|
|[LastChild](customxmlnode-lastchild-property-office.md)|Gets a  **CustomXMLNode** object corresponding to the last child element of the current node. If the node has no child elements (or if it is not of type **msoCustomXMLNodeElement** ), the property returns **Nothing**. Read-only.|
|[NamespaceURI](customxmlnode-namespaceuri-property-office.md)|Gets the unique address identifier for the namespace of the  **CustomXMLNode** object. Read-only.|
|[NextSibling](customxmlnode-nextsibling-property-office.md)|Gets the next sibling node (element, comment, or processing instruction) of the current node. If the node is the last sibling at its level, the property returns  **Nothing**. Read-only.|
|[NodeType](customxmlnode-nodetype-property-office.md)|Gets the type of the current node. Read-only.|
|[NodeValue](customxmlnode-nodevalue-property-office.md)|Gets or sets the value of the current node. Read/write.|
|[OwnerDocument](customxmlnode-ownerdocument-property-office.md)|Gets the object representing the Microsoft Excel workbook, Microsoft PowerPoint presentation, or the Microsoft Word document associated with this node. Read-only.|
|[OwnerPart](customxmlnode-ownerpart-property-office.md)|Gets the object representing the part associated with this node. Read-only.|
|[Parent](customxmlnode-parent-property-office.md)|Gets the  **Parent** object for the **CustomXMLNode** object. Read-only.|
|[ParentNode](customxmlnode-parentnode-property-office.md)|Gets the parent element node of the current node. If the current node is at the root level, the property returns  **Nothing**. Read-only.|
|[PreviousSibling](customxmlnode-previoussibling-property-office.md)|Gets the previous sibling node (element, comment, or processing instruction) of the current node. If the current node is the first sibling at its level, the property returns  **Nothing**. Read-only.|
|[Text](customxmlnode-text-property-office.md)|Gets or sets the text for the current node. Read/write.|
|[XML](customxmlnode-xml-property-office.md)|Gets the XML representation of the current node and its children, if any exist. Read-only.|
|[XPath](customxmlnode-xpath-property-office.md)|Gets a  **String** with the canonicalized XPath for the current node. If the node is no longer in the Document Object Model (DOM), the property returns an error message. Read-only.|

