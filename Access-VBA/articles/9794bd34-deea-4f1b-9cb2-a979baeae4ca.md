
# ShapeNode Properties (Excel)

 **Last modified:** July 28, 2015


## Properties



|**Name**|**Description**|
|:-----|:-----|
| [Application](98e77d56-875c-7696-2b2d-5f36409fa129.md)|When used without an object qualifier, this property returns an  ** [Application](19b73597-5cf9-4f56-8227-b5211f657f6f.md)**object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an  **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
| [Creator](10c4e270-6b82-85be-2428-3d7509249335.md)|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long**.|
| [EditingType](78a17ed7-7e30-d5f3-4af8-636d65079218.md)|If the specified node is a vertex, this property returns a value that indicates how changes made to the node affect the two segments connected to the node. Read-only  ** [MsoEditingType](5fe5c4f6-6467-c6a7-197c-ff700c384b92.md)**.|
| [Parent](ebb2ff4b-3939-e850-a3ad-1f93f9ded7c3.md)|Returns the parent object for the specified object. Read-only.|
| [Points](fe09c78f-44c9-4e66-df7b-c23720216ec5.md)|Returns the position of the specified node as a coordinate pair. Each coordinate is expressed in points. Read-only  **Variant**.|
| [SegmentType](716e8171-1fd6-941e-209f-e48f5468940f.md)|Returns a value that indicates whether the segment associated with the specified node is straight or curved. If the specified node is a control point for a curved segment, this property returns  **msoSegmentCurve**. Read-only  **MsoSegmentType**.|
