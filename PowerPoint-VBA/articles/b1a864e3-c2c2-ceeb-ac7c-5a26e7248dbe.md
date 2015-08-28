
# ConnectorFormat.EndConnect Method (PowerPoint)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Attaches the end of the specified connector to a specified shape. 


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **EndConnect**( **_ConnectedShape_**,  **_ConnectionSite_**)

 _expression_A variable that represents a  **ConnectorFormat** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|ConnectedShape|Required| ** [Shape](1da93849-99e0-827e-ced3-c6cf7f8569f3.md)**|The shape to attach the end of the connector to. The specified  **Shape** object must be in the same ** [Shapes](eb208855-254e-1a0f-884b-4a5edcfd584d.md)** collection as the connector.|
|ConnectionSite|Required| **Long**|A connection site on the shape specified by ConnectedShape. Must be an integer between 1 and the integer returned by the  **ConnectionSiteCount** property of the specified shape. If you want the connector to automatically find the shortest path between the two shapes it connects, specify any valid integer for this argument and then use the **RerouteConnections**method after the connector is attached to shapes at both ends.|

## Remarks
<a name="sectionSection1"> </a>

If there's already a connection between the end of the connector and another shape, that connection is broken. If the end of the connector isn't already positioned at the specified connecting site, this method moves the end of the connector to the connecting site and adjusts the size and position of the connector. Use the  **BeginConnect** method to attach the beginning of the connector to a shape.

When you attach a connector to an object, the size and position of the connector are automatically adjusted, if necessary.


## Example
<a name="sectionSection2"> </a>

This example adds two rectangles to the first slide in the active presentation and connects them with a curved connector. Notice that the  **RerouteConnections** method makes it irrelevant what values you supply for the ConnectionSite arguments used with the **BeginConnect** and **EndConnect** methods.


```
Set myDocument = ActivePresentation.Slides(1)

Set s = myDocument.Shapes

Set firstRect = s.AddShape(msoShapeRectangle, 100, 50, 200, 100)

Set secondRect = s.AddShape(msoShapeRectangle, 300, 300, 200, 100)

With s.AddConnector(msoConnectorCurve, 0, 0, 100, 100) _

        .ConnectorFormat

    .BeginConnect ConnectedShape:=firstRect, ConnectionSite:=1

    .EndConnect ConnectedShape:=secondRect, ConnectionSite:=1

    .Parent.RerouteConnections

End With
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [ConnectorFormat Object](54504fab-8279-1012-db7f-3f19a4840637.md)
#### Other resources


 [ConnectorFormat Object Members](446eda0c-4992-d38f-b054-355de3058011.md)
