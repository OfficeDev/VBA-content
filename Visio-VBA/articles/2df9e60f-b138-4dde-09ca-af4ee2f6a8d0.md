
# Shape.DistanceFrom Property (Visio)

 **Last modified:** July 28, 2015

 _**Applies to:** Visio 2013 Preview_

Returns the distance from one shape to another, measured between the closest points on the two shapes. Both shapes must be on the same page or in the same master. Read-only.


## Syntax

 _expression_. **DistanceFrom**( **_OtherShape_**,  **_Flags_**)

 _expression_A variable that represents a  **Shape** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|OtherShape|Required| **[IVSHAPE]**|The other  **Shape** object involved in the comparison.|
|Flags|Required| **Integer**|Flags that influence the type of entries returned in results.|

### Return Value

Double


## Remarks

The  **DistanceFrom** property returns:




- Zero and raises an exception if the shapes being compared are in different masters or on different pages.
    
- Zero if the shapes being compared are overlapping.
    
- Zero if one shape contains the other shape or one shape is contained within the other shape.
    


The Flags argument can be any combination of the values of the constants defined in the following table. These constants are also defined in **VisSpatialRelationFlags** in the Microsoft Visio type library.



|**Constant **|**Value **|**Description **|
|:-----|:-----|:-----|
| **visSpatialIncludeDataGraphics**|&amp;H40|Includes data graphic callout shapes and their sub-shapes. By default, data graphic callout shapes and their subshapes are not included. If the parent shape is itself a data graphic callout, searches are made between the parent shape's geometry and non-callout shapes, unless this flag is set.|
| **visSpatialIncludeHidden**|&amp;H10 |Consider hidden Geometry sections. By default, hidden Geometry sections do not influence the result. |
| **visSpatialIgnoreVisible**|&amp;H20 |Do not consider visible Geometry sections. By default, visible Geometry sections influence the result. |
Use the NoShow cell to determine whether a Geometry section is hidden or visible. Hidden Geometry sections have a value of TRUE and visible Geometry sections have a value of FALSE in the NoShow cell.

If the parent shape or OtherShape has no geometry, or ifFlags excludes consideration of all geometry of either shape, the **DistanceFrom** property returns a large number (1E+30) that should be construed as infinite.

The  **DistanceFrom** property does not consider the width of a shape's line, shadows, line ends, control points, or connection points when comparing two shapes.

