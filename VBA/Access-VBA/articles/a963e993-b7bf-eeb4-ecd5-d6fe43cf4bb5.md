
# Overview of Multidimensional Schemas and Data

 **Last modified:** June 29, 2011

 _ **Applies to:** Access 2013 | Access 2016_

 **In this article**
[Understanding Multidimensional Schemas](#sectionSection0)
[Dimensions](#sectionSection1)
[Hierarchies](#sectionSection2)
[Levels](#sectionSection3)
[Members](#sectionSection4)
[Understanding Multidimensional Schemas](#sectionSection5)
[Dimensions](#sectionSection6)
[Hierarchies](#sectionSection7)
[Levels](#sectionSection8)
[Members](#sectionSection9)



## Understanding Multidimensional Schemas
<a name="sectionSection0"> </a>

The central metadata object in ADO MD is the  _cube_, which consists of a structured set of related dimensions, hierarchies, levels, and members.

A  _dimension_ is an independent category of data from your multidimensional database, derived from your business entities. A dimension typically contains items to be used as query criteria for the measures of the database.

A  _hierarchy_ is a path of aggregation of a dimension. A dimension may have multiple levels of granularity, which have parent-child relationships. A hierarchy defines how these levels are related.

A  _level_ is a step of aggregation in a hierarchy. For dimensions with multiple layers of information, each layer is a level.

A  _member_ is a data item in a dimension. Typically, you create a caption or describe a measure of the database using members.

Cubes are represented by [CubeDef](199235b7-3d98-f655-27bc-94f66e994e06.md) objects in ADO MD. Dimensions, hierarchies, levels, and members are also represented by their corresponding ADO MD objects:[Dimension](12f43cfc-c74e-a2e8-7f6e-75fc68472c4b.md), [Hierarchy](26e4e690-59ad-fb87-66b0-f3310df42d0c.md), [Level](ddbcabce-8777-1068-98a3-be209084f497.md), and [Member](d80c024a-07dc-7a35-f8f2-b4d5b19d89e4.md).


## Dimensions
<a name="sectionSection1"> </a>

The dimensions of a cube depend on your business entities and types of data to be modeled in the database. Typically, each dimension is an independent entry point or mechanism for selecting data.

For example, a cube containing sales data has the following five dimensions: Salesperson, Geography, Time, Products, and Measures. The Measures dimension contains actual sales data values, while the other dimensions represent ways to categorize and group the sales data values.

The Geography dimension has the following set of members:




```
 
{All, North America, Europe, Canada, USA, UK, Germany, Canada-West, 
Canada-East, USA-NW, USA-SW, USA-NE, USA-SE, England, Scotland,  
Wales,Ireland, Germany-North, Germany-South, Ottawa, Toronto,  
Vancouver, Calgary, Seattle, Boise, Los Angeles, Houston,  
Shreveport, Miami, Boston, New York, London, Dover, Glasgow,  
Edinburgh, Cardiff, Pembroke, Belfast, Berlin,  
Hamburg, Munich, Stuttgart} 

```


## Hierarchies
<a name="sectionSection2"> </a>

Hierarchies define the ways in which the levels of a dimension can be "rolled up" or grouped. A dimension can have more than one hierarchy.


## Levels
<a name="sectionSection3"> </a>

In the example Geography dimension pictured in the previous figure, each box represents a level in the hierarchy.

Each level has a set of members, as follows:


- The World  `= {All}`
    
- Continents  `= {North America, Europe}`
    
- Countries  `= {Canada, USA, UK, Germany}`
    
- Regions  `= {Canada-East, Canada-West, USA-NE, USA-NW, USA-SE, USA-SW, England, Ireland, Scotland, Wales, Germany-North, Germany-South}`
    
- Cities  `= {Ottawa, Toronto, Vancouver, Calgary, Seattle, Boise, Los Angeles, Houston, Shreveport, Miami, Boston, New York, London, Dover, Glasgow, Edinburgh, Cardiff, Pembroke, Belfast, Berlin, Hamburg, Munich, Stuttgart}`
    

## Members
<a name="sectionSection4"> </a>

Members at the leaf level of a hierarchy have no children, and members at the root level have no parent. All other members have at least one parent and at least one child. For example, a partial traversal of the hierarchy tree in the Geography dimension yields the following parent-child relationships:


-  `{All} (parent of) {Europe, North America}`
    
-  `{North America} (parent of) {Canada, USA}`
    
-  `{USA} (parent of) {USA-NE, USA-NW, USA-SE, USA-SW}`
    
-  `{USA-NW} (parent of) {Boise, Seattle}`
    
Members can be consolidated along one or more hierarchies per dimension. 

This example also illustrates another characteristic: Some members of the Week level of the Year-Week hierarchy do not appear in any level of the Year-Quarter hierarchy. Thus, a hierarchy need not include all members of a dimension.


## Understanding Multidimensional Schemas
<a name="sectionSection5"> </a>

The central metadata object in ADO MD is the  _cube_, which consists of a structured set of related dimensions, hierarchies, levels, and members.

A  _dimension_ is an independent category of data from your multidimensional database, derived from your business entities. A dimension typically contains items to be used as query criteria for the measures of the database.

A  _hierarchy_ is a path of aggregation of a dimension. A dimension may have multiple levels of granularity, which have parent-child relationships. A hierarchy defines how these levels are related.

A  _level_ is a step of aggregation in a hierarchy. For dimensions with multiple layers of information, each layer is a level.

A  _member_ is a data item in a dimension. Typically, you create a caption or describe a measure of the database using members.

Cubes are represented by [CubeDef](199235b7-3d98-f655-27bc-94f66e994e06.md) objects in ADO MD. Dimensions, hierarchies, levels, and members are also represented by their corresponding ADO MD objects:[Dimension](12f43cfc-c74e-a2e8-7f6e-75fc68472c4b.md), [Hierarchy](26e4e690-59ad-fb87-66b0-f3310df42d0c.md), [Level](ddbcabce-8777-1068-98a3-be209084f497.md), and [Member](d80c024a-07dc-7a35-f8f2-b4d5b19d89e4.md).


## Dimensions
<a name="sectionSection6"> </a>

The dimensions of a cube depend on your business entities and types of data to be modeled in the database. Typically, each dimension is an independent entry point or mechanism for selecting data.

For example, a cube containing sales data has the following five dimensions: Salesperson, Geography, Time, Products, and Measures. The Measures dimension contains actual sales data values, while the other dimensions represent ways to categorize and group the sales data values.

The Geography dimension has the following set of members:




```
 
{All, North America, Europe, Canada, USA, UK, Germany, Canada-West, 
Canada-East, USA-NW, USA-SW, USA-NE, USA-SE, England, Scotland,  
Wales,Ireland, Germany-North, Germany-South, Ottawa, Toronto,  
Vancouver, Calgary, Seattle, Boise, Los Angeles, Houston,  
Shreveport, Miami, Boston, New York, London, Dover, Glasgow,  
Edinburgh, Cardiff, Pembroke, Belfast, Berlin,  
Hamburg, Munich, Stuttgart} 

```


## Hierarchies
<a name="sectionSection7"> </a>

Hierarchies define the ways in which the levels of a dimension can be "rolled up" or grouped. A dimension can have more than one hierarchy.


## Levels
<a name="sectionSection8"> </a>

In the example Geography dimension pictured in the previous figure, each box represents a level in the hierarchy.

Each level has a set of members, as follows:


- The World  `= {All}`
    
- Continents  `= {North America, Europe}`
    
- Countries  `= {Canada, USA, UK, Germany}`
    
- Regions  `= {Canada-East, Canada-West, USA-NE, USA-NW, USA-SE, USA-SW, England, Ireland, Scotland, Wales, Germany-North, Germany-South}`
    
- Cities  `= {Ottawa, Toronto, Vancouver, Calgary, Seattle, Boise, Los Angeles, Houston, Shreveport, Miami, Boston, New York, London, Dover, Glasgow, Edinburgh, Cardiff, Pembroke, Belfast, Berlin, Hamburg, Munich, Stuttgart}`
    

## Members
<a name="sectionSection9"> </a>

Members at the leaf level of a hierarchy have no children, and members at the root level have no parent. All other members have at least one parent and at least one child. For example, a partial traversal of the hierarchy tree in the Geography dimension yields the following parent-child relationships:


-  `{All} (parent of) {Europe, North America}`
    
-  `{North America} (parent of) {Canada, USA}`
    
-  `{USA} (parent of) {USA-NE, USA-NW, USA-SE, USA-SW}`
    
-  `{USA-NW} (parent of) {Boise, Seattle}`
    
Members can be consolidated along one or more hierarchies per dimension. 

This example also illustrates another characteristic: Some members of the Week level of the Year-Week hierarchy do not appear in any level of the Year-Quarter hierarchy. Thus, a hierarchy need not include all members of a dimension.

