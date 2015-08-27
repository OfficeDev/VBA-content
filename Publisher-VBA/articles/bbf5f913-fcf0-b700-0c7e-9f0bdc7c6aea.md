
# Shapes.AddCallout Method (Publisher)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Adds a new  ** [Shape](666cb7f0-62a8-f419-9838-007ef29506ee.md)** object representing a borderless line callout to the specified ** [Shapes](52e069a6-d54b-a11a-1cba-96174329cb02.md)** collection.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **AddCallout**( **_Type_**,  **_Left_**,  **_Top_**,  **_Width_**,  **_Height_**)

 _expression_A variable that represents a  **Shapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Type|Required| **MsoCalloutType**|The type of callout line.|
|Left|Required| **Variant**|The position of the left edge of the shape representing the line callout.|
|Top|Required| **Variant**|The position of the top edge of the shape representing the line callout.|
|Width|Required| **Variant**|The width of the shape representing the line callout.|
|Height|Required| **Variant**|The height of the shape representing the line callout.|

### Return Value

Shape


## Remarks
<a name="sectionSection1"> </a>

For the Left, Top, Width, and Height arguments, numeric values are evaluated in points; strings can be in any units supported by Microsoft Publisher (for example, "2.5 in").

The Type parameter can be one of these  **MsoCalloutType** constants.



| **msoCalloutOne**|A horizontal or vertical single-segment callout line.|
| **msoCalloutTwo**|A freely-rotating single-segment callout line.|
| **msoCalloutThree**|A two-segment callout line.|
| **msoCalloutFour**|A three-segment callout line.|

## Example
<a name="sectionSection2"> </a>

The following example adds a new freely-rotating callout line to the first page of the active publication.


```
Dim shpCallout As Shape 
 
Set shpCallout = ActiveDocument.Pages(1).Shapes.AddCallout _ 
 (Type:=msoCalloutTwo, _ 
 Left:=144, Top:=216, _ 
 Width:=36, Height:=72)
```

