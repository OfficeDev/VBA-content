
# Viewer.LayerDeleted Property (Visio Viewer)

 **Last modified:** March 09, 2015

 _**Applies to:** Visio 2013_

 **In this article**
 [Syntax](#sectionSection1)
 [Remarks](#sectionSection2)
 [Example](#sectionSection3)


Gets a value that indicates whether the layer at the specified index in the drawing open in Microsoft Visio Viewer is deleted. Read-only.

## Syntax
<a name="sectionSection1"> </a>

 _expression_. **LayerDeleted**( **_LayerIndex_**)

 _expression_An expression that returns a  **Viewer** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|LayerIndex|Required| **Long**|The index of the layer in the collection of layers in the drawing open in Visio Viewer.|

### Return Value

 **Boolean**


## Remarks
<a name="sectionSection2"> </a>

The collection of layers is one-based, so the index of the first layer in the collection is 1. If there are no layers in the drawing, the  **LayerDeleted** property returns **False**.


## Example
<a name="sectionSection3"> </a>

The following code gets a value that indicates whether the layer at index position 1 is deleted.


```
Debug.Print vsoViewer.LayerDeleted(1)
```

