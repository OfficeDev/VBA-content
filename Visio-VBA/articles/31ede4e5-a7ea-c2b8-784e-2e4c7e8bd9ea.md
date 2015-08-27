
# Viewer.GetErrorMessage Method (Visio Viewer)

 **Last modified:** March 09, 2015

 _**Applies to:** Visio 2013_

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Returns a string that describes the specified error message code in Microsoft Visio Viewer.

## Syntax
<a name="sectionSection0"> </a>

 _expression_. **GetErrorMessage**( **_ErrorCode_**)

 _expression_An expression that returns a  **Viewer** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|ErrorCode|Required| **Long**|The error message code for which you want to get a description.|

### Return Value

String


## Remarks
<a name="sectionSection1"> </a>

If you pass an error code that Visio Viewer does not recognize, the  **GetErrorMessage** method will return either a string saying so, or nothing.

If you pass the value that the  ** [LastErrorCode](cbef3230-128c-3976-04da-eec6da9f6225.md)** property returns, the **GetErrorMessage** method returns the last error code that Visio Viewer returned.


## Example
<a name="sectionSection2"> </a>

The following code shows how to use the  **GetErrorMessage** method to get a description of the last error code that Visio Viewer returned.


```
Debug.Print vsoViewer.GetErrorMessage(vsoViewer.LastErrorCode)
```

