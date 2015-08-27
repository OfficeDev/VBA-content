
# Viewer.ReviewerCount Property (Visio Viewer)

 **Last modified:** June 07, 2012

 _**Applies to:** Visio 2013_

 **In this article**
 [Syntax](#sectionSection1)
 [Remarks](#sectionSection2)
 [Example](#sectionSection3)


Gets the count of reviewers in the current document open in Microsoft Visio Viewer. Read-only.

## Syntax
<a name="sectionSection1"> </a>

 _expression_. **ReviewerCount**

 _expression_An expression that returns a  **Viewer** object.


### Return Value

 **Long**


## Remarks
<a name="sectionSection2"> </a>

The collection of reviewers is one-based, so the index of the first reviewer in the collection is 1.


## Example
<a name="sectionSection3"> </a>

The following code gets the number of reviewers in the drawing open in Visio Viewer and displays it in the  **Immediate** window.


```
Debug.Print vsoViewer.ReviewerCount
```

