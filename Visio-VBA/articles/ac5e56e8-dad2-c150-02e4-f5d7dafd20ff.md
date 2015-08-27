
# Curve.Start Property (Visio)

 **Last modified:** July 28, 2015

 _**Applies to:** Visio 2013 Preview_

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Returns the start of a  **Curve** object's parameter domain. Read-only.

## Syntax
<a name="sectionSection0"> </a>

 _expression_. **Start**

 _expression_A variable that represents a  **Curve** object.


### Return Value

Double


## Remarks
<a name="sectionSection1"> </a>

The  **Start** property of a **Curve** object returns the value of the starting point in the curve's parameter domain. A **Curve** object describes itself in terms of its parameter domain, which is the range [Start(),End()], where Start() produces the curve's starting point. Note that the **Start** value is not a coordinate pair. Rather, it represents the relative position along the curve of the starting point. For a line, for example, the value of **Start** typically is 0, the value of **End** is 1, and you can use the **Point** method of the **Curve** object to determine the coordinates of any point along the curve by determining the relative location of the point between the start and endpoints.


## Example
<a name="sectionSection2"> </a>

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **Start** property to display the value of the starting point of a curve. It uses the **Point** method to find the midpoint of the curve.


```
 
Sub Start_Example() 
 
 Dim vsoShape As Visio.Shape 
 Dim vsoPaths As Visio.Paths 
 Dim vsoPath As Visio.Path 
 Dim vsoCurve As Visio.Curve 
 Dim dblStartpoint As Double 
 Dim dblEndpoint As Double 
 Dim dblX As Double 
 Dim dblY As Double 
 Dim intOuterLoopCounter As Integer 
 Dim intInnerLoopCounter As Integer 
 
 'Draw a shape and get its Paths collection. 
 Set vsoPaths = ActivePage.DrawOval(1, 1, 4, 4).Paths 
 
 'Iterate through the Path objects in the Paths collection. 
 For intOuterLoopCounter = 1 To vsoPaths.Count 
 
 Set vsoPath = vsoPaths.Item(intOuterLoopCounter) 
 Debug.Print "Path object " &amp; intOuterLoopCounter 
 
 'Iterate through the curves in a Path object. 
 For intInnerLoopCounter = 1 To vsoPath.Count 
 
 Set vsoCurve = vsoPath(intInnerLoopCounter) 
 Debug.Print "Curve number " &amp; intInnerLoopCounter 
 
 'Display the start point of the curve. 
 dblStartpoint = vsoCurve.Start 
 Debug.Print "Startpoint = " &amp; dblStartpoint 
 
 'Display the endpoint of the curve. 
 dblEndpoint = vsoCurve.End 
 Debug.Print "Endpoint = " &amp; dblEndpoint 
 
 'Find the midpoint of the curve. 
 vsoCurve.Point ((dblEndpoint - dblStartpoint) / 2), dblX, dblY 
 Debug.Print "Midpoint: x = " &amp; dblx; ", y = " &amp; dblY 
 
 Next intInnerLoopCounter 
 Debug.Print "This path has " &amp; intInnerLoopCounter - 1 &amp; " curve object(s)." 
 
 Next intOuterLoopCounter 
 Debug.Print "This shape has " &amp; intOuterLoopCounter - 1 &amp; " path object(s)." 
 
End Sub 

```

