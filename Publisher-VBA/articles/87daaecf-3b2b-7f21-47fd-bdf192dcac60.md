
# LineFormat.BeginArrowheadLength Property (Publisher)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Returns or sets an  **MsoArrowheadLength**constant indicating the length of the arrowhead at the beginning of the specified line. Read/write.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **BeginArrowheadLength**

 _expression_A variable that represents a  **LineFormat** object.


### Return Value

MsoArrowheadLength


## Remarks
<a name="sectionSection1"> </a>

The  **BeginArrowheadLength** property value can be one of the ** [MsoArrowheadLength](http://msdn.microsoft.com/library/e39957f3-ffdd-17fe-dc60-1c3f8c5b14ce%28Office.15%29.aspx)** constants declared in the Microsoft Office type library.

Use the  ** [EndArrowheadLength](3e46e63b-54b2-edbf-0dc1-fba2c3a5d945.md)** property to return or set the length of the arrowhead at the end of the line.


## Example
<a name="sectionSection2"> </a>

This example adds a line to the active publication. There is a short, narrow oval on the line's starting point and a long, wide triangle on its endpoint.


```
With ActiveDocument.Pages(1).Shapes _ 
 .AddLine(BeginX:=100, BeginY:=100, _ 
 EndX:=200, EndY:=300).Line 
 .BeginArrowheadLength = msoArrowheadShort 
 .BeginArrowheadStyle = msoArrowheadOval 
 .BeginArrowheadWidth = msoArrowheadNarrow 
 .EndArrowheadLength = msoArrowheadLong 
 .EndArrowheadStyle = msoArrowheadTriangle 
 .EndArrowheadWidth = msoArrowheadWide 
End With 

```

