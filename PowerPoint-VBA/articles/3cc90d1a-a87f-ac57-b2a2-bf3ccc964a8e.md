
# Axis.BaseUnitIsAuto Property (PowerPoint)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


 **True** if Microsoft Word chooses appropriate base units for the specified category axis. The default is **True**. Read/write  **Boolean**.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **BaseUnitIsAuto**

 _expression_A variable that represents an  ** [Axis](38d5e006-ac32-7bdb-f9f0-e8a858dcbf49.md)** object.


## Remarks
<a name="sectionSection1"> </a>

You cannot set this property for a value axis.


## Example
<a name="sectionSection2"> </a>




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the category axis for the first chart in the active document to use a time scale, with the base unit automatically chosen by Word.




```


With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart

            .Axes(xlCategory).CategoryType = xlTimeScale

            .Axes(xlCategory).BaseUnitIsAuto = True

        End With

    End If

End With
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Axis Object](38d5e006-ac32-7bdb-f9f0-e8a858dcbf49.md)
#### Other resources


 [Axis Object Members](6c4c7cca-d62e-a7c0-b724-30d1be8a44c9.md)
