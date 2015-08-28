
# Axis.MinorUnitScale Property (PowerPoint)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Returns or sets the minor unit scale value for the category axis when the  ** [CategoryType](bbcb485d-9464-33c8-ca9b-e3463bc9e884.md)** property is set to **xlTimeScale**. Read/write  ** [XlTimeUnit](7da25d66-7339-9cb2-13da-81dda86a55b4.md)**.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **MinorUnitScale**

 _expression_A variable that represents an  ** [Axis](38d5e006-ac32-7bdb-f9f0-e8a858dcbf49.md)** object.


## Remarks
<a name="sectionSection1"> </a>

 **MinorUnitScale** can be one of the following **XlTimeUnit** constants:


-  **xlMonths**
    
-  **xlDays**
    
-  **xlYears**
    

## Example
<a name="sectionSection2"> </a>




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the category axis to use a time scale and sets the major and minor units.




```
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.Axes(xlCategory)

            .CategoryType = xlTimeScale

            .MajorUnit = 5

            .MajorUnitScale = xlDays

            .MinorUnit = 1

            .MinorUnitScale = xlDays

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
