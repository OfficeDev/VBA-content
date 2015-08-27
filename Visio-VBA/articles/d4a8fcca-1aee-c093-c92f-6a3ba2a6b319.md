
# ApplicationSettings.SnapStrengthGuidesX Property (Visio)

 **Last modified:** July 28, 2015

 _**Applies to:** Visio 2013 Preview_

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Specifies the distance in pixels along the x-axis that guides pull when snapping is enabled. Read/write.

## Syntax
<a name="sectionSection0"> </a>

 _expression_. **SnapStrengthGuidesX**

 _expression_A variable that represents an  **ApplicationSettings** object.


### Return Value

Long


## Remarks
<a name="sectionSection1"> </a>

Setting the  **SnapStrengthGuidesX** property is equivalent to setting the **Guides** option under **Snap strength** on the **Advanced** tab in the **Snap &amp; Glue** dialog box (click the **Visual Aids** arrow on the **View** tab). Setting snap strength in the UI sets both _x_ and _y_ values to the same value.

The minimum allowable value for the  **SnapStrengthGuidesX** property is 0 (zero), and the maximum is 999. Attempting to set a value outside that range returns an error. The default value is 10.


## Example
<a name="sectionSection2"> </a>

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **SnapStrengthGuidesX** property to print the current snap strength guides _x_-axis setting in the Immediate window. It also shows how to get an  **ApplicationSettings** object from the Visio **Application** object.


```
Public Sub SnapStrengthGuidesX_Example() 
 
 Dim vsoApplicationSettings As Visio.ApplicationSettings 
 Dim lngSnapStrength As Long 
 
 Set vsoApplicationSettings = Visio.Application.Settings 
 lngSnapStrength = vsoApplicationSettings.SnapStrengthGuidesX 
 
 Debug.Print lngSnapStrength 
 
End Sub 
 

```

