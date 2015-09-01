
# Options.DisableFeaturesbyDefault Property (Word)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


 **True** for Microsoft Word to disable in all documents all features introduced after the version of Word specified in the ** [DisableFeaturesIntroducedAfterbyDefault](a7cf788b-f5c1-2d7e-b3de-1261b2a65c45.md)**. The default value is  **False**. Read/write  **Boolean**.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **DisableFeaturesbyDefault**

 _expression_A variable that represents a  ** [Options](873b7b99-3fe1-fd89-9ece-a9355cb827dc.md)** object.


## Remarks
<a name="sectionSection1"> </a>

The  **DisableFeaturesByDefault** property sets a global option for the application. If you want to disable features introduced after Word 97 for Windows for the document only, use the ** [DisableFeatures](40a62de3-f74e-d604-d3fc-dfb26abeb313.md)**property.


## Example
<a name="sectionSection2"> </a>

This example disables all features introduced after Word for Windows 95, versions 7.0 and 7.0a, for all documents.


```
Sub FeaturesDisableByDefault() 
 With Application.Options 
 
 'Checks whether features are disabled 
 If .DisableFeaturesbyDefault = True Then 
 
 'If they are, disables all features after Word for Windows 95 
 .DisableFeaturesIntroducedAfterbyDefault = wd70 
 Else 
 
 'If not, turns on the disable features option and disables 
 'all features introduced after Word for Windows 95 
 .DisableFeaturesbyDefault = True 
 .DisableFeaturesIntroducedAfterbyDefault = wd70 
 End If 
 End With 
End Sub
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Options Object](873b7b99-3fe1-fd89-9ece-a9355cb827dc.md)
#### Other resources


 [Options Object Members](76cd9dfe-6bbb-4c3d-0bfc-79a62bedd15e.md)
