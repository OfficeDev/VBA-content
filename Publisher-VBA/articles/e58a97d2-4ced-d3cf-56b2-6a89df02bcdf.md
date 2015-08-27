
# PictureFormat.OriginalHasAlphaChannel Property (Publisher)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Returns an  **MsoTriState** constant depending on whether the original, linked picture contains an alpha channel. Read-only.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **OriginalHasAlphaChannel**

 _expression_A variable that represents an  **PictureFormat** object.


### Return Value

MsoTriState


## Remarks
<a name="sectionSection1"> </a>

This property only applies to linked pictures. Returns "Permission Denied" for shapes representing embedded or pasted pictures.

Use either of the following properties to determine whether a shape represents a linked picture:


-  The ** [Type](bb712dd4-5d81-10e0-9b4c-4af6a09a3c71.md)** property of the ** [Shape](666cb7f0-62a8-f419-9838-007ef29506ee.md)** object
    
- The  ** [IsLinked](2215cee8-864d-7228-8692-a428385d2be2.md)** property of the ** [PictureFormat](aa30ea9d-b91f-acdf-2e60-8a9f506f28b4.md)** object
    


An alpha channel is a special 8-bit channel used by some image processing software to contain additional data, such as masking information or transparency information.

The  **OriginalHasAlphaChannel** property value can be one of the **MsoTriState** constants declared in the Microsoft Office type library and shown in the following table.



|**Constant**|**Description**|
|:-----|:-----|
| **msoFalse**| The original, linked picture does not contain an alpha channel.|
| **msoTriStateMixed**| Indicates a combination of **msoTrue** and **msoFalse** for the specified shape range.|
| **msoTrue**|The original, linked picture contains an alpha channel.|

## Example
<a name="sectionSection2"> </a>

The following example returns whether the first shape on the first page of the active publication contains an alpha channel. If the picture is linked, and the original picture contains an alpha channel, that is also returned. This example assumes the shape is a picture.


```
With ActiveDocument.Pages(1).Shapes(1).PictureFormat 
 If .HasAlphaChannel = msoTrue Then 
 Debug.Print .Filename 
 Debug.Print "This picture contains an alpha channel." 
 
 If .IsLinked = msoTrue Then 
 If .OriginalHasAlphaChannel = msoTrue Then 
 Debug.Print "The linked picture " &amp; _ 
 "also contains an alpha channel." 
 End If 
 End If 
 End If 
End With 

```

