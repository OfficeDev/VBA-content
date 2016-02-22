
# Page.Export Method (Visio)

Exports an object from Microsoft Visio to a file format such as .bmp, .dib, .dwg, .dxf, .emf, .emz, .gif, .htm, .jpg, .png, .svg, .svgz, .tif, or .wmf.


## Syntax

 _expression_ . **Export**( **_FileName_** )

 _expression_ A variable that represents a **Page** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FileName_|Required| **String**|The fully qualified path and name of the file to receive the exported object.|

### Return Value

Nothing


## Remarks

The file name extension indicates which export filter to use. If the filter is not installed, the  **Export** method returns a compiler error in your Visual Basic or VBA project. The **Export** method uses the default preference settings for the specified filter and does not prompt the user for non-default arguments.

The  **Export** method of a **Page** object supports saving to HTML file format using the extension .htm or .html. When pages are exported, Visio uses the settings that were last selected in the **Save As** dialog box.

If the specified file already exists, Visio replaces it without prompting the user.

Starting with Visio, you can use various properties and methods of the  **[ApplicationSettings](f2e24211-ecc6-e0f5-4c00-fc50f98a3505.md)** object that relate to raster images to configure settings for export to .bmp, .gif, .jpg, .png, and .tif file types.

If your Visual Studio solution includes the  **Microsoft.Office.Interop.Visio** reference, this method maps to the following types:


-  **Microsoft.Office.Interop.Visio.IVPage.Export(string)**
    

## Example

This example shows how to export a Visio page as a bitmap (.bmp) file. It assumes a drawing page is open and active in Microsoft Visio.


```vb
Public Sub Export_Example() 
    Dim vsoPage As Visio.Page 
    Set vsoPage = ActivePage 
    vsoPage.Export ("C:\\myExportedPage.bmp") 
End Sub
```

