
# Series Members (Excel)
Represents a series in a chart.

 **Last modified:** July 28, 2015


## Methods



|**Name**|**Description**|
|:-----|:-----|
| [ApplyDataLabels](959a4d12-ed48-48fc-04cf-7a1880cd7e1f.md)|Applies data labels to a series.|
| [ClearFormats](0c94178c-493b-9738-3b85-67448d13a458.md)|Clears the formatting of the object.|
| [Copy](4a9261ae-9ad9-b591-f326-6f78e42637bf.md)|If the series has a picture fill, then this method copies the picture to the Clipboard.|
| [DataLabels](bde8faa1-269c-1dbe-e39e-3701a634f214.md)|Returns an object that represents either a single data label (a  ** [DataLabel](bb342572-8761-b326-548a-98455172f9a8.md)** object) or a collection of all the data labels for the series (a ** [DataLabels](3d79271e-c702-e785-6984-d838d060a8c5.md)**collection).|
| [Delete](931e1d33-aa05-6461-d5f3-4246925f5850.md)|Deletes the object.|
| [ErrorBar](0f127c27-09d3-a0e0-7a1d-5e3544039658.md)|Applies error bars to the series.  **Variant**.|
| [Paste](73e689cb-b2aa-61d7-e84c-113091d09a44.md)|Pastes a picture from the Clipboard as the marker on the selected series.|
| [Points](9b6f08a1-3fbe-e9bc-a509-345a3d2d78b3.md)|Returns an object that represents a single point (a  ** [Point](48ed9aec-2d29-ec4d-8e55-fca13982c358.md)** object) or a collection of all the points (a ** [Points](918dc385-ed61-262e-033f-ba829f5ee8b2.md)**collection) in the series. Read-only|
| [Select](9317a166-df2d-0c06-b1fb-4e3ecc7a645e.md)|Selects the object.|
| [Trendlines](d42609e1-011c-6cb3-286d-192284cd8ab8.md)|Returns an object that represents a single trendline (a  ** [Trendline](5c04b065-57f4-a059-7c22-50612bd727ea.md)** object) or a collection of all the trendlines (a ** [Trendlines](752cde45-c628-7550-6c88-07405821e348.md)**collection) for the series.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
| [Application](b36ee17b-a3dd-7458-6b65-8fb7723f41bd.md)|When used without an object qualifier, this property returns an  ** [Application](19b73597-5cf9-4f56-8227-b5211f657f6f.md)**object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an  **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
| [ApplyPictToEnd](40d4dca5-1747-c9c6-a117-29939bf4cd74.md)| **True** if a picture is applied to the end of the point or all points in the series. Read/write **Boolean**.|
| [ApplyPictToFront](b40a8808-734f-a00e-3fa1-4e1cf5ac0a52.md)| **True** if a picture is applied to the front of the point or all points in the series. Read/write **Boolean**.|
| [ApplyPictToSides](300e9c75-4108-32bc-01ab-c622843e6fbd.md)| **True** if a picture is applied to the sides of the point or all points in the series. Read/write **Boolean**.|
| [AxisGroup](0d5c9331-667a-e3d2-ff33-3ff353bd4c8d.md)|Returns or sets the group for the specified series. Read/write|
| [BarShape](27af7eea-6ad3-e906-c5f8-d9e29314b32d.md)|Returns or sets the shape used with the 3-D bar or column chart. Read/write  ** [XlBarShape](63a7cea6-e741-8e5b-94f3-16acfe22cb34.md)**.|
| [BubbleSizes](41e56271-ec4c-7f9e-9642-174c8435e7d6.md)|Returns or sets a string that refers to the worksheet cells containing the x-value, y-value and size data for the bubble chart. When you return the cell reference, it will return a string describing the cells in A1-style notation. To set the size data for the bubble chart, you must use R1C1-style notation. Applies only to bubble charts. Read/write  **Variant**.|
| [ChartType](5eff6ce3-1cba-eb92-0039-59f2ab65ddbc.md)|Returns or sets the chart type. Read/write  ** [XlChartType](bba4ee89-ee91-f55a-d2e0-59a73e5bfabe.md)**.|
| [Creator](f0c855a2-6901-be4f-13e2-426b97d34ef8.md)|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long**.|
| [ErrorBars](1a607e6f-e70a-e39c-4cc3-6060eb64e654.md)|Returns an  ** [ErrorBars](646de974-bf6f-99c8-20dd-9ca514b7a304.md)**object that represents the error bars for the series. Read-only.|
| [Explosion](e70678f5-ee1a-f5c2-7e5f-0c26f5282be4.md)|Returns or sets the explosion value for a pie-chart or doughnut-chart slice. Returns 0 (zero) if there's no explosion (the tip of the slice is in the center of the pie). Read/write  **Long**.|
| [Format](786f242a-57a8-b856-e826-4548a15f8e98.md)|Returns the  ** [ChartFormat](edac71b7-ed38-6658-2cbf-6493dc1ad3ed.md)** object. Read-only.|
| [Formula](c3b75251-55c0-150f-6a41-94d7f6444520.md)|Returns or sets a  **String** value that represents the object's formula in A1-style notation and in the language of the macro.|
| [FormulaLocal](6e2a0912-5006-d223-30a6-618642de035d.md)|Returns or sets the formula for the object, using A1-style references in the language of the user. Read/write  **String**.|
| [FormulaR1C1](d7b821f2-6e5c-21bc-b080-ddf666b466c4.md)|Returns or sets the formula for the object, using R1C1-style notation in the language of the macro. Read/write  **String**.|
| [FormulaR1C1Local](06037c27-3371-c2ac-4754-a5bb7ebb2058.md)|Returns or sets the formula for the object, using R1C1-style notation in the language of the user. Read/write  **String**.|
| [Has3DEffect](2bde474a-0e53-e435-d202-e97b23e90fd2.md)| **True** if the series has a three-dimensional appearance. Read/write **Boolean**.|
| [HasDataLabels](10f879c9-4d34-d20b-facc-44ebc950aaa2.md)| **True** if the series has data labels. Read/write **Boolean**.|
| [HasErrorBars](03d9a6e6-8c15-2bdb-1bca-ed9fb95e9cb3.md)| **True** if the series has error bars. This property isn't available for 3-D charts. Read/write **Boolean**.|
| [HasLeaderLines](9401e5a6-5acc-7503-02e6-b6dc42f381bb.md)| **True** if the series has leader lines. Read/write **Boolean**.|
| [InvertColor](889cef2a-8211-c1b2-0668-8e0c48a894ec.md)|Returns or sets the fill color for negative data points in a series. Read/write|
| [InvertColorIndex](fa2e87a4-57ad-395d-b631-fbca99560dae.md)|Returns or sets the fill color for negative data points in a series. Read/write|
| [InvertIfNegative](06c963ac-6e81-5f45-b8b9-8c61bf0c02b6.md)| **True** if Microsoft Excel inverts the pattern in the item when it corresponds to a negative number. Read/write **Boolean**.|
| [IsFiltered](90c1564c-439c-de1f-8082-0de3372c0566.md)|This setting controls whether the series has been filtered out from the chart. The default value is  **False**.  **Boolean** Read/Write.|
| [LeaderLines](d08a982c-8ac0-3f72-3f94-d72b3081f013.md)|Returns a  **LeaderLines** object that represents the leader lines for the series. Read-only.|
| [MarkerBackgroundColor](a7149fae-47a7-b24d-c177-28afde2ab29d.md)|Sets the marker background color as an RGB value or returns the corresponding color index value. Applies only to line, scatter, and radar charts. Read/write  **Long**.|
| [MarkerBackgroundColorIndex](90f57719-ff91-5b9c-6338-d238c6e234d6.md)|Returns or sets the marker background color as an index into the current color palette, or as one of the following  ** [XlColorIndex](b925578b-d654-61fa-03fa-67631ea8c5d1.md)** constants: **xlColorIndexAutomatic** or **xlColorIndexNone**. Applies only to line, scatter, and radar charts. Read/write  **Long**.|
| [MarkerForegroundColor](bdbb30c9-b997-7e7c-d592-cca04c2cfa71.md)|Sets the marker foreground color as an RGB value or returns the corresponding color index value. Applies only to line, scatter, and radar charts. Read/write  **Long**.|
| [MarkerForegroundColorIndex](6c13b34c-e21c-50d3-302f-ed234b7e2647.md)|Returns or sets the marker foreground color as an index into the current color palette, or as one of the following  ** [XlColorIndex](b925578b-d654-61fa-03fa-67631ea8c5d1.md)** constants: **xlColorIndexAutomatic** or **xlColorIndexNone**. Applies only to line, scatter, and radar charts. Read/write  **Long**.|
| [MarkerSize](d1e499ae-d59c-3493-c741-9607c3c27a17.md)|Returns or sets the data-marker size, in points. Can be a value from 2 through 72. Read/write  **Long**.|
| [MarkerStyle](fec57799-b01b-a8f8-2c26-1e7b11dd9777.md)|Returns or sets the marker style for a point or series in a line chart, scatter chart, or radar chart. Read/write  ** [XlMarkerStyle](404f138e-b3ed-556e-23e8-105114c2f66b.md)**.|
| [Name](64da2964-f896-a9f9-6c84-eeaa227ba86d.md)|Returns or sets a  **String** value representing the name of the object.|
| [Parent](5744fbd2-f82a-d488-8a1d-e93bc618bf59.md)|Returns the parent object for the specified object. Read-only.|
| [PictureType](098dac46-ec2d-ea2d-71e9-1094a5f0b23a.md)|Returns or sets a  ** [XlChartPictureType](7d4f70ea-4a66-1b88-49cf-85200c8eebff.md)** value that represents the way pictures are displayed on a column or bar picture chart.|
| [PictureUnit2](6c29fd60-2e42-3f7a-1fc0-67309ea75a3a.md)|Returns or sets the unit for each picture on the chart if the  ** [PictureType](098dac46-ec2d-ea2d-71e9-1094a5f0b23a.md)** property is set to **xlStackScale** (if not, this property is ignored). Read/write **Double**.|
| [PlotColorIndex](45bf641a-7b1e-1f0f-9662-5a903c08c2a1.md)|Returns an index value that is used internally to associate series formatting with chart elements. Read-only|
| [PlotOrder](c74ba422-ca4d-db60-02aa-7b512bdd0241.md)|Returns or sets the plot order for the selected series within the chart group. Read/write  **Long**.|
| [Shadow](8b1ad20a-764d-595b-5c3f-b7e6b68421a7.md)|Returns or sets a  **Boolean** value that determines if the object has a shadow.|
| [Smooth](24cb3fc6-a6ed-71ca-1aab-c1ea23480d00.md)| **True** if curve smoothing is turned on for the line chart or scatter chart. Applies only to line and scatter charts. Read/write.|
| [Type](18a300d5-ed08-af06-37ca-812b35d876ef.md)|Returns or sets a Long value that represents the series type.|
| [Values](3db2577e-ef0e-75ea-412b-531d7e67c098.md)|Returns or sets a  **Variant** value that represents a collection of all the values in the series.|
| [XValues](63715a3c-9d2d-6213-ac99-2c583773b45a.md)|Returns or sets an array of x values for a chart series. The  **XValues** property can be set to a range on a worksheet or to an array of values, but it cannot be a combination of both. Read/write **Variant**.|
