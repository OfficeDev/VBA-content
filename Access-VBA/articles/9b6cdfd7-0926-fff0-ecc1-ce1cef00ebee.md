
# ChartObjects Members (Excel)
A collection of all the  ** [ChartObject](b546e6f2-7ac6-2dea-eba2-f98f68f3df65.md)** objects on the specified chart sheet, dialog sheet, or worksheet.

 **Last modified:** July 28, 2015


## Methods



|**Name**|**Description**|
|:-----|:-----|
| [Add](46f28b34-83a5-b3d9-c19b-a1dc8e05dff7.md)|Creates a new embedded chart.|
| [Copy](66e30b0c-a304-00fa-e573-e975c530c46c.md)|Copies the object to the Clipboard.|
| [CopyPicture](df79e18c-624b-424d-cd3e-d9432ed87aac.md)|Copies the selected object to the Clipboard as a picture.  **Variant**.|
| [Cut](842104f6-4317-8cac-5dd2-2ce2b1071052.md)|Cuts the object to the Clipboard.|
| [Delete](a39fca6c-1b6a-5693-b554-37788ec193c7.md)|Deletes the object.|
| [Duplicate](085e07e1-7b08-befb-1351-b9de3df26ddc.md)|Duplicates the object and returns a reference to the new copy.|
| [Item](0dbc6680-73ee-73a8-c3d8-f05faf6dd596.md)|Returns a single object from a collection.|
| [Select](ef89d037-34d4-3c17-edb7-352b52e5ae4b.md)|Selects the object.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
| [Application](2ff0a431-a796-e1c6-d15d-7e70aba1e426.md)|When used without an object qualifier, this property returns an  ** [Application](19b73597-5cf9-4f56-8227-b5211f657f6f.md)**object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an  **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
| [Count](28d3d9fd-cf58-8b95-3f14-c336bcee1bb5.md)|Returns a  **Long** value that represents the number of objects in the collection.|
| [Creator](8cfd1fc7-b6a8-5d1a-9dc8-58ca5521d3a8.md)|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long**.|
| [Height](a0801e22-cd20-9750-a69a-121be0fd9749.md)|Returns or sets a  **Double** value that represents the height, in points, of the object.|
| [Left](9d9b8505-3d6b-f37f-b35c-0a092721fe7a.md)|Returns or sets a  **Double** value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).|
| [Locked](6d9fc386-3dcc-c52f-d590-2749dac2378f.md)|Returns or sets a  **Boolean** value that indicates if the objects are locked.|
| [Parent](4c5453db-8e90-1ae0-2fb2-990c1d336f20.md)|Returns the parent object for the specified object. Read-only.|
| [Placement](954e98e5-8b88-6918-3cbd-f8e982c0a47e.md)|Returns or sets a  **Variant** value, containing an ** [XlPlacement](ad52cbf4-3d51-d9fe-5e31-be181f7775d3.md)** constant, that represents the way the objects are attached to the cells below them.|
| [PrintObject](310a4571-e5e4-14c8-56a0-6d70a59f4588.md)| **True** if the objects will be printed when the document is printed. Read/write **Boolean**.|
| [ProtectChartObject](e0685fbd-84a5-36c4-a5ab-06127937f2c8.md)| **True** if the embedded chart frame cannot be moved, resized, or deleted through the user interface. Read/write **Boolean**.|
| [ShapeRange](4813fce5-ad3f-861c-d6dc-63fb617ed4da.md)|Returns a  ** [ShapeRange](e1b8229c-73a0-4a77-5e00-4bcec9032260.md)** object that represents the specified object or objects. Read-only.|
| [Top](260fb609-ca58-61f8-44a9-d3183d7937f1.md)|Returns or sets a  **Double** value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).|
| [Visible](c7e1fad7-1ed3-d76b-f637-2dfda5fe9b53.md)|Returns or sets a  **Boolean** value that determines whether the object is visible. Read/write.|
| [Width](835cb1e6-937c-de90-af37-309b9bebb070.md)|Returns or sets a  **Double** value that represents the width, in points, of the object.|
