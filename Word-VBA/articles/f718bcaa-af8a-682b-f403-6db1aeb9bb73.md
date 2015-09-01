
# Cell Members (Word)
Represents a single table cell. The  **Cell** object is a member of the ** [Cells](ceaa5b45-518d-d6ea-1ce8-5a34f6e37046.md)**collection. The  **Cells** collection represents all the cells in the specified object.

 **Last modified:** July 28, 2015


## Methods



|**Name**|**Description**|
|:-----|:-----|
| [AutoSum](5f8c36c3-2e26-8e5f-16c4-49d4c04144c1.md)|Inserts an = (Formula) field that calculates and displays the sum of the values in table cells above or to the left of the cell specified in the expression.|
| [Delete](01e6d989-e86c-9a3b-b0e3-d6eb1f2a7183.md)|Deletes a table cell or cells and optionally controls how the remaining cells are shifted.|
| [Formula](0fec018a-5a6f-f5ec-ed1c-a963e53c27b3.md)|Inserts an = (Formula) field that contains the specified formula into a table cell.|
| [Merge](79d929bd-9578-e937-405f-8ad970ae883c.md)|Merges the specified table cell with another table cell. The result is a single table cell.|
| [Select](d7228170-2b1f-51e2-9fc1-0cbfffa3b74d.md)|Selects the specified object.|
| [SetHeight](1c26425e-66f0-0558-5981-7161d730e8e1.md)|Sets the height of table cells.|
| [SetWidth](fd9fbeb1-a8c7-a6bf-1c9e-b63954848baf.md)|Sets the width of columns or cells in a table.|
| [Split](c7eb0d00-ff7e-a737-2083-e16f46ead256.md)|Splits a single table cell into multiple cells.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
| [Application](ccce55d3-b2ec-bd03-f1f5-46df97b5a07d.md)|Returns an  ** [Application](d1cf6f8f-4e88-bf01-93b4-90a83f79cb44.md)** object that represents the Microsoft Word application.|
| [Borders](a62d45e4-02ff-60ab-b0e6-93929cce64d1.md)|Returns a  ** [Borders](6dd1d4cc-2dcf-22c7-a299-4721a5543ba3.md)**collection that represents all the borders for the specified object.|
| [BottomPadding](5f265dc2-a9c4-d307-69a8-1f73407a4301.md)|Returns or sets the amount of space (in points) to add below the contents of a single cell or all the cells in a table. Read/write  **Single**.|
| [Column](b3f5f0a1-4d17-9d66-f689-9eb6308132fe.md)|Returns a  **Column** object that represents the table column containing the specified cell. Read-only.|
| [ColumnIndex](cb30b08a-b95f-da3f-ceae-7c83a5d2ec9e.md)|Returns the number of the table column that contains the specified cell. Read-only  **Long**.|
| [Creator](9a50df51-61ab-01d1-30fe-6c5f6622ce4c.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.|
| [FitText](ba600e01-1892-557d-95e8-fc9cdea8ef6b.md)| **True** if Microsoft Word visually reduces the size of text typed into a cell so that it fits within the column width. Read/write **Boolean**.|
| [Height](746d61a9-d3e2-c28d-3dac-a892c33be2c7.md)|Returns or sets the height of the specified table cell. .|
| [HeightRule](cff7f223-5f3f-c31f-e12a-3d28c96d47ec.md)|Returns or sets a  **WdRowHeightRule** constant that represents the rule for determining the height of the specified cells or rows. Read/write.|
| [ID](46c973be-38d4-18b3-ea4e-0d29d89313d7.md)|Returns or sets the identifying label for the specified object when the current document is saved as a Web page. Read/write  **String**.|
| [LeftPadding](b80dba74-7f12-0258-de03-e9941b6b1f4c.md)|Returns or sets the amount of space (in points) to add to the left of the contents of a single cell or all the cells in a table. Read/write  **Single**.|
| [NestingLevel](6eff7eac-72b9-1b33-af2c-0dd410576e92.md)|Returns the nesting level of the specified cell. Read-only  **Long**.|
| [Next](b4171c7c-6703-9cdf-a964-09e32874fbb6.md)|Returns a  **Cell** object that represents the next table cell in the **Cells** collection. Read-only.|
| [Parent](ef27abde-9789-52f2-ac30-b346404939d6.md)|Returns an  **Object** that represents the parent object of the specified **Cell** object.|
| [PreferredWidth](2b59ace4-bd3e-8a30-b81e-0f57d29f8a02.md)|Returns or sets the preferred width (in points or as a percentage of the window width) for the specified cell. Read/write  **Single**.|
| [PreferredWidthType](5880af18-b1a2-cb53-c224-147453e84f0e.md)|Returns or sets the preferred unit of measurement to use for the width of the specified cell. Read-only  **WdPreferredWidthType**.|
| [Previous](64bc6592-e7ae-15bc-456e-1ba0cb1b2935.md)|Returns a  **Cell** object that represents the previous table cell in the ** [Cells](ceaa5b45-518d-d6ea-1ce8-5a34f6e37046.md)** collection. Read-only.|
| [Range](579a25ad-91fa-a7c9-7eb8-4307521aeddd.md)|Returns a  ** [Range](15a7a1c4-5f3f-5b6e-60e9-29688de3f274.md)**object that represents the portion of a document that's contained in the specified object.|
| [RightPadding](6e71d162-7a8a-9ff2-38ec-c7867804d28b.md)|Returns or sets the amount of space (in points) to add to the right of the contents of a single cell or all the cells in a table. Read/write  **Single**.|
| [Row](b395a2f8-2eb4-1443-1298-56e3d3ad068b.md)|Returns a  ** [Row](38a05858-829a-ea5c-ce63-7f7343bf7b88.md)**object that represents the row containing the specified cell.|
| [RowIndex](745fabed-ba99-2e69-0d87-a7b520ac78cf.md)|Returns the number of the row that contains the specified cell. Read-only  **Long**.|
| [Shading](ab2f5789-ba6e-fa8a-d0a9-4c8b7922aa92.md)|Returns a  ** [Shading](e136509a-1be1-29e4-7b37-1faf659e37ba.md)**object that refers to the shading formatting for the specified object.|
| [Tables](2e18a6ae-590b-0f4f-41b5-cd34e15c9375.md)|Returns a  ** [Tables](068a3d0f-0b19-3927-cb0a-7fb0d0fd8e52.md)**collection that represents all the nested tables inside the specified table cell. Read-only.|
| [TopPadding](03c8bd07-dde2-6ad3-1291-7b0c0ada424a.md)|Returns or sets the amount of space (in points) to add above the contents of a single cell or all the cells in a table. Read/write  **Single**.|
| [VerticalAlignment](fc4308f0-755e-251b-f7f2-6d86b78dc0b0.md)|Returns or sets the vertical alignment of text in one or more cells of a table. Read/write  **WdCellVerticalAlignment**.|
| [Width](87c0422d-5f4f-44a3-902a-cb751b459ef9.md)|Returns or sets the width of a table cell, in points. Read/write  **Long**.|
| [WordWrap](16255023-d6c3-3c27-402f-490970b7af33.md)| **True** if Microsoft Word wraps text to multiple lines and lengthens the cell so that the cell width remains the same. Read/write **Boolean**.|
