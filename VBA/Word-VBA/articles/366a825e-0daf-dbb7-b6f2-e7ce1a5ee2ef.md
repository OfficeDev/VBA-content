
# Chart Object (Word)

Represents a chart in a document.


## Remarks

The Example section describes the following properties and methods for returning a  **Chart** object:




- The  **[Chart](33d577fe-58b0-8e2f-a859-5bd3b34ed892.md)** property.
    
- The  **[AddChart](http://msdn.microsoft.com/library/1b168e7b-543a-a817-51b0-8171beecc946%28Office.15%29.aspx)** method.
    



## Example

The  **[InlineShapes](88c632b2-80de-c96a-8879-a98461b38bd0.md)** collection contains an object for each inline shape, including charts, in a document. Use **InlineShapes** ( _Index_ ), where Index is the index number of an inline shape, to return a single **InlineShape** object. Use the **[HasChart](f8b88eef-ec41-fc03-f58b-e346d240a121.md)** property to determine whether the **InlineShape** object represents a chart. If the **HasChart** property is set to **True**, use the **[Chart](33d577fe-58b0-8e2f-a859-5bd3b34ed892.md)** property to return a **Chart** object.

You can also use the  **[Type](0f85b99c-025b-9dff-b4f2-b74ab627efcc.md)** property to determine whether the **InlineShape** object represents a chart. If the **Type** property is set to **WdInlineShapeChart**, the inline shape represents a chart.

The following example verifies whether the first inline shape in the active document represents a chart. If so, the example changes the fore color of the first series on the chart.




```
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SeriesCollection(1).Format.Fill.ForeColor.RGB = rgbRed 
 End If 
End With
```

The following example creates a new 3-D column chart and adds it to the active document.




```
ActiveDocument.InlineShapes.AddChart Type:=xl3DColumn 

```


## Methods



|**Name**|
|:-----|
|[ApplyChartTemplate](10d2c95e-1f67-1301-9b98-3a0b09f60df5.md)|
|[ApplyDataLabels](3d4ce40f-7194-ad96-4ae6-15434c6dd491.md)|
|[ApplyLayout](f23d8a12-65d5-3336-4381-76bfc4b73507.md)|
|[Axes](37f422b5-31f2-92ce-c04e-a837b0a3d407.md)|
|[ChartWizard](5c4c4cb1-3ef7-e3c3-d441-6f92cb8e7771.md)|
|[ClearToMatchColorStyle](6bdf902e-61af-8cce-9925-2b8e943617b0.md)|
|[ClearToMatchStyle](33ea5fc1-9a71-a8d6-e714-91ff69c506b3.md)|
|[Copy](2343456a-0f47-bed5-f931-0b02b6ef8db1.md)|
|[CopyPicture](90f41c1a-8a96-0959-6c9a-b10f7f4744b0.md)|
|[Delete](ed16a8dc-6470-27ed-12d7-ab6e9ff06fe8.md)|
|[Export](49660450-ae9f-c59e-8974-b04327a72dc0.md)|
|[FullSeriesCollection](5fed6ff0-0dad-87ba-4db5-21ae919f25fd.md)|
|[GetChartElement](e9ebb101-1625-9a6a-1da9-dbb02c49f01c.md)|
|[Paste](e159d28e-c2ff-9105-3b52-278fe55b078c.md)|
|[Refresh](1f53620e-1a79-b932-bbf2-2a6cd95d524c.md)|
|[SaveChartTemplate](d980f663-7e73-7b55-9f7c-1fc9da84c0bd.md)|
|[Select](1ad91c5a-26a2-a7ad-faa6-c824245482bb.md)|
|[SeriesCollection](b9688aef-839a-b45b-1596-d8f02225aa05.md)|
|[SetBackgroundPicture](6bc2d271-86dd-cd4f-a7b8-323f6f7fe332.md)|
|[SetDefaultChart](e914b44a-5de9-ca9d-a513-96943802a194.md)|
|[SetElement](d172a9df-b081-0077-18ef-f75bf0d6f26a.md)|
|[SetSourceData](8c5b056a-6680-7e4e-ce67-a3b76b2d7d25.md)|

## Properties



|**Name**|
|:-----|
|[Application](a76fdfbb-1f9f-18b7-6127-fb7a85a6e8ed.md)|
|[AutoScaling](911bf146-f3fa-7c05-a0eb-9e2062ed4a93.md)|
|[BackWall](39ed0473-7d45-0584-48f1-307f9a481789.md)|
|[BarShape](e29af332-162c-4a9e-0281-f546bd00f27c.md)|
|[CategoryLabelLevel](74f01367-c625-94a8-4a0f-6bbc42dade3c.md)|
|[ChartArea](b16d78c0-7663-3ef9-c17a-02e7a024b344.md)|
|[ChartColor](d0f33ca3-90e5-c8d6-d2ac-117ea62ae7cc.md)|
|[ChartData](d8234dd3-148f-b69a-8a4e-f22474080eab.md)|
|[ChartGroups](ae4da68e-1e80-f683-b1ef-eb26aa753420.md)|
|[ChartStyle](53db7507-4fbf-15af-ea31-7ce5466f58f5.md)|
|[ChartTitle](1804d06a-bb2b-5995-7750-2ada70ddd1d4.md)|
|[ChartType](ad75b5bc-b323-8f67-cf1a-b4d6b6969eed.md)|
|[Creator](24057d70-7bab-728d-92de-3670b9e0e392.md)|
|[DataTable](1cae3588-5bc4-5ec4-c3f3-cc642d0755a6.md)|
|[DepthPercent](fd1a83dc-e68d-82be-d2bf-5f7a87cb08ac.md)|
|[DisplayBlanksAs](573752ec-7c2a-a5e0-bd05-626c81fb5d48.md)|
|[Elevation](2913dce4-e35a-31ff-3ea0-237c85dbad23.md)|
|[Floor](1544e584-3b0f-92a8-cc8f-6b12ed66109e.md)|
|[GapDepth](09147a74-c8bb-4fc5-0389-c8f46e0be67d.md)|
|[HasAxis](b5b7effe-48c6-75d9-fdc4-7a9ff148f0e9.md)|
|[HasDataTable](62af9540-9a69-0e19-b884-4f2b5947152f.md)|
|[HasLegend](057fedc3-4f23-9c28-3196-836523d83656.md)|
|[HasTitle](5995f349-3809-e842-69a6-e9227b731021.md)|
|[HeightPercent](b05873d9-a7b9-8980-28e7-057a90f7bb94.md)|
|[Legend](b1ffdbfb-854c-bd65-dd63-d3b8d0547f67.md)|
|[Parent](1763bd1d-04bd-d4fc-b304-c52c084100b3.md)|
|[Perspective](d88ab2dc-822a-de51-a2b9-bcce667cd0e2.md)|
|[PivotLayout](adf7d083-4f81-92f8-887d-5555d553d35d.md)|
|[PlotArea](440f7d57-c681-098e-45d6-a2f7aca6de07.md)|
|[PlotBy](ae2774d0-0f58-2224-9104-61d00fa63a86.md)|
|[PlotVisibleOnly](59b7f58e-a1b2-56cd-89e8-529228d2979c.md)|
|[RightAngleAxes](d7f01a8f-aa76-3e92-2db2-370176066145.md)|
|[Rotation](a141124f-f33c-95e1-6ba9-8ecffdef434c.md)|
|[SeriesNameLevel](e77240d4-273c-460e-d10a-c43f67f6f955.md)|
|[Shapes](bbc30f17-b984-683f-cd6d-9080f3c29897.md)|
|[ShowAllFieldButtons](95ad77fa-fef3-3927-0f0f-9e6fd7701316.md)|
|[ShowAxisFieldButtons](08ee0734-d5b9-b57a-fa5f-ffa1c5ded498.md)|
|[ShowDataLabelsOverMaximum](3a460343-126c-5d83-38c2-c7fe7d2c59d1.md)|
|[ShowLegendFieldButtons](da28865f-d513-3f43-45e7-d1cb25cda18c.md)|
|[ShowReportFilterFieldButtons](716bcdfb-0e94-85c3-1a3d-2da6a6867f36.md)|
|[ShowValueFieldButtons](9b650a6f-8cdb-9aef-d19e-6a2e339e7768.md)|
|[SideWall](dd1ededa-f19a-d0b8-4e88-4af4720c7768.md)|
|[Walls](f45ae75a-c96c-4441-af81-aedf23787194.md)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)