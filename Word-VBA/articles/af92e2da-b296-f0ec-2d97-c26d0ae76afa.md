
# ChartGroup Members (Word)
Represents one or more series plotted in a chart with the same format.

 **Last modified:** July 28, 2015


## Methods



|**Name**|**Description**|
|:-----|:-----|
| [CategoryCollection](63bd5ac0-15dc-16a3-843f-0b082bb81ea0.md)|Returns all the visible categories in the chart group, or the specified visible category.|
| [FullCategoryCollection](bba2ee13-b2db-9ed6-9581-b86dedfa51c9.md)|Returns all the categories in the chart group, or the specified category, whether visible or filtered out.|
| [SeriesCollection](4b4b7383-0967-cd2f-979c-eda9ef691459.md)|Returns all the series in the chart group.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
| [Application](3729126b-3431-98c8-8e8e-e76db2133145.md)|When used without an object qualifier, returns an  ** [Application](d1cf6f8f-4e88-bf01-93b4-90a83f79cb44.md)**object that represents the Microsoft Word application. When used with an object qualifier, returns an  **Application** object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.|
| [AxisGroup](4559ca47-ed2c-f122-0f38-a12cac8836d8.md)|Returns the type of axis group. Read/write  ** [XlAxisGroup](ed3ff1ce-28de-165d-bbfa-f3d770f32522.md)**.|
| [BubbleScale](4776723c-4d6e-1009-8d00-6f837fbd4803.md)|Returns or sets the scale factor for bubbles in the specified chart group. Read/write  **Long**.|
| [Creator](6c08be09-c7cd-ab41-3f75-fee9f26f6226.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.|
| [DoughnutHoleSize](5f4098ee-7d94-ace4-b412-1c7071434973.md)|Returns or sets the size of the hole in a doughnut chart group. Read/write  **Long**.|
| [DownBars](ee556f66-cce6-aa8d-a837-ee8b0b93ba89.md)|Returns the down bars on a line chart. Read-only  ** [DownBars](d0cf170e-0c58-2d01-a4b2-1eaf65dbfa3c.md)**.|
| [DropLines](eebe1c74-5682-4680-56d2-f0190fec5950.md)|Returns the drop lines for a series on a line chart or area chart. Read-only  ** [DropLines](4691b002-8512-7cd3-5a20-561232e18d88.md)**.|
| [FirstSliceAngle](0b5b9e0b-1210-6fc6-9e2c-2913cdb552cc.md)|Returns or sets the angle, in degrees (clockwise from vertical), of the first pie-chart or doughnut-chart slice. Read/write  **Long**.|
| [GapWidth](7f8d7f6b-9086-19c2-c4f4-d947491631ec.md)|For bar and column charts, returns or sets the space, as a percentage of the bar or column width, between bar or column clusters. For pie-of-pie and bar-of-pie charts, returns or sets the space between the primary and secondary sections of the chart. Read/write  **Long**.|
| [Has3DShading](095f5bc7-86aa-2c09-c52c-6e6d5a4deb16.md)| **True** if a chart group has three-dimensional shading. Read/write **Boolean**.|
| [HasDropLines](34743dd3-73f6-d125-a240-23984d31fa47.md)| **True** if the line chart or area chart has drop lines. Read/write **Boolean**.|
| [HasHiLoLines](5713e885-9f36-6b6c-2622-a813cba2077b.md)| **True** if the line chart has high-low lines. Read/write **Boolean**.|
| [HasRadarAxisLabels](0b086c3c-1eaa-1e65-fcb1-969c8b2c64c7.md)| **True** if a radar chart has axis labels. Read/write **Boolean**.|
| [HasSeriesLines](56e85d95-4743-4afd-5bdf-d00065608708.md)| **True** if a stacked column chart or bar chart has series lines or if a pie-of-pie chart or bar-of-pie chart has connector lines between the two sections. Read/write **Boolean**.|
| [HasUpDownBars](9c39f015-f8cc-633c-54a0-b68fc420d8f6.md)| **True** if a line chart has up and down bars. Read/write **Boolean**.|
| [HiLoLines](452f4e5d-7ad8-76ad-5067-2df8a074d6d1.md)|Returns the high-low lines for a series on a line chart. Read-only  ** [HiLoLines](9f1ed891-7e95-8dd0-745a-ce28555284a9.md)**.|
| [Index](cdc754ff-90de-d8c8-ff07-2f00a55ea959.md)|Returns the index number of the object within the collection of similar objects. Read-only  **Long**.|
| [Overlap](e2d219f6-7edd-69c7-015f-8304cf95dbc3.md)|Specifies how bars and columns are positioned. Read/write  **Long**.|
| [Parent](2f7f4f9f-412a-49cc-9c8c-39885f10c6a9.md)|Returns the parent for the specified object. Read-only  **Object**.|
| [RadarAxisLabels](30b37487-bef9-b333-7df7-546d85a92047.md)|Returns the radar axis labels for the specified chart group. Read-only  ** [TickLabels](d94e90dc-0b0e-f4af-078e-6f2b97729db5.md)**.|
| [SecondPlotSize](68f4d170-62c8-eb34-26a2-693aa96fc5f1.md)|Returns or sets the size, as a percentage of the primary pie, of the secondary section of either a pie-of-pie chart or a bar-of-pie chart. Read/write  **Long**.|
| [SeriesLines](23f36b19-99ed-f4d5-23b5-a8cd35bbf75c.md)|Returns the series lines for a 2-D stacked bar, 2-D stacked column, pie-of-pie, or bar-of-pie chart. Read-only  ** [SeriesLines](7521c592-c5aa-8e50-6268-840a41b3a282.md)**.|
| [ShowNegativeBubbles](6102a2dd-2ef8-2047-f14a-ca64e88a0565.md)| **True** if negative bubbles are shown for the chart group. Read/write **Boolean**.|
| [SizeRepresents](9611e92a-725c-fbe8-41bf-ef57d2166e4d.md)|Returns or sets what the bubble size represents on a bubble chart. Read/write  **Long**.|
| [SplitType](0bebc2f8-4dd6-8a74-993b-9e16357f38d0.md)|Returns or sets the way the two sections of either a pie-of-pie chart or a bar-of-pie chart are split. Read/write  ** [XlChartSplitType](8305530b-c62d-acaf-9b74-8c67797e2339.md)**.|
| [SplitValue](102826a5-834e-1b23-9888-6fb9b193ac96.md)|Returns or sets the threshold value separating the two sections of either a pie-of-pie chart or a bar-of-pie chart. Read/write  **Variant**.|
| [UpBars](8581ad5f-94a1-0e12-3880-14ce2a7e9f03.md)|Returns the up bars on a line chart. Read-only  ** [UpBars](22dff1d2-8f1b-8c48-354c-570906e0f830.md)**.|
| [VaryByCategories](e7ee35a4-ddb7-83ef-3c9b-0076f601bb19.md)| **True** if Microsoft Word assigns a different color or pattern to each data marker. Read/write **Boolean**.|
