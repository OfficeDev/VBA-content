
# TimelineState Members (Excel)
The timeline-specific state of a  [SlicerCache Object (Excel)](6e6533e3-0503-a1d3-9ecd-f7997233565f.md) object.

 **Last modified:** July 28, 2015


## Methods



|**Name**|**Description**|
|:-----|:-----|
| [SetFilterDateRange](c0ceea5c-9aa2-39a2-ce58-e37befeb0175.md)|Sets the Timeline's filter.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
| [Application](5b919557-9aeb-acc7-f717-8457f57e44fb.md)|Returns an  ** [Application](19b73597-5cf9-4f56-8227-b5211f657f6f.md)** object that represents the Microsoft Excel application. Read-only.|
| [Creator](aa6e35bb-531c-f501-23ef-f727db51f320.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.|
| [EndDate](1d33ce70-32ed-a439-eb34-7305fd9557f2.md)|Returns the end of the filtering date range (equals to  [TimelineState.StartDate Property (Excel)](3de8df53-1a36-428e-50dd-c7f45aa73b25.md) if range is a single day). **Variant** Read-only|
| [FilterType](8ba72a5e-0b0b-2d15-ccea-fb2cda537aae.md)|Returns the type of the date filter.  [XlPivotFilterType Enumeration (Excel)](0ae3f0fe-02e3-b0f7-1506-1961c4adcd6c.md) Read-only|
| [FilterValue1](6e10c4c3-465c-e097-8b3d-a76f8e2594e0.md)|Returns the 1st value associated with the date filter (semantics vary by filter type).  **Variant** Read-only|
| [FilterValue2](c48ba531-70fd-25db-e61f-a8cccd99ca82.md)|Returns the 2nd value associated with the date filter (semantics vary by filter type).  **Variant** Read-only|
| [Parent](2d7c5eb8-dbf8-9c71-8606-06b665094ac7.md)|Returns an  **Object** that represents the parent object of the specified [TimelineState Object (Excel)](bb92fe09-3cce-8e10-3795-2b9089c27801.md) object. Read-only.|
| [SingleRangeFilterState](aca37428-83e9-cb54-f32a-675dfcac5d9f.md)| **True** when the filtering state is a contiguous date range; **False** otherwise. **Boolean** Read-only|
| [StartDate](3de8df53-1a36-428e-50dd-c7f45aa73b25.md)|Returns the start of the filtering date range.  **Variant** Read-only|
