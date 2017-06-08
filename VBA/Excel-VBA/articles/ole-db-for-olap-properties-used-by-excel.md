---
title: OLE DB for OLAP Properties Used by Excel
ms.prod: excel
ms.assetid: 5caa2240-1f7b-08d7-c87b-ec30f3bcb441
ms.date: 06/08/2017
---


# OLE DB for OLAP Properties Used by Excel

Microsoft Excel uses an OLE DB for OLAP (OnLine Analytical Processing) provider to connect to OLAP cubes. When connecting to an OLAP cube, Excel reads and sets various OLE DB and OLE DB for OLAP properties. Excel considers Connection properties, Data Source Information Properties, Schema Rowset Queries, and Cell properties. 

Rather than address all the existing properties, this topic focuses on the properties that have a unique relationship with Excel. OLAP connections in Excel are used for PivotTables and OLAP Formulas. When you are testing an existing OLAP provider, it is recommended that you have Excel read a set of provider properties to determine whether an OLAP provider supports the features required for OLAP PivotTable design and functionality. If the provider does not support certain capabilities, the features that depend on these capabilities are either disabled or limited. Other properties are set in order to get desired behavior, and if these properties are not implemented for an OLAP provider, Excel might not work with it.

## Connection Properties



|**Property Set**|**Property**|**Set if**|**Set to**|
|:-----|:-----|:-----|:-----|
|DBPROPSET_MSOLAPINIT|DBPROP_MSMD_SAFETY_OPTIONS|Supported|OLAPUDFSecurity reg key or DBPROPVAL_MSMD_SAFETY_OPTIONS_ALLOW_SAFESee also: [Safety Options Property](http://msdn.microsoft.com/en-us/library/aa237323.aspx).|
|DBPROPSET_MSOLAPINIT|DBPROP_MSMD_MDXCOMPATIBILITY|Supported|DBPROP_MSMD_MDXCOMPATIBILITY_70See also: [MDX Compatibility Property](http://msdn.microsoft.com/en-us/library/aa256070.aspx).|
|DBPROPSET_MSOLAPINIT|DBPROP_MSMD_SOURCE_DSN_SUFFIX|DBPROP_MSMD_SOURCE_DSN in DBPROPSET_MSOLAPINIT is present|String "Prompt=CompleteRequired;Window Handle=0x<hwnd>"See also: [Source_DSN_Suffix Property](http://msdn.microsoft.com/en-us/library/aa237431.aspx).|
|DBPROPSET_MSOLAPINIT|DBPROP_MSMD_MDX_MISSING_MEMBER_MODE|Supported|If property is supported, Excel sets it to the string "Error". Ignored if not supported.|
|DBPROPSET_DBINIT| [DBPROP_INIT_LCID](http://msdn.microsoft.com/en-us/library/ms719750.aspx)|Supported|Set before making the connection. It is possible to specify any LCID to be used in the connection. If translations are turned on for the connection, Excel sets this to the UI language (default). If property is not supported, Excel has no problem other than losing the functionality of having translations based on UI language.|
|DBPROPSET_DBINIT| [DBPROP_INIT_PROMPT](http://msdn.microsoft.com/en-us/library/ms714342.aspx)|Supported|Not OLAP specific. Set before making the connection. If setting this property fails because a certain value is not supported, Excel ignores the failure.|
|DBPROPSET_DBINIT| [DBPROP_AUTH_PERSIST_SENSITIVE_AUTHINFO](http://msdn.microsoft.com/en-us/library/ms714905.aspx)|Supported|Not OLAP specific. Set before making the connection. Excel appears to always set this to True.|
|DBPROPSET_DBINIT| [DBPROP_INIT_HWND](http://msdn.microsoft.com/en-us/library/ms715949.aspx)|Supported|Not OLAP specific. Set before making the connection. Excel sets this to the main window of the application so the provider displays the alert using the correct parent window.|
|DBPROPSET_DBINIT| [DBPROP_INIT_ASYNCH](http://msdn.microsoft.com/en-us/library/ms711533.aspx)|Supported|Not OLAP specific.Set before making the connection. Excel sets this property to DBPROPVAL_ASYNCH_INITIALIZE based on a registry setting (you can also disable it by using a registry setting). If property is not supported, Excel ignores it and does not set it.|
|DBPROPSET_DBINIT|DBPROP_CMD_PROMPT|Supported|Not OLAP specific. Set before making the connection.|
|DBPROPSET_DBINIT|DBPROP_CMD_HWND|Supported|Not OLAP specific. Set before making the connection.|

## Data Source Information



|**Property Set**|**Property**|**Value**|**Use**|
|:-----|:-----|:-----|:-----|
|DBPROPSET_MDX_EXTENSIONS|DBPROP_MSMD_MDX_DDL_EXTENSIONS|If bit set for DBPROPVAL_MDX_DLL_CREATESESSIONCUBE.|The grouping feature of OLAP PivotTables is enabled if  `CREATE SESSION CUBE` is supported.|
|DBPROPSET_MDX_EXTENSIONS|DBPROP_MSMD_MDX_DDL_EXTENSIONS|If bit set for DBPROPVAL_MDX_DDL_REFRESHCUBE.|If  `REFRESH CUBE` command is supported, Excel executes it when an OLAP PivotTable is refreshed.|
|DBPROPSET_MDX_EXTENSIONS|DBPROP_MSMD_MDX_CALCMEMB_EXTENSIONS|If bit set for DBPROPVAL_MDX_CALCMEMB_ADD.|The show calculated members feature in OLAP PivotTable is enabled if  `ADDCALCULATEDMEMBERS` is supported in MDX (Multidimensional Expressions).|
|DBPROPSET_DATASOURCEINFO| [MDPROP_MDX_FORMULAS ](http://msdn.microsoft.com/en-us/library/ms709719.aspx)|If both bits set MDPROPVAL_MF_SCOPE_SESSION, MDPROPVAL_MF_CREATE_CALCMEMBERS.|If the provider supports creating session members ( `CREATE SESSION MEMBER`), Excel enables this feature in OLAP PivotTables (only available in the object model in Excel).|
|DBPROPSET_SESSION|DBPROP_VISUALMODE|If supported (and subselect not supported, see MDPROP_MDX_SUBQUERIES below). |Enables control of Include hidden items in totals (toggle visual totals).|
|DBPROPSET_DATASOURCEINFO|MDPROP_MDX_SUBQUERIES|If the two lowest bits are set (with this, Excel does not support non-visual totals, see DBPROP_VISUALMODE above).|Enables Label, Date, and Value filtering in Excel PivotTables. Generally uses Excel MDX query construction. Note that this property is introduced with SQL Server 2005 Service Pack 2. Value is always  `VARIANT_TRUE` in msolap90.dll.|
|DBPROPSET_DATASOURCEINFO|MDPROP_MDX_DRILL_FUNCTIONS||If the two lowest bits of this property are set, Excel interprets it as the server supporting tuple-based drilling with the  `DrillDownLevel` and `DrillDownMember` functions.However, Excel only allows attribute drilling if the lowest two bits of  `MDPROP_MDX_SUBQUERIES` are also set (subselects supported).|
|DBPROPSET_DATASOURCEINFO| [MDPROP_FLATTENING_SUPPORT](http://msdn.microsoft.com/en-us/library/ms720900.aspx)|Check that it is set to MDPROPVAL_FS_FULL_SUPPORT.|Read by Excel, and if it is not set to  `MDPROPVAL_FS_FULL_SUPPORT`, an error occurs because Excel does not consider it an OLAP provider.|
|DBPROPSET_DATASOURCEINFO| [MDPROP_NAMED_LEVELS](http://msdn.microsoft.com/en-gb/library/ms713691.aspx)|Excel checks that the lowest bit is set (MDPROPVAL_NL_NAMEDLEVELS).|If the lowest bit of this property is not set, Excel fails.|
|DBPROPSET_DATASOURCEINFO| [MDPROP_MDX_SET_FUNCTIONS](http://msdn.microsoft.com/en-us/library/ms711600.aspx)||Excel queries for this property, but it has no feature-relevant effect.|
|DBPROPSET_DATASOURCEINFO| [DBPROP_DBMSVER](http://msdn.microsoft.com/en-us/library/ms719676.aspx)|Excel checks whether this value is a string.|Excel does not check the actual value of this property; it only verifies whether it is a string. If it is not a string, Excel fails to connect.|
|DBPROPSET_DATASOURCEINFO| [DBPROP_DATASOURCE_TYPE](http://msdn.microsoft.com/en-us/library/ms722595.aspx)|Excel checks whether the second lowest bit is set (DBPROPVAL_DST_MDP).|If the lowest bit is set, the provider is considered a multidimensional (OLAP) provider.|
|DBPROPSET_ROWSET| [DBPROP_ROWSET_ASYNCH](http://msdn.microsoft.com/en-us/library/ms717927.aspx)|If supported.|Excel tries to set this to  `DBPROPVAL_ASYNCH_INITIALIZE` but if this fails, Excel falls back into synchronous mode.If supported, it enables Excel to support the user pressing the  **Esc** key to stop query execution before it is finished.|

## Schema Rowset Queries



|**Schema Rowset**|**Column**|**Value**|**Controls**|
|:-----|:-----|:-----|:-----|
| [MDSCHEMA_CUBES](http://msdn.microsoft.com/en-us/library/aa179343.aspx)|IS_DRILLTHROUGH_ENABLED|TRUE|If set to TRUE, the drill-through (Show Details) feature is enabled for cells in the OLAP PivotTable values area.|
| [MDSCHEMA_HIERARCHIES](http://msdn.microsoft.com/en-us/library/aa179350.aspx)|STRUCTURE|MD_STRUCTURE_UNBALANCED|Excel has special handling of filtering for unbalanced hierarchies, so these are marked as such for control purposes.|
| [MDSCHEMA_HIERARCHIES](http://msdn.microsoft.com/en-us/library/ms126062.aspx)|HIERARCHY_ORIGIN|MD_ORIGIN_ATTRIBUTE set and not MD_ORIGIN_USER_DEFINED|Excel has special handling of attribute hierarchies in OLAP PivotTables, so attribute hierarchies are marked as such.|
| [MDSCHEMA_HIERARCHIES](http://msdn.microsoft.com/en-us/library/ms126062.aspx)|HIERARCHY_DISPLAY_FOLDER||Based on this property, the PivotTable Field List displays hierarchies in folders under their dimensions.|
| [MDSCHEMA_MEASUREGROUPS](http://msdn.microsoft.com/en-us/library/ms126178.aspx)|MEASUREGROUP_NAME|| **Measures** are listed in a folder representing their measure group in the PivotTable Field List.|
| [MDSCHEMA_MEASUREGROUPS](http://msdn.microsoft.com/en-us/library/ms126178.aspx)|MEASUREGROUP_CAPTION|| **Measures** are listed in a folder representing their measure group with this caption in the PivotTable Field List.|
| [MDSCHEMA_SETS](http://msdn.microsoft.com/en-us/library/ms126290.aspx)|SET_DISPLAY_FOLDER||Excel reads the display folder property to enable it to place sets in display folders in the PivotTable Field List.|
| [MDSCHEMA_SETS](http://msdn.microsoft.com/en-us/library/ms126290.aspx)|SET_CAPTION||Excel reads the set caption for displaying in the PivotTable report and in the PivotTable Field List.|
| [MDSCHEMA_KPIS](http://msdn.microsoft.com/en-us/library/ms126258.aspx)|KPI_DISPLAY_FOLDER||KPIs (key performance indicators) defined on the server are listed in the PivotTable field list, and the components (value, goal, status, and trend) can be added to the values area. Excel reads this property to place the KPI in the correct display folder in the PivotTable Field List.|
| [MDSCHEMA_KPIS](http://msdn.microsoft.com/en-us/library/ms126258.aspx)|KPI_PARENT_KPI_NAME||Excel reads this property to place child KPIs in subfolders under their parent KPI in the PivotTable Field List (if display folders are defined, those are used instead).|
| [MDSCHEMA_KPIS](http://msdn.microsoft.com/en-us/library/ms126258.aspx)|KPI_TREND_GRAPHIC||Excel reads this property and, based on the value, maps it to the closest conditional formatting icon set in Excel when Trend is added to the PivotTable.|
| [MDSCHEMA_KPIS](http://msdn.microsoft.com/en-us/library/ms126258.aspx)|KPI_STATUS_GRAPHIC||Excel reads this property and, based on the value, maps it to the closest conditional formatting icon set in Excel when Status is added to the PivotTable.|
| [MDSCHEMA_ACTIONS](http://msdn.microsoft.com/en-us/library/ms126032.aspx)|||Additional Actions feature. Excel exposes server-defined actions in the shortcut menu of an OLAP PivotTable report when actions exist on the server for the selected context.|
| [MDSCHEMA_MEASURES](http://msdn.microsoft.com/en-us/library/ms126250.aspx)|MEASURE_DISPLAY_FOLDER||Read by Excel so it can place measures in the correct display folder in the PivotTable Field List.|
| [MDSCHEMA_MEASURES](http://msdn.microsoft.com/en-us/library/ms126250.aspx)|EXPRESSION||Read by Excel to determine whether a measure is calculated. If it is a string and not empty, Excel considers it a calculated measure.|
| [MDSCHEMA_PROPERTIES](http://msdn.microsoft.com/en-us/library/ms126309.aspx)|PROPERTY_NAME|"MEMBER_VALUE" This schema also used for getting regular member properties. The "MEMBER_VALUE" value is a special case, but there are other usage.|Excel gets the member value property of the key attribute in a dimension by restricting to "MEMBER_VALUE" in the PROPERTY_NAME column.If the data type (DATA_TYPE) of the MEMBER_VALUE property of the key attribute of a Time dimension is  **Date**, the PivotTable exposes date filtering instead of label filtering. The actual date filtering is done based on the member value property of the key independent of which hierarchy of that dimension is filtered.<table><tr><th>**Note**</th></tr><tr><td>Date filtering requires support for subselects (see MDPROP_MDX_SUBQUERIES above).</td></tr></table>|
|MDSCHEMA_DISCOVER|RESTRICTIONS||Depending on usage, Excel restricts on hierarchies, levels, or measures when reading the MDSCHEMA_DISCOVER rowset to get the RESTRICTIONS. Excel reads schema row by row and finds list of restrictions for all other relevant schemas to obtain the index of the restrictions that affect Excel. The RESTRICTIONS column has a chapter handle to another row set from which Excel looks at the NAME column. In the NAME column, Excel expects to find the strings HIERARCHY_VISIBILITY, MEASURE_VISIBILITY, LEVEL_VISIBILITY (if the provider supports restriction on visibility). If Excel cannot find <xxx&gt;_VISIBILITY strings (or if MDSCHEMA_DISCOVER is not supported) it will assume that provider doesn't support returning hidden items, and it will not query for them.|
| [MDSCHEMA_LEVELS](http://msdn.microsoft.com/en-us/library/ms126038.aspx)|LEVEL_ATTRIBUTE_HIERARCHY_NAME||Used by Excel to hide special grouping levels with system-generated names. Note that this is not needed with Microsoft SQL Server 2005 Analysis Services Service Pack 2.|
| [MDSCHEMA_LEVELS](http://msdn.microsoft.com/en-us/library/ms126038.aspx)|CUSTOM_ROLLUP_SETTINGS|0|If not 0, Excel assumes the level has custom rollup. Excel checks this for all levels of each hierarchy, and if custom rollup is present, some operations are disabled (such as grouping).|

## Cell Properties



|**Property Name**|**Use**|
|:-----|:-----|
| **Language**|<p>LCID for determining how to interpret  `FORMAT_STRING` when it is **CURRENCY**.</p> <p>Excel uses this property to determine which currency symbol to use when formatting values with  `FORMAT_STRING` set to **Currency**.</p>  <p>[Retrieving Cell Properties](http://msdn.microsoft.com/en-us/library/ms715853.aspx)</p><p>Example of calculated measure definition specifying the LANGUAGE property for the client application to pick up: </p><p>```CREATE MEMBER CURRENTCUBE.[Measures].[Internet Gross Profit]``` </p><p> ```AS``` </p><p>```[Measures].[Internet Sales Amount]```  </p><p>```-  ```</p><p>```[Measures].[Internet Total Product Cost], ```</p><p>``` ```</p><p>```FORMAT_STRING = "Currency", ```</p><p>```BACK_COLOR = 12615680 /*R=0, G=128, B=192*/, ```</p><p>```FORE_COLOR = 65408 /*R=128, G=255, B=0*/, ```</p><p>```FONT_FLAGS = 3 /*Bold, Italic*/, ```</p><p>```NON_EMPTY_BEHAVIOR = { [Internet Sales Amount],[Internet Total Product Cost] }, ```</p><p>```VISIBLE = 1, ```</p><p>```LANGUAGE = 1033 /*Telling client application to display US currency symbol*/;```</p>|



