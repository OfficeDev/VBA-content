---
title: "Свойство ShapeRange.ConnectorFormat (издатель)"
keywords: vbapb10.chm2293814
f1_keywords: vbapb10.chm2293814
ms.prod: publisher
api_name: Publisher.ShapeRange.ConnectorFormat
ms.assetid: 1a1516bd-ef27-0b37-09dd-45af8a531a76
ms.date: 06/08/2017
ms.openlocfilehash: 2c4bc085f7d73d2c1c4b8ee9a107d80ca4fd0a14
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shaperangeconnectorformat-property-publisher"></a>Свойство ShapeRange.ConnectorFormat (издатель)

Возвращает объект **[ConnectorFormat](connectorformat-object-publisher.md)** , который содержит соединитель свойства форматирования. Применяется к **фигуры** или **ShapeRange** объектов, представляющих соединители.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ConnectorFormat**

 переменная _expression_A, представляющий объект **ShapeRange** .


## <a name="example"></a>Пример

В этом примере добавляется два прямоугольника для первой страницы в активной публикации и связывает их с искривленной формы.


```vb
Dim shpRect1 As Shape 
Dim shpRect2 As Shape 
 
With ActiveDocument.Pages(1).Shapes 
 
 ' Add two new rectangles. 
 Set shpRect1 = .AddShape(Type:=msoShapeRectangle, _ 
 Left:=100, Top:=50, Width:=200, Height:=100) 
 Set shpRect2 = .AddShape(Type:=msoShapeRectangle, _ 
 Left:=300, Top:=300, Width:=200, Height:=100) 
 
 ' Add a new curved connector. 
 With .AddConnector(Type:=msoConnectorCurve, _ 
 BeginX:=0, BeginY:=0, EndX:=100, EndY:=100) _ 
 .ConnectorFormat 
 
 ' Connect the new connector to the two rectangles. 
 .BeginConnect ConnectedShape:=shpRect1, ConnectionSite:=1 
 .EndConnect ConnectedShape:=shpRect2, ConnectionSite:=1 
 
 ' Reroute the connector to create the shortest path. 
 .Parent.RerouteConnections 
 End With 
 
End With 

```


