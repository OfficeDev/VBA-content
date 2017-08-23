---
title: "Свойство Shape.ConnectorFormat (издатель)"
keywords: vbapb10.chm2228278
f1_keywords: vbapb10.chm2228278
ms.prod: publisher
api_name: Publisher.Shape.ConnectorFormat
ms.assetid: 280c424c-530c-55ab-da4f-65b858ee3dd8
ms.date: 06/08/2017
ms.openlocfilehash: 86d38b9e40c251227e2d1a4199a916d298d012c6
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapeconnectorformat-property-publisher"></a>Свойство Shape.ConnectorFormat (издатель)

Возвращает объект **[ConnectorFormat](connectorformat-object-publisher.md)** , который содержит соединитель свойства форматирования. Применяется к **фигуры** или **ShapeRange** объектов, представляющих соединители. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ConnectorFormat**

 переменная _expression_A, представляющий объект **фигуры** .


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


