---
title: "Свойство ConnectorFormat.BeginConnectionSite (издатель)"
keywords: vbapb10.chm3211522
f1_keywords: vbapb10.chm3211522
ms.prod: publisher
api_name: Publisher.ConnectorFormat.BeginConnectionSite
ms.assetid: 24a9246e-270f-7289-971d-8763acfaf02d
ms.date: 06/08/2017
ms.openlocfilehash: 67b9720484b0feaad86543fbc18358ecca65efd1
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="connectorformatbeginconnectionsite-property-publisher"></a>Свойство ConnectorFormat.BeginConnectionSite (издатель)

Возвращает значение типа **Long** , указывающее подключения сайта, к которому подключен начала соединитель. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **BeginConnectionSite**

 переменная _expression_A, представляет собой объект- **ConnectorFormat** .


### <a name="return-value"></a>Возвращаемое значение

Длинный


## <a name="remarks"></a>Заметки

Если в начало указанный соединитель не будет присоединен к фигуры, это свойство приводит к ошибке.

Свойство **[EndConnectionSite](connectorformat-endconnectionsite-property-publisher.md)** используется для возврата сайта, к которому подключен конца соединителя.


## <a name="example"></a>Пример

В этом примере предполагается, что первая страница в активной публикации уже содержит две фигуры, подключенное соединителем с именем Conn1To2. Код добавляет прямоугольник и соединитель для первой страницы. В начало новый соединитель будет присоединена к на одном узле подключения в начале соединитель с именем Conn1To2 и end новый соединитель будет присоединена к точке, один на новый прямоугольник.


```vb
Dim shpNew As Shape 
Dim intSite As Integer 
Dim shpOld As Shape 
 
With ActiveDocument.Pages(1).Shapes 
 
 ' Add new rectangle. 
 Set shpNew = .AddShape(Type:=msoShapeRectangle, _ 
 Left:=450, Top:=190, Width:=200, Height:=100) 
 
 ' Add new connector. 
 .AddConnector(Type:=msoConnectorCurve, _ 
 BeginX:=0, BeginY:=0, EndX:=10, EndY:=10) _ 
 .Name = "Conn1To3" 
 
 ' Get connection site number of old shape, and set 
 ' reference to old shape. 
 With .Item("Conn1To2").ConnectorFormat 
 intSite = .BeginConnectionSite 
 Set shpOld = .BeginConnectedShape 
 End With 
 
 ' Connect new connector to old shape and new rectangle. 
 With .Item("Conn1To3").ConnectorFormat 
 .BeginConnect ConnectedShape:=shpOld, _ 
 ConnectionSite:=intSite 
 .EndConnect ConnectedShape:=shpNew, _ 
 ConnectionSite:=1 
 End With 
End With
```


