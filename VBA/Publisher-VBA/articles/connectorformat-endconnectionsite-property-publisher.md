---
title: "Свойство ConnectorFormat.EndConnectionSite (издатель)"
keywords: vbapb10.chm3211525
f1_keywords: vbapb10.chm3211525
ms.prod: publisher
api_name: Publisher.ConnectorFormat.EndConnectionSite
ms.assetid: 61d38281-7a48-99e1-bda7-67e61b7225a2
ms.date: 06/08/2017
ms.openlocfilehash: 23c1e62af69000afd4b430b7f3364568c4abb86c
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="connectorformatendconnectionsite-property-publisher"></a>Свойство ConnectorFormat.EndConnectionSite (издатель)

Возвращает значение типа **Long** , указывающее подключения сайта, к которому подключен конца соединителя. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **EndConnectionSite**

 переменная _expression_A, представляющий объект **ConnectorFormat** .


### <a name="return-value"></a>Возвращаемое значение

Длинный


## <a name="remarks"></a>Заметки

Если в конец указанный соединитель не будет присоединен к фигуры, это свойство приводит к ошибке.

Свойство **[BeginConnectionSite](connectorformat-beginconnectionsite-property-publisher.md)** используется для возврата сайта, к которому подключен начала соединитель.


## <a name="example"></a>Пример

В этом примере предполагается, что первая страница в активной публикации уже содержит две фигуры, подключенное соединителем с именем Conn1To2. Код добавляет прямоугольник и соединитель для первой страницы. Конец новый соединитель будет присоединено на том же сайте подключения в конец соединителя с именем Conn1To2 и подключения к сайту один на новый прямоугольник будет присоединено начала нового соединителя.


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
 intSite = .EndConnectionSite 
 Set shpOld = .EndConnectedShape 
 End With 
 
 ' Connect new connector to old shape and new rectangle. 
 With .Item("Conn1To3").ConnectorFormat 
 .EndConnect ConnectedShape:=shpOld, _ 
 ConnectionSite:=intSite 
 .BeginConnect ConnectedShape:=shpNew, _ 
 ConnectionSite:=1 
 End With 
End With 

```


