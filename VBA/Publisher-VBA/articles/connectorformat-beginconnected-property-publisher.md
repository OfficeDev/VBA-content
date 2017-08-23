---
title: "Свойство ConnectorFormat.BeginConnected (издатель)"
keywords: vbapb10.chm3211520
f1_keywords: vbapb10.chm3211520
ms.prod: publisher
api_name: Publisher.ConnectorFormat.BeginConnected
ms.assetid: ed70561e-b63e-530d-87be-1e6b7d87c425
ms.date: 06/08/2017
ms.openlocfilehash: 28438ef68ad560e5845c6a658dc1d35988768950
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="connectorformatbeginconnected-property-publisher"></a>Свойство ConnectorFormat.BeginConnected (издатель)

Возвращает константу **MsoTriState**, указывающее, подключен ли в начало соединительной фигуры. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **BeginConnected**

 переменная _expression_A, представляет собой объект- **ConnectorFormat** .


### <a name="return-value"></a>Возвращаемое значение

MsoTriState


## <a name="remarks"></a>Заметки

Значение свойства **BeginConnected** может иметь одно из ** [MsoTriState](http://msdn.microsoft.com/library/2036cfc9-be7d-e05c-bec7-af05e3c3c515%28Office.15%29.aspx)** объявленные константы в библиотеке типов, Microsoft Office.

Свойство **[EndConnected](connectorformat-endconnected-property-publisher.md)** определяет, подключен ли конца соединителя фигуры.


## <a name="example"></a>Пример

Если третий фигуры на первой странице в активной публикации соединитель, начало подключена к фигуры, в этом примере хранит номера сайта подключения, содержит ссылку на подключенных фигуры и отключается фигуру начала соединитель.


```vb
Dim intSite As Integer 
Dim shpConnected As Shape 
 
With ActiveDocument.Pages(1).Shapes(3) 
 
 ' Test whether shape is a connector. 
 If .Connector Then 
 With .ConnectorFormat 
 
 ' Test whether connector is connected to another shape. 
 If .BeginConnected Then 
 
 ' Store connection site number. 
 intSite = .BeginConnectionSite 
 
 ' Set reference to connected shape. 
 Set shpConnected = .BeginConnectedShape 
 
 ' Disconnect connector and shape. 
 .BeginDisconnect 
 End If 
 End With 
 End If 
End With 

```


