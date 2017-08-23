---
title: "Свойство ConnectorFormat.EndConnected (издатель)"
keywords: vbapb10.chm3211523
f1_keywords: vbapb10.chm3211523
ms.prod: publisher
api_name: Publisher.ConnectorFormat.EndConnected
ms.assetid: ace997de-5a11-6b52-ac87-e914adb4212d
ms.date: 06/08/2017
ms.openlocfilehash: 5eac4dae4b0b03ff51657abceed4156daf3d2503
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="connectorformatendconnected-property-publisher"></a>Свойство ConnectorFormat.EndConnected (издатель)

Возвращает константу **MsoTriState** , указывающее, подключен ли в конец указанный соединитель фигуры. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **EndConnected**

 переменная _expression_A, представляющий объект **ConnectorFormat** .


### <a name="return-value"></a>Возвращаемое значение

MsoTriState


## <a name="remarks"></a>Заметки

Свойство **[BeginConnected](connectorformat-beginconnected-property-publisher.md)** определяет, подключен ли в начале соединитель фигуры.

Значение свойства **EndConnected** может иметь одно из **MsoTriState** константы объявляются в библиотеке типов Microsoft Office и показаны в следующей таблице.



|**Константы**|**Описание**|
|:-----|:-----|
| **msoFalse**| В конец указанный соединитель не подключен к фигуре.|
| **msoTriStateMixed**|Возвращаемое значение. Указывает сочетание **msoTrue** и **msoFalse** в диапазоне указанные форму.|
| **msoTrue**| В конец указанный соединитель подключена к фигуры.|

## <a name="example"></a>Пример

Если третий фигуры на первой странице в активной публикации соединителя, подключенной к фигуры, end, в этом примере хранит номера сайта подключения, содержит ссылку на подключенных фигуры и отключается конца соединительной линии фигуры.


```vb
Dim intSite As Integer 
Dim shpConnected As Shape 
 
With ActiveDocument.Pages(1).Shapes(3) 
 
 ' Test whether shape is a connector. 
 If .Connector Then 
 With .ConnectorFormat 
 
 ' Test whether connector is connected to another shape. 
 If .End Connected Then 
 
 ' Store connection site number. 
 intSite = .EndConnectionSite 
 
 ' Set reference to connected shape. 
 Set shpConnected = .EndConnectedShape 
 
 ' Disconnect connector and shape. 
 .EndDisconnect 
 End If 
 End With 
 End If 
End With 

```


