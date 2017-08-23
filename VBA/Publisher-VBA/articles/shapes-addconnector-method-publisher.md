---
title: "Метод Shapes.AddConnector (издатель)"
keywords: vbapb10.chm2162705
f1_keywords: vbapb10.chm2162705
ms.prod: publisher
api_name: Publisher.Shapes.AddConnector
ms.assetid: fd1ef969-7960-2555-e355-9804c86f6c01
ms.date: 06/08/2017
ms.openlocfilehash: 66c90b627d54b7eb6f2a3e3e8ddc508317f4cefd
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapesaddconnector-method-publisher"></a>Метод Shapes.AddConnector (издатель)

Добавляет новый объект **[фигуры](shape-object-publisher.md)** , представляющее соединитель для указанной коллекции **[фигур](shapes-object-publisher.md)** .


## <a name="syntax"></a>Синтаксис

 _выражение_. **AddConnector** ( **_Тип_**, **_НачалоX_**, **_BeginY_**, **_EndX_**, **_EndY_**)

 переменная _expression_A, представляет собой объект- **фигур** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Тип|Обязательный| **MsoConnectorType**|Тип соединителя для добавления.|
|BeginX|Обязательное свойство.| **Variant**|X координата начальную точку соединитель.|
|BeginY|Обязательное свойство.| **Variant**|Начальную точку соединитель по оси y.|
|EndX|Обязательное свойство.| **Variant**|Координата x конечной точки соединителя.|
|EndY|Обязательное свойство.| **Variant**|Конечная точка соединителя по оси y.|

### <a name="return-value"></a>Возвращаемое значение

Shape


## <a name="remarks"></a>Заметки

Для BeginX, BeginY, EndX и EndY параметров числовые значения вычисляются в точках; строк может быть в любой устройств, поддерживаемых Microsoft Publisher (например, «2,5 дюйма»).

Новый соединитель не связанные с какой другой фигурой; Используйте методы **[BeginConnect](connectorformat-beginconnect-method-publisher.md)** и **[EndConnect](connectorformat-endconnect-method-publisher.md)** для подключения нового соединителя на другую фигуру.

Параметр типа может иметь одно из следующих констант **MsoConnectorType** .



| **msoConnectorCurve**| Добавляет искривленной формы. | | **msoConnectorElbow**| Добавляет соединителя форме локтя. | | **msoConnectorStraight**| Добавляет линейное соединитель. | | **msoConnectorTypeMixed**| Не используется с помощью этого метода. |

## <a name="example"></a>Пример

В следующем примере добавляется новый соединитель линейное для первой страницы active публикации.


```vb
Dim shpConnect As Shape 
 
Set shpConnect = ActiveDocument.Pages(1).Shapes.AddConnector _ 
 (Type:=msoConnectorStraight, _ 
 BeginX:=144, BeginY:=144, _ 
 EndX:=180, EndY:=72)
```


