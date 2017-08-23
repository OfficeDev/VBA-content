---
title: "Метод ShapeRange.Item (издатель)"
keywords: vbapb10.chm2293760
f1_keywords: vbapb10.chm2293760
ms.prod: publisher
api_name: Publisher.ShapeRange.Item
ms.assetid: f316bbac-b0be-0281-585b-c32dcb709b66
ms.date: 06/08/2017
ms.openlocfilehash: 4cd46ed1759c2947a0a5d852b03a9dd267e046ae
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shaperangeitem-method-publisher"></a>Метод ShapeRange.Item (издатель)

Возвращает объект отдельных в указанном семействе сайтов.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Элемент** ( **_Индекс_**)

 переменная _expression_A, представляющий объект **ShapeRange** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Индекс|Обязательное свойство.| **Variant**|Номер или имя поля или поля элемента списка, чтобы возвратить.|

### <a name="return-value"></a>Возвращаемое значение

Shape


## <a name="example"></a>Пример

Этот пример возвращает первую фигуру внутри сгруппированных фигуры.


```vb
Dim shpTemp As Shape 
 
Set shpTemp = ActiveDocument.Pages(Index:=1) _ 
 .Shapes(1).GroupItems.Item(Index:=1)
```


