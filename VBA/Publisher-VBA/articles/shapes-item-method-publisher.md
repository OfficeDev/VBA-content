---
title: "Метод Shapes.Item (издатель)"
keywords: vbapb10.chm2162688
f1_keywords: vbapb10.chm2162688
ms.prod: publisher
api_name: Publisher.Shapes.Item
ms.assetid: 174bbabb-e19f-4638-6dd8-780a8617fd70
ms.date: 06/08/2017
ms.openlocfilehash: e58ca50c4fc0bb800fa44ef18f79742c04c81607
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapesitem-method-publisher"></a>Метод Shapes.Item (издатель)

Возвращает объект отдельных в указанном семействе сайтов.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Элемент** ( **_Индекс_**)

 переменная _expression_A, представляет собой объект- **фигур** .


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


