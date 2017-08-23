---
title: "Метод GroupShapes.Item (издатель)"
keywords: vbapb10.chm3342336
f1_keywords: vbapb10.chm3342336
ms.prod: publisher
api_name: Publisher.GroupShapes.Item
ms.assetid: d0e2f8a6-6529-a274-410b-744c2bb55774
ms.date: 06/08/2017
ms.openlocfilehash: 85fed34475658711812cf98ff592182aff90e975
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="groupshapesitem-method-publisher"></a>Метод GroupShapes.Item (издатель)

Возвращает объект отдельных в указанном семействе сайтов.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Элемент** ( **_Индекс_**)

 переменная _expression_A, представляет собой объект- **GroupShapes** .


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


