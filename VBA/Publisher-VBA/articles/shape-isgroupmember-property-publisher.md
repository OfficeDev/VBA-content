---
title: "Свойство Shape.IsGroupMember (издатель)"
keywords: vbapb10.chm2228337
f1_keywords: vbapb10.chm2228337
ms.prod: publisher
api_name: Publisher.Shape.IsGroupMember
ms.assetid: bbd9b662-b47d-d5cf-6858-e208c44f88a0
ms.date: 06/08/2017
ms.openlocfilehash: 5018e4c75eacbbb38bf7db5bb5a3ffafbebaed3a
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapeisgroupmember-property-publisher"></a>Свойство Shape.IsGroupMember (издатель)

Возвращает **значение True** , если указанный фигуры должна быть членом группы, **значение False** в противном случае. Только для чтения **типа Boolean**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **IsGroupMember**

 переменная _expression_A, представляющий объект **фигуры** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="remarks"></a>Заметки

Объект, возвращенный свойством **ParentGroupShape** можно использовать для определения родительскую фигуру для группы.


## <a name="example"></a>Пример

Возвращает значение **True** , если первую фигуру active публикации, является участником группы можно использовать следующий оператор.


```
blnGrouped = Application.ActiveDocument.MasterPages _ 
 .Item.Shapes(1).IsGroupMember
```


