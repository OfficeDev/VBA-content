---
title: "Свойство LayoutGuides.VerticalBaseLineSpacing (издатель)"
keywords: vbapb10.chm1114134
f1_keywords: vbapb10.chm1114134
ms.prod: publisher
api_name: Publisher.LayoutGuides.VerticalBaseLineSpacing
ms.assetid: 49391fbd-86c0-b53f-ff57-009af9341e74
ms.date: 06/08/2017
ms.openlocfilehash: cf72093e97132c3fc0826f97ea413e4f6b485990
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="layoutguidesverticalbaselinespacing-property-publisher"></a>Свойство LayoutGuides.VerticalBaseLineSpacing (издатель)

Возвращает значение типа **одного** , представляющий интервал по вертикали базового на указанный объект **LayoutGuides** . Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **VerticalBaseLineSpacing**

 переменная _expression_A, представляет собой объект- **LayoutGuides** .


### <a name="return-value"></a>Возвращаемое значение

Один


## <a name="remarks"></a>Заметки

Установка свойств объекта **страницы** макета руководство должны возвращаться из коллекции **макетом** .


## <a name="example"></a>Пример

В этом примере задается интервал по вертикали базового объекта **LayoutGuides** 12 для второй главную страницу в активном документе.


```vb
Dim objLayout As LayoutGuides 
Set objLayout = ActiveDocument.MasterPages(2).LayoutGuides 
objLayout.VerticalBaseLineSpacing = 12 

```


