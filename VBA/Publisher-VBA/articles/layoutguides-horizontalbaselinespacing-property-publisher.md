---
title: "Свойство LayoutGuides.HorizontalBaseLineSpacing (издатель)"
keywords: vbapb10.chm1114132
f1_keywords: vbapb10.chm1114132
ms.prod: publisher
api_name: Publisher.LayoutGuides.HorizontalBaseLineSpacing
ms.assetid: 19899a25-c1a5-9c81-f022-d842a3d6c7d8
ms.date: 06/08/2017
ms.openlocfilehash: 38212f352c058a643b8d8c4a856ba16297b45702
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="layoutguideshorizontalbaselinespacing-property-publisher"></a>Свойство LayoutGuides.HorizontalBaseLineSpacing (издатель)

Возвращает значение типа **одного** , который представляет расстояние между горизонтальной базового на указанный объект **LayoutGuides** . Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **HorizontalBaseLineSpacing**

 переменная _expression_A, представляет собой объект- **LayoutGuides** .


### <a name="return-value"></a>Возвращаемое значение

Один


## <a name="remarks"></a>Заметки

Установка свойств объекта **страницы** макета руководство должны возвращаться из коллекции **макетом** .


## <a name="example"></a>Пример

В этом примере задается интервал горизонтальной базового объекта макета руководства по 20 для второй главную страницу в активном документе.


```vb
Dim objLayout As LayoutGuides 
Set objLayout = ActiveDocument.MasterPages(2).LayoutGuides 
objLayout.HorizontalBaseLineSpacing = 20 

```


