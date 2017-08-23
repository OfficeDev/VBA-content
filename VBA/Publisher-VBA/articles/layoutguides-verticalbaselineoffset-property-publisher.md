---
title: "Свойство LayoutGuides.VerticalBaseLineOffset (издатель)"
keywords: vbapb10.chm1114133
f1_keywords: vbapb10.chm1114133
ms.prod: publisher
api_name: Publisher.LayoutGuides.VerticalBaseLineOffset
ms.assetid: 9a2f031c-4469-ca26-3e79-dfa556762e05
ms.date: 06/08/2017
ms.openlocfilehash: c7e4a83723bf751b79068aedb259b328a2d12b6d
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="layoutguidesverticalbaselineoffset-property-publisher"></a>Свойство LayoutGuides.VerticalBaseLineOffset (издатель)

Возвращает значение типа **одного** , который представляет смещение вертикальной базового на указанный объект **LayoutGuides** . Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **VerticalBaseLineOffset**

 переменная _expression_A, представляет собой объект- **LayoutGuides** .


### <a name="return-value"></a>Возвращаемое значение

Один


## <a name="remarks"></a>Заметки

Установка свойств объекта **страницы** макета руководство должны возвращаться из коллекции **макетом** .


## <a name="example"></a>Пример

В этом примере задает смещение вертикальной базового объекта руководства макет до 12 для второй главную страницу в активном документе.


```vb
Dim objLayout As LayoutGuides 
Set objLayout = ActiveDocument.MasterPages(2).LayoutGuides 
objLayout.VerticalBaseLineOffset = 12 

```


