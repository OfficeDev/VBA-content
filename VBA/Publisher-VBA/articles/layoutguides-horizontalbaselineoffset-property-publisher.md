---
title: "Свойство LayoutGuides.HorizontalBaseLineOffset (издатель)"
keywords: vbapb10.chm1114131
f1_keywords: vbapb10.chm1114131
ms.prod: publisher
api_name: Publisher.LayoutGuides.HorizontalBaseLineOffset
ms.assetid: b80d2114-8132-db13-a50d-ce904dbe5919
ms.date: 06/08/2017
ms.openlocfilehash: 0200e5eb18914e1ddc7beafa5a67cd4039fe6587
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="layoutguideshorizontalbaselineoffset-property-publisher"></a>Свойство LayoutGuides.HorizontalBaseLineOffset (издатель)

Возвращает значение типа **одного** , который представляет смещение горизонтальной базового на указанный объект **LayoutGuides** . Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **HorizontalBaseLineOffset**

 переменная _expression_A, представляет собой объект- **LayoutGuides** .


### <a name="return-value"></a>Возвращаемое значение

Один


## <a name="remarks"></a>Заметки

Установка свойств объекта **страницы** макета руководство должны возвращаться из коллекции **макетом** .


## <a name="example"></a>Пример

В этом примере задает смещение горизонтальной базового объекта руководства макет до 12 для второй главную страницу в активном документе.


```vb
Dim objLayout As LayoutGuides 
Set objLayout = ActiveDocument.MasterPages(2).LayoutGuides 
objLayout.HorizontalBaseLineSpacing = 12 

```


