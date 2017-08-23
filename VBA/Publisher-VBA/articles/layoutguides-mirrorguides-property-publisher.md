---
title: "Свойство LayoutGuides.MirrorGuides (издатель)"
keywords: vbapb10.chm1114119
f1_keywords: vbapb10.chm1114119
ms.prod: publisher
api_name: Publisher.LayoutGuides.MirrorGuides
ms.assetid: 8e6ff709-21e0-2286-5d75-c7ebea05fd26
ms.date: 06/08/2017
ms.openlocfilehash: a81d16ea465f5b38b31e93f53e445e320b214fb7
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="layoutguidesmirrorguides-property-publisher"></a>Свойство LayoutGuides.MirrorGuides (издатель)

Возвращает или задает значение **Boolean** , указывающее, создается ли Microsoft Publisher положения руководство зеркала для публикации сгиб книги. **Значение true,** Если Publisher создает зеркала руководство по позиции для отдельных страниц влево и вправо в публикации сгиб книги; **Значение false,** Если же направляющие полей, строк и столбцов, применяются на всех страницах публикации. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **MirrorGuides**

 переменная _expression_A, представляет собой объект- **LayoutGuides** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="remarks"></a>Заметки

Если свойство **MirrorGuides** имеет **значение True**, параметры полей применить к страницам вправо и дублируются для страниц влево. Кроме того, если задано значение **True**, свойство **MirrorGuides** устанавливает публикации для использования двумя главными страницами вместо одной по умолчанию. — Это первый главной страницы для всех страниц влево и второй — для всех страниц вправо в публикации. Для получения дополнительных сведений см **[макетом](masterpages-object-publisher.md)** объекта.


## <a name="example"></a>Пример

В следующем примере задается Publisher для создания зеркальной руководства для публикации сгиб книги и задает внутри и вне поля каждого двух страницах. В примере задается слева и значения правого поля страницы вправо и Publisher отображают эти значения для страниц влево.


```vb
With ActiveDocument.LayoutGuides 
 .MirrorGuides = True 
 .MarginLeft = 48 
 .MarginRight = 96 
End With
```


