---
title: "Свойство PageSetup.HorizontalGap (издатель)"
keywords: vbapb10.chm6946818
f1_keywords: vbapb10.chm6946818
ms.prod: publisher
api_name: Publisher.PageSetup.HorizontalGap
ms.assetid: e8ee51e0-59b3-8fb6-21f6-87d67a96dd66
ms.date: 06/08/2017
ms.openlocfilehash: 40e304116822dccc71f98184a7ca01b11bdd7eee
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pagesetuphorizontalgap-property-publisher"></a>Свойство PageSetup.HorizontalGap (издатель)

Возвращает значение **типа Variant** , который представляет расстояние между правым краем одну страницу публикации и левого края к следующей странице публикации в той же строке, при печати нескольких страниц на одном листе бумаги. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **HorizontalGap**

 переменная _expression_A, представляет собой объект- **PageSetup** .


### <a name="return-value"></a>Возвращаемое значение

Variant


## <a name="remarks"></a>Заметки

Числовые значения вычисляются как точки; строковые значения можно в любое устройство, поддерживаемый Microsoft Publisher (например, «2,5 дюйма»). Допустимый диапазон допустимых значений — от 0 до различие между ширину листа и ширину страницы.

Это свойство применяется только к публикации, где печати нескольких страниц на одном листе принтера. Использование этого свойства для другой публикации возникает ошибка.


