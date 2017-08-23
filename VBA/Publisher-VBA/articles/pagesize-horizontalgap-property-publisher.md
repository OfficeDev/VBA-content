---
title: "Свойство PageSize.HorizontalGap (издатель)"
keywords: vbapb10.chm8847368
f1_keywords: vbapb10.chm8847368
ms.prod: publisher
api_name: Publisher.PageSize.HorizontalGap
ms.assetid: 14c14534-c1c7-db2d-c7bf-8b7fd66c245e
ms.date: 06/08/2017
ms.openlocfilehash: d9285348503dc22c910c2081a592c57d8c730d5d
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pagesizehorizontalgap-property-publisher"></a>Свойство PageSize.HorizontalGap (издатель)

Возвращает значение **типа Variant** , представляющее расстояние между правым краем одну страницу публикации и левого края к следующей странице публикации в той же строке по размеру пустая страница, представленного объектом **PageSize** родительского при печати нескольких страниц на одном листе бумаги. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **HorizontalGap**

 переменная _expression_A, представляет собой объект- **PageSize** .


### <a name="return-value"></a>Возвращаемое значение

Variant


## <a name="remarks"></a>Заметки

Размер пустая страница, представленного объектом **PageSize** родительского соответствует одному значков, отображаемых в разделе **Пустая страница размеры** в диалоговом окне **Параметры страницы** в интерфейсе пользователя Microsoft Publisher.

Числовые значения вычисляются как точки; строковые значения можно в любое устройство, поддерживаемый Microsoft Publisher (например, «2,5 дюйма»). Допустимый диапазон допустимых значений — от 0 до различие между ширину листа и ширину страницы.


