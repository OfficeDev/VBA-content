---
title: "Свойство PageSize.VerticalGap (издатель)"
keywords: vbapb10.chm8847369
f1_keywords: vbapb10.chm8847369
ms.prod: publisher
api_name: Publisher.PageSize.VerticalGap
ms.assetid: cc6e66ff-9a74-d88f-cfde-2f5bee66432f
ms.date: 06/08/2017
ms.openlocfilehash: 3d8555231e12b1e5d05979a3d9bbf1b76af2316f
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pagesizeverticalgap-property-publisher"></a>Свойство PageSize.VerticalGap (издатель)

Возвращает значение **типа Variant** , представляющее расстояние в точках между нижний край одной страницы публикации и верхнего края страницы публикации сразу под размер пустая страница, представленного объектом **PageSize** родительского. Это свойство применяется только при печати нескольких страниц на одном листе бумаги. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **VerticalGap**

 переменная _expression_A, представляет собой объект- **PageSize** .


### <a name="return-value"></a>Возвращаемое значение

Variant


## <a name="remarks"></a>Заметки

Размер пустая страница, представленного объектом **PageSize** родительского соответствует одному значков, отображаемых в разделе **Пустая страница размеры** в диалоговом окне **Параметры страницы** в интерфейсе пользователя Microsoft Publisher.

Числовые значения вычисляются как точки. Строковые значения можно в любое устройство, поддерживаемый Microsoft Publisher (например, «2,5 дюйма»). Допустимый диапазон допустимых значений — от 0 до различие между высота листа и высота страницы.


