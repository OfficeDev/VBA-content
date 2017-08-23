---
title: "Свойство PageSetup.TopMargin (издатель)"
keywords: vbapb10.chm6946837
f1_keywords: vbapb10.chm6946837
ms.prod: publisher
api_name: Publisher.PageSetup.TopMargin
ms.assetid: 4eee9b1e-6c76-7831-13bc-25926c3c0f10
ms.date: 06/08/2017
ms.openlocfilehash: 53a15e8b828bd80927fe7c045a2271ae7ffebd2b
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pagesetuptopmargin-property-publisher"></a>Свойство PageSetup.TopMargin (издатель)

Возвращает значение **типа Variant** , представляющее расстояние между верхней границы листа принтера и верхнего края страницы публикации. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **TopMargin**

 переменная _expression_A, представляет собой объект- **PageSetup** .


### <a name="return-value"></a>Возвращаемое значение

Variant


## <a name="remarks"></a>Заметки

Числовые значения вычисляются как точки. Строковые значения можно в любое устройство, поддерживаемый Microsoft Publisher (например, «2,5 дюйма»). Допустимый диапазон допустимых значений — от 0 до различие между высота листа и Высота страниц публикации.

Свойство **TopMargin** возвращает значение только при печати нескольких страниц на одном листе бумаги. При попытке использовать в других ситуациях, Microsoft Publisher возвращает значение nothing.


