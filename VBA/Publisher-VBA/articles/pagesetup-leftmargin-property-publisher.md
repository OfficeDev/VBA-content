---
title: "Свойство PageSetup.LeftMargin (издатель)"
keywords: vbapb10.chm6946819
f1_keywords: vbapb10.chm6946819
ms.prod: publisher
api_name: Publisher.PageSetup.LeftMargin
ms.assetid: 19fbb72e-bb6e-18e9-28f3-c7e99b071bfb
ms.date: 06/08/2017
ms.openlocfilehash: 609438dba4832760eca79dad954cb15d395a0b22
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pagesetupleftmargin-property-publisher"></a>Свойство PageSetup.LeftMargin (издатель)

Возвращает **Variant** , представляющий расстояние (в точках) между левой границей страницы принтера и левого края страницы публикации при печати нескольких страниц на одном листе. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **LeftMargin**

 переменная _expression_A, представляет собой объект- **PageSetup** .


### <a name="return-value"></a>Возвращаемое значение

Variant


## <a name="remarks"></a>Заметки

Числовые значения вычисляются как точки. Строковые значения можно в любое устройство, поддерживаемый Microsoft Publisher (например, 2,5 дюйма). Допустимый диапазон допустимых значений — от 0 до различие между ширину листа и ширину страницы.

Свойство **LeftMargin** возвращает значение только при печати нескольких страниц на одном листе бумаги. Если предпринимается попытка использовать свойство **LeftMargin** в других ситуациях, Microsoft Publisher возвращает **значение Nothing**.


