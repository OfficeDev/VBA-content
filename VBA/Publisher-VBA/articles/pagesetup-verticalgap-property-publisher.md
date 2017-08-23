---
title: "Свойство PageSetup.VerticalGap (издатель)"
keywords: vbapb10.chm6946838
f1_keywords: vbapb10.chm6946838
ms.prod: publisher
api_name: Publisher.PageSetup.VerticalGap
ms.assetid: 191d66c4-d168-625a-47b7-028167a98af9
ms.date: 06/08/2017
ms.openlocfilehash: 15ae19ee6e99edaa7aae33aa6c1c046f9c7d3ffb
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pagesetupverticalgap-property-publisher"></a>Свойство PageSetup.VerticalGap (издатель)

Возвращает **Variant** , представляющий расстояние (в точках) между нижний край одной страницы публикации и верхнего края страницы публикации под ней при печати нескольких страниц публикации на странице одного принтера. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **VerticalGap**

 переменная _expression_A, представляет собой объект- **PageSetup** .


### <a name="return-value"></a>Возвращаемое значение

Variant


## <a name="remarks"></a>Заметки

Можно использовать свойство **VerticalGap** , если необходимо использовать печати нескольких страниц на одном листе бумаги. Если размер страницы, включая значения для свойства **VerticalGap** и **HorizontalGap** больше половины размер бумаги, Microsoft Publisher выводится сообщение об ошибке.


