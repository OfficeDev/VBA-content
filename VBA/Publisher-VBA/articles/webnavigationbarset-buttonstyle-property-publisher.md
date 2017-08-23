---
title: "Свойство WebNavigationBarSet.ButtonStyle (издатель)"
keywords: vbapb10.chm8519685
f1_keywords: vbapb10.chm8519685
ms.prod: publisher
api_name: Publisher.WebNavigationBarSet.ButtonStyle
ms.assetid: 39251032-d51e-3895-af18-cb4b613a38f4
ms.date: 06/08/2017
ms.openlocfilehash: 1140eacd99c5695abf0f5a3f5caf168f34504aa9
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="webnavigationbarsetbuttonstyle-property-publisher"></a>Свойство WebNavigationBarSet.ButtonStyle (издатель)

Задает или возвращает константу **PbWizardNavBarButtonStyle** , представляющий стиль кнопок панели навигации: большой, небольшой или только текст. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ButtonStyle**

 переменная _expression_A, представляет собой объект- **WebNavigationBarSet** .


### <a name="return-value"></a>Возвращаемое значение

PbWizardNavBarButtonStyle


## <a name="remarks"></a>Заметки

Значение свойства **ButtonStyle** может иметь одно из **[PbWizardNavBarButtonStyle](pbwizardnavbarbuttonstyle-enumeration-publisher.md)** константы в библиотеке типов, Microsoft Publisher.


## <a name="example"></a>Пример

В следующем примере задается стиль кнопки для **pbnbButtonStyleLarge** для первого набора панель навигации Web активного документа.


```vb
ActiveDocument.WebNavigationBarSets(1).ButtonStyle = pbnbButtonStyleLarge
```


