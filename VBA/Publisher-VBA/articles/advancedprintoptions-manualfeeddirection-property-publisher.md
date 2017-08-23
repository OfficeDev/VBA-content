---
title: "Свойство AdvancedPrintOptions.ManualFeedDirection (издатель)"
keywords: vbapb10.chm7077929
f1_keywords: vbapb10.chm7077929
ms.prod: publisher
api_name: Publisher.AdvancedPrintOptions.ManualFeedDirection
ms.assetid: 6c241594-d113-c3bd-5669-d3046e824c4e
ms.date: 06/08/2017
ms.openlocfilehash: 59a0abccea9c70b666aa567d0f2949ab698777b9
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="advancedprintoptionsmanualfeeddirection-property-publisher"></a>Свойство AdvancedPrintOptions.ManualFeedDirection (издатель)

Получает или задает ориентации (книжная или альбомная) как подача конвертов принтера вручную веб-канала. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ManualFeedDirection**

 переменная _expression_A, представляющий объект **AdvancedPrintOptions** .


### <a name="return-value"></a>Возвращаемое значение

PbOrientationType


## <a name="remarks"></a>Заметки

Значение свойства **ManualFeedDirection** в сочетании с ** [AdvancedPrintOptions.ManualFeedAlign](advancedprintoptions-manualfeedalign-property-publisher.md)** параметр свойство соответствует параметру **способ подачи конвертов** в диалоговом окне **Настройка конверта** в интерфейсе пользователя Microsoft Publisher. (В меню **файл** выберите пункт **Настройка печати**. На вкладке **Сведения о** нажмите кнопку **Дополнительные параметры принтера**. На вкладке **Мастера установки принтера** выберите **Окно настройки конвертов**)

Возможные значения для **ManualFeedDirection** : **pbOrientationLandscape** (2) и **pbOrientationPortrait** (1).


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект AdvancedPrintOptions](advancedprintoptions-object-publisher.md)

