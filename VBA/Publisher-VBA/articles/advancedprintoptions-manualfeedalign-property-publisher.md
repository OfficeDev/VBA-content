---
title: "Свойство AdvancedPrintOptions.ManualFeedAlign (издатель)"
keywords: vbapb10.chm7077928
f1_keywords: vbapb10.chm7077928
ms.prod: publisher
api_name: Publisher.AdvancedPrintOptions.ManualFeedAlign
ms.assetid: 5c2dc0a7-981f-731d-6a85-0971c7e19a62
ms.date: 06/08/2017
ms.openlocfilehash: b71688f89cc776754f09e45b76e0e0119d3b87bb
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="advancedprintoptionsmanualfeedalign-property-publisher"></a>Свойство AdvancedPrintOptions.ManualFeedAlign (издатель)

Получает или задает выравнивание (левому или правому, или центр) где подача конвертов принтера вручную веб-канала. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ManualFeedAlign**

 переменная _expression_A, представляющий объект **AdvancedPrintOptions** .


### <a name="return-value"></a>Возвращаемое значение

 **PbPlacementType**


## <a name="remarks"></a>Заметки

Значение свойства **ManualFeedAlign** в сочетании с ** [AdvancedPrintOptions.ManualFeedDirection](advancedprintoptions-manualfeeddirection-property-publisher.md)** параметр свойство соответствует параметру **способ подачи конвертов** в диалоговом окне **Настройка конверта** в интерфейсе пользователя Microsoft Publisher. (В меню **файл** выберите пункт **Настройка печати**. На вкладке **Сведения о** нажмите кнопку **Дополнительные параметры принтера**. На вкладке **Мастера установки принтера** выберите **Окно настройки конвертов**).

Возможные значения для **ManualFeedAlign** : **pbPlacementCenter** (3), **pbPlacementLeft** (1) и **pbPlacementRight** (2).


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект AdvancedPrintOptions](advancedprintoptions-object-publisher.md)

