---
title: "Свойство WebOptions.RelyOnVML (издатель)"
keywords: vbapb10.chm8257543
f1_keywords: vbapb10.chm8257543
ms.prod: publisher
api_name: Publisher.WebOptions.RelyOnVML
ms.assetid: 8cd29d64-48a6-d33e-cb9d-6b1ea094069a
ms.date: 06/08/2017
ms.openlocfilehash: 21c56e4c37148003ec5e973797920443a44375f5
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="weboptionsrelyonvml-property-publisher"></a>Свойство WebOptions.RelyOnVML (издатель)

Возвращает или задает **логическое** значение, указывающее, создаются ли файлы изображений из графические объекты при сохранении веб-публикации. Если **значение True**, файлы еще не создан. Если **значение False**, изображения, созданные. Значение по умолчанию — **False**. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **RelyOnVML**

 переменная _expression_A, представляет собой объект- **WebOptions** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="remarks"></a>Заметки

Можно уменьшить размеры файлов, не создавая изображения для графические объекты. Обратите внимание на то, что веб-браузер должен поддерживать VML Vector Markup Language () для отображения графических объектов. Microsoft Internet Explorer версии 5.0 и более поздних версий поддерживают VML, поэтому свойство **RelyOnVML** удалось установить значение **True** , если для этих браузеров. Для браузеров, не поддерживающих VML графический объект не появится при сохранении веб-публикации с помощью этого свойства включено.

Если не уверены какие браузеры будет использоваться для просмотра веб-сайта, это свойство должно быть присвоено **значение False**.


## <a name="example"></a>Пример

В следующем примере предполагается, что конечные пользователи Microsoft Internet Explorer версии 5.0 и поэтому указывает, что не следует создавать изображений из графические объекты при сохранении веб-публикации.


```vb
Dim theWO As WebOptions 
 
Set theWO = Application.WebOptions 
 
With theWO 
 .RelyOnVML = True 
End With
```


