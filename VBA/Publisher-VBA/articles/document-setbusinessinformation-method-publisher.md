---
title: "Метод Document.SetBusinessInformation (издатель)"
keywords: vbapb10.chm196757
f1_keywords: vbapb10.chm196757
ms.prod: publisher
api_name: Publisher.Document.SetBusinessInformation
ms.assetid: 8549f75f-2fb6-6ac6-ecaf-54a0a9b22dc7
ms.date: 06/08/2017
ms.openlocfilehash: 6813be18459f6dcf371d4ad7b3ca4a7c895fec31
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="documentsetbusinessinformation-method-publisher"></a>Метод Document.SetBusinessInformation (издатель)

Применяет набор указанного деловых данных, состоящий из логотип изображения и бизнес-сведений о контакте (например, название компании и адрес), к текущей публикации.


## <a name="syntax"></a>Синтаксис

 _выражение_. **SetBusinessInformation** ( **_Имя_**)

 переменная _expression_A, представляющий объект **Document** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Имя|Обязательное свойство.| **String**|Имя бизнес-информация, применяемый набор.|

## <a name="remarks"></a>Заметки

Вызов метода **SetBusinessInformation** соответствует выбору бизнес-данных (в списке **выберите набор бизнес-информация** ) и затем кнопку **Обновление публикации** в диалоговое окно " **Бизнес-информация** " (меню " **Правка** ") в Microsoft Publisher пользовательского интерфейса (UI). Необходимо создать и изменить набора деловых данных в этом диалоговом окне перед их применения программным путем с помощью метода **SetBusinessInformation** .


## <a name="example"></a>Пример

Следующие Microsoft Visual Basic для приложений (VBA) макроса показано, как использовать метод **SetBusinessInformation** для применения конкретных бизнес-информация, задайте значение текущей публикации. Прежде чем запускать этот код, замените _BISetName_ имя набор бизнес-информации, который ранее был создан в пользовательском Интерфейсе Publisher.


```vb
Public Sub SetBusinessInformation_Example() 
 
 ThisDocument.SetBusinessInformation "BISetName" 
 
End Sub
```


