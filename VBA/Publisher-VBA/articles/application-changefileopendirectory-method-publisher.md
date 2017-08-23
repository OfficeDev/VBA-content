---
title: "Метод Application.ChangeFileOpenDirectory (издатель)"
keywords: vbapb10.chm131124
f1_keywords: vbapb10.chm131124
ms.prod: publisher
api_name: Publisher.Application.ChangeFileOpenDirectory
ms.assetid: 9178881c-2f7f-9063-31d1-14d4745f0666
ms.date: 06/08/2017
ms.openlocfilehash: 7d25e9056f0bd237fa8706de484d9b183991a6ad
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="applicationchangefileopendirectory-method-publisher"></a>Метод Application.ChangeFileOpenDirectory (издатель)

Задает папку, в которой Microsoft Publisher выполняется поиск документов. Содержимое указанной папки, перечислены в следующий раз отображается диалоговое окно " **Открыть публикацию** " (меню " **файл** ").


## <a name="syntax"></a>Синтаксис

 _выражение_. **ChangeFileOpenDirectory** ( **_Dir_**)

 переменная _expression_A, представляющий объект **приложения** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Dir|Обязательное свойство.| **String**|Путь к каталогу.|

## <a name="remarks"></a>Заметки

Publisher выполняет поиск указанной папки для документов, пока пользователь изменяет папку в диалоговом окне **Открыть публикацию** или текущей завершается сеанс Publisher. Используйте свойство **[PathForPublications](options-pathforpublications-property-publisher.md)** объекта **Параметры** для изменения папки по умолчанию для документов в каждом сеансе Publisher.


## <a name="example"></a>Пример

В этом примере изменяется папки, в которой Publisher выполняется поиск документов. (Обратите внимание на то, что действительный путь к файлу для работы этого примера необходимо заменить PathToDirectory.)


```vb
Sub ChangeOpenPath() 
 ChangeFileOpenDirectory Dir:="PathToDirectory" 
End Sub
```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект приложения](application-object-publisher.md)

