---
title: "Метод Application.Open (издатель)"
keywords: vbapb10.chm131128
f1_keywords: vbapb10.chm131128
ms.prod: publisher
api_name: Publisher.Application.Open
ms.assetid: 560ac406-f058-8fd8-4b6d-978ff369de9b
ms.date: 06/08/2017
ms.openlocfilehash: 592cb1e3f094d12074748f572f550a4ca2cbe4a5
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="applicationopen-method-publisher"></a>Метод Application.Open (издатель)

Возвращает объект **[Document](document-object-publisher.md)** , представляющий открываемые публикации.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Открыть** ( **_Имя файла_**, **_только для чтения_**, **_AddToRecentFiles_**, **_SaveChanges_**)

 переменная _expression_A, представляющий объект **приложения** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Имя файла|Обязательное свойство.| **String**|Имя публикации (допускаются пути).|
|ReadOnly|Необязательный| **Boolean**| **Значение true,** чтобы открыть публикации с доступом только для чтения. Значение по умолчанию — **False**.|
|AddToRecentFiles|Необязательный| **Boolean**| **Значение true** (по умолчанию) для добавления имени файла в список недавно использованных файлов в нижней части меню "файл".|
|SaveChanges|Необязательный| **PbSaveOptions**|Указывает, что следует сделать Microsoft Publisher Если уже открытой публикации с несохраненными изменениями.|
|OpenConflictDocument|Необязательный| **Boolean**| **Значение true,** для открытия публикации локального конфликта автономной конфликта. Значение по умолчанию — **False**.|

### <a name="return-value"></a>Возвращаемое значение

Документ


## <a name="remarks"></a>Заметки

Так как Publisher однодокументного интерфейса, метод **Откройте** работает только в том случае, когда откройте новый экземпляр объекта Publisher. В следующем примере показано, как создать новый видимым экземпляр Publisher. После завершения работы с второй экземпляр, свойство [Visible](window-visible-property-publisher.md)окна приложения можно задать значение **False**, но процесс продолжает работать в фоновом режиме, даже если он не отображается. Чтобы закрыть второй экземпляр, необходимо задать объект равна **Nothing**.

Параметр SaveChanges может иметь одно из **PbSaveOption** константы объявляются в библиотеке типов издателя и показаны в следующей таблице.



|**Константы**|**Описание**|
|:-----|:-----|
| **pbDoNotSaveChanges**|Закройте открыть публикацию без сохранения изменений. |
| **pbPromptToSaveChanges**|Запрашивать у пользователя, следует ли сохранить изменения в открытой публикации. По умолчанию.|
| **pbSaveChanges**|Сохраните открыть публикацию перед закрытием.|

## <a name="example"></a>Пример

В этом примере создается второй экземпляр издателя и открывает указанной публикации с доступом только для чтения. 

Для работы этого примера необходимо заменить _PathToFile_ путь к существующей публикации.




```vb
Sub OpenNewPub() 
 Dim appPub As New Publisher.Application 
 appPub.Open FileName:="PathToFile", _ 
 ReadOnly:=True, AddToRecentFiles:=False, _ 
 SaveChanges:=pbPromptToSaveChanges 
 appPub.ActiveWindow.Visible = True 
End Sub
```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект приложения](application-object-publisher.md)

