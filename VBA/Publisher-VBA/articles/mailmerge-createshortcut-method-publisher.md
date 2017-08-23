---
title: "Метод MailMerge.CreateShortcut (издатель)"
keywords: vbapb10.chm6225942
f1_keywords: vbapb10.chm6225942
ms.prod: publisher
api_name: Publisher.MailMerge.CreateShortcut
ms.assetid: 96878925-41ce-4873-931e-d5c05307a94a
ms.date: 06/08/2017
ms.openlocfilehash: 5063e925eaf57caf16648954a5078f34e44e10ea
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="mailmergecreateshortcut-method-publisher"></a>Метод MailMerge.CreateShortcut (издатель)

Создается ярлык для файла, содержащего список получателей или продуктов для публикации слияния почты.


## <a name="syntax"></a>Синтаксис

 _выражение_. **CreateShortcut** ( **_Имя файла_**)

 переменная _expression_A, представляет собой объект- **слияния** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Имя файла|Обязательное свойство.| **String**|Имя файла списка список рассылки или продукт, для которого должен быть создан ярлык на панели.|

## <a name="remarks"></a>Заметки

Метод **CreateShortcut** соответствует команда **Сохранить ярлык в список получателей** в области задач **слияния почты** и **Слияния почты** и команды **Сохранить ярлык для списка продуктов** в области задач, **Объединение в каталог** , в интерфейсе пользователя Microsoft Publisher.

Список рассылки получателя файлы с расширением .ols (для ярлыка списка Microsoft Office).


## <a name="example"></a>Пример

Следующие Microsoft Visual Basic для приложений (VBA) макроса показано, как использовать метод **CreateShortcut** для создания ярлыка в список получателей слияния почты. Прежде чем запустить этот макрос, убедитесь, что активный документ подключен к источнику данных. Если активный документ не подключен к источнику данных, можно использовать ** [MailMerge.OpenDataSource](mailmerge-opendatasource-method-publisher.md)** метод для подключения.

Кроме того перед выполнением кода замените _имя пользователя_ в путь к папке на сохраненный файл с именем допустимого пользователя на вашем компьютере или замените путь к папке и имя файла, в которое включен путь и имя файла.

Обратите внимание, что путь к папке, в этом примере типичные для путей к папкам в Microsoft Windows Vista. Необходимо иметь разрешение на сохранение файлов в папке.




```vb
Public Sub CreateShortcut_Example() 
 
 Dim pubMailMerge As Publisher.MailMerge 
 Set pubMailMerge = ThisDocument.MailMerge 
 
 pubMailMerge.CreateShortcut ("C:\Users\username\Documents\My Data Sources\MyRecipientList") 
 
End Sub
```


