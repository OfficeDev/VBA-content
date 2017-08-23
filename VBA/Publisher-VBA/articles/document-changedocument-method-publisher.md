---
title: "Метод Document.ChangeDocument (издатель)"
keywords: vbapb10.chm196756
f1_keywords: vbapb10.chm196756
ms.prod: publisher
api_name: Publisher.Document.ChangeDocument
ms.assetid: c6defa92-99fb-973b-6bb2-e3c2a1b0a4f3
ms.date: 06/08/2017
ms.openlocfilehash: a94393d0ee2bfb2f84689dca6c0147c73464f597
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="documentchangedocument-method-publisher"></a>Метод Document.ChangeDocument (издатель)

Изменяет текущий публикации, с помощью мастера, а также разработки, указанной вами.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ChangeDocument** ( **_Мастер_** **_разработки_**)

 переменная _expression_A, представляющий объект **Document** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Мастер|Обязательное свойство.| **PbWizard**|Тип мастера. Возможные значения см.|
|Разработка|Необязательный| **Длинный**|Тип проекта.|

## <a name="remarks"></a>Заметки

Возможные значения для параметра мастера объявляются в перечислении **[PbWizard](pbwizard-enumeration-publisher.md)** в библиотеке типов, Publisher.


## <a name="example"></a>Пример

Следующие Microsoft Visual Basic для приложений (VBA) макроса показано, как использовать метод **ChangeDocument** для изменения мастера, используемые текущей publicaton для брошюра.


```vb
Public Sub ChangeDocument_Example() 
 
 ThisDocument.ChangeDocument pbWizardBrochures 
 
End Sub
```


