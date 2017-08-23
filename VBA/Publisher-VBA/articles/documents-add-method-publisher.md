---
title: "Метод Documents.Add (издатель)"
keywords: vbapb10.chm8650756
f1_keywords: vbapb10.chm8650756
ms.prod: publisher
api_name: Publisher.Documents.Add
ms.assetid: 1e3536c8-8fc0-8c95-3a4c-b16fe8a99098
ms.date: 06/08/2017
ms.openlocfilehash: d15a565d8437edc7b0c0f118c56ba6c92bc7229b
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="documentsadd-method-publisher"></a>Метод Documents.Add (издатель)

Добавляет новый объект **Document** , представляющий новую публикацию в коллекцию **документов** .


## <a name="syntax"></a>Синтаксис

 _выражение_. **Добавление** ( **_PbWizard_**, **_desid_**)

 _expression_An выражение, возвращающее объект **документы** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|PbWizard|Необязательный| **PbWizard**|Мастер для создания новой публикации.|
|desid|Необязательный| **Длинный**|КОД структуры, чтобы применить новые публикации.|

### <a name="return-value"></a>Возвращаемое значение

Документ


## <a name="remarks"></a>Заметки

Значение параметра PbWizard должно быть константа из перечисления **[PbWizard](pbwizard-enumeration-publisher.md)** , указанному в библиотеке типов Microsoft Publisher 2007.

Значение параметра desid должно быть идентификатор структуры для применения. Можно определить идентификатор разработки, создания новой публикации с использованием мастера и макет в пользовательском интерфейсе издателя и выполнения следующих Visual Basic для приложений (VBA).




```vb
Public Sub FindDesignID() 
 
 Dim pbWizard As Wizard 
 Dim pbWizardProperty As WizardProperty 
 
 Set pbWizard = ThisDocument.Wizard 
 Set pbWizardProperty = pbWizard.Properties(1) 
 
 Debug.Print pbWizardProperty.CurrentValueId 
 
End Sub
```


