/**
 * Copyright 2023 HyogoICT
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *       http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
/**
 * Copyright 2023 Google LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *       http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
/* eslint-disable @typescript-eslint/no-unused-vars */
/**
 * Copyright 2024 T.Yamada
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 * http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

'use strict';
function createForm() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  if (!sheet) {
    throw new Error('Sheet not found');
  }
  const formTitle = sheet.getRange('B1').getDisplayValue();
  const formDescription = sheet.getRange('B2').getDisplayValue();
  const form = FormApp.create(formTitle);
  form.setDescription(formDescription).setIsQuiz(true);
  const firstRow = 5;
  const lastRow = sheet.getLastRow();
  const dataRows = lastRow - (firstRow - 1);
  const data = sheet.getRange(firstRow, 1, dataRows, 11).getDisplayValues();

  type Question = {
    title: string;
    helpText: string;
    point: number;
    type: string;
    answer: number | number[];
    choice: string[];
    feedback: string;
  };

  const questionsList: Question[] = data.map(question => {
    let answer: number | number[];
    if (isNaN(Number(question[4]))) {
      answer = question[4].split(',').map(Number);
    } else {
      answer = Number(question[4]);
    }
    return {
      title: question[0],
      helpText: question[1],
      point: Number(question[2]),
      type: question[3],
      answer: answer,
      choice: [question[5], question[6], question[7], question[8], question[9]],
      feedback: question[10],
    };
  });
  questionsList.forEach(question => {
    let item;
    switch (question.type) {
      case '多肢選択': {
        item = form.addMultipleChoiceItem();
        break;
      }
      case '多肢選択（複数可）': {
        item = form.addCheckboxItem();
        break;
      }
      case '記述（1行）': {
        item = form.addTextItem();
        const feedbackForTextItem = FormApp.createFeedback()
          .setText(question.feedback)
          .build();
        item
          .setTitle(question.title)
          .setHelpText(question.helpText)
          .setPoints(question.point)
          .setGeneralFeedback(feedbackForTextItem);
        return;
      }
      default: {
        throw new Error(`その問題形式には対応していません: ${question.type}`);
      }
    }
    const choiceList: GoogleAppsScript.Forms.Choice[] = [];
    if (Array.isArray(question.choice)) {
      question.choice.forEach((choice, index) => {
        if (choice !== '') {
          let isCorrect = false;
          if (Array.isArray(question.answer)) {
            if (question.answer.every(Number.isFinite)) {
              isCorrect = question.answer.includes(index + 1);
            } else {
              throw new Error(
                `Invalid answer format for CheckboxItem: ${question.answer}`
              );
            }
          } else {
            if (Number.isFinite(question.answer)) {
              isCorrect = question.answer === index + 1;
            } else {
              throw new Error(
                `Invalid answer format for MultipleChoiceItem: ${question.answer}`
              );
            }
          }
          const choiceObj = item.createChoice(String(choice), isCorrect);
          console.log(
            'Answer is ' +
              question.answer +
              ' And this is ' +
              (index + 1) +
              ' so, ' +
              isCorrect
          );
          choiceList.push(choiceObj);
        }
      });
    }
    const feedback = FormApp.createFeedback()
      .setText(question.feedback ? question.feedback : ' ')
      .build();
    item
      .setTitle(question.title)
      .setHelpText(question.helpText)
      .setPoints(question.point)
      .setChoices(choiceList)
      .setFeedbackForCorrect(feedback)
      .setFeedbackForIncorrect(feedback);
  });
  sheet.getRange('B3').setValue(form.getPublishedUrl());
}
