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
  const data = sheet.getRange(firstRow, 1, dataRows, 12).getDisplayValues();

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
    // Concatinate A and B columns to create a title
    const titleA = question[0];
    const titleB = question[1];
    const title = `${titleA} ${titleB}`;
    let answer: number | number[];
    if (isNaN(Number(question[5]))) {
      answer = question[5].split(',').map(Number);
    } else {
      answer = Number(question[5]);
    }
    return {
      title: title,
      helpText: question[2],
      point: Number(question[3]),
      type: question[4],
      answer: answer,
      choice: [
        question[6],
        question[7],
        question[8],
        question[9],
        question[10],
      ],
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

function addQuestionNumToSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  if (!sheet) {
    throw new Error('Sheet not found');
  }
  const firstRow = 5;
  const lastRow = sheet.getLastRow();
  const titles = sheet.getRange('B' + firstRow + ':B' + lastRow).getValues();

  // A列に問題番号を設定
  for (let i = 0; i < titles.length; i++) {
    if (titles[i][0]) {
      // タイトルが空でない場合にのみ処理
      const questionNumber = `問${i + 1}`;
      sheet.getRange(firstRow + i, 1).setValue(questionNumber);
    }
  }
}

function createTestSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getActiveSheet();
  // A列とB列から問題番号と問題文を取得
  const questionsRange = sourceSheet.getRange(
    'A4:B' + sourceSheet.getLastRow()
  );
  const questionsValues = questionsRange.getValues();
  // G列からK列までの選択肢を取得
  const choicesRange = sourceSheet.getRange('G4:K' + sourceSheet.getLastRow());
  const choicesValues = choicesRange.getValues();

  const questionsList = questionsValues.map((question, index) => ({
    number: question[0], // 問題番号
    question: question[1], // 問題文
    choices: choicesValues[index], // 選択肢
  }));

  const middleIndex = Math.ceil(questionsList.length / 2);
  const totalColumns = sourceSheet.getMaxColumns();
  const offset = Math.floor(totalColumns / 2); // 中央の列を基準にオフセットを計算

  // 新しいシートを作成し、B1の値をシート名として設定
  const sheetName = sourceSheet.getRange('B1').getValue() || 'New Test Sheet';
  const newSheet = ss.insertSheet(sheetName);

  // 左側のセクションにデータを配置
  questionsList.slice(0, middleIndex).forEach((item, index) => {
    const row = index + 4; // B4から開始
    newSheet.getRange(row, 1).setValue(item.number); // A列: 問題番号
    newSheet.getRange(row, 2).setValue(item.question); // B列: 問題文
    item.choices.forEach((choice, cIndex) => {
      newSheet.getRange(row, 7 + cIndex).setValue(choice); // G列から選択肢
    });
  });

  // 右側のセクションにデータを配置
  questionsList.slice(middleIndex).forEach((item, index) => {
    const row = index + 4; // 右側のセクションもB4から開始
    newSheet.getRange(row, 1 + offset).setValue(item.number); // 中央の列からA列相当の位置: 問題番号
    newSheet.getRange(row, 2 + offset).setValue(item.question); // 中央の列からB列相当の位置: 問題文
    item.choices.forEach((choice, cIndex) => {
      newSheet.getRange(row, 7 + offset + cIndex).setValue(choice); // 中央の列からG列相当の位置から選択肢
    });
  });
}

function sortingName() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getActiveSheet();
  const range = sourceSheet.getRange('A1:Z' + sourceSheet.getLastRow())
  const values = range.getValues();

  // A列の文字列にB列の文字が含まれているか確認し、条件に合う場合にB列以降をA列を基準に並べ替え
  const sortedValues = values.map(row => {
    if (row[0].includes(row[1])) { // A列がB列の文字を含むか確認
      const base = row.slice(1); // B列以降を取得
      base.sort((a, b) => { // A列を基準に並べ替え
        return a.toString().localeCompare(b.toString(), undefined, {numeric: true});
      });
      return [row[0], ...base]; // 並べ替えたデータを結合
    }
    return row; // 条件に合わない場合はそのまま
  });

  // 並べ替えたデータをスプレッドシートに書き戻す
  range.setValues(sortedValues);
}