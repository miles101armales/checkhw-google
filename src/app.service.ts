import { Injectable } from '@nestjs/common';
import { google, sheets_v4 } from 'googleapis';
import path from 'path';
import * as fs from 'fs';
import { UpdateStatusDto } from './dto/update-hw.dto';

// Маппинг между номером домашнего задания и столбцом
const assignmentColumnMap: Record<number, string> = {
  1: 'J',
  2: 'K',
  3: 'L',
  4: 'M',
  5: 'N',
  6: 'O',
  7: 'P',
  8: 'Q',
  9: 'R',
  10: 'S',
  11: 'T',
  12: 'U',
  13: 'V',
  14: 'W',
  15: 'X',
  16: 'Y',
  17: 'Z',
  18: 'AA',
};

@Injectable()
export class AppService {
  private sheets: sheets_v4.Sheets;
  private spreadsheetId = '1RTbBj4QhR1hK_M9NOJigZyYJqpUoNSNy1huzouhT-Yk'; // Укажите свой ID таблицы

  constructor() {
    const auth = this.authenticate();
    this.sheets = google.sheets({ version: 'v4', auth });
  }

  // Метод для аутентификации через OAuth2
  private authenticate() {
    const credentials = JSON.parse(
      fs.readFileSync('credentials.json', 'utf-8'),
    );

    // Создаем клиента с использованием ключей сервисного аккаунта
    const oAuth2Client = new google.auth.JWT({
      email: credentials.client_email,
      key: credentials.private_key,
      scopes: ['https://www.googleapis.com/auth/spreadsheets'], // Доступ к Google Sheets API
    });

    return oAuth2Client;
  }

  // Основной метод для обновления статуса домашнего задания
  async updateAssignmentStatus(dto: UpdateStatusDto): Promise<void> {
    const { email, assignmentNumber, status } = dto;

    const range = 'Выполнение ДЗ!A2:AA530'; // Диапазон данных, который нужно просканировать

    // Шаг 1: Получаем данные с Google Sheets
    const response = await this.sheets.spreadsheets.values.get({
      spreadsheetId: this.spreadsheetId,
      range,
    });

    console.log(email);

    const rows = response.data.values;
    if (!rows) throw new Error('Нет данных в таблице');

    // Шаг 2: Находим строку с нужным email
    const rowIndex = rows.findIndex((row) => {
      return row[2] == email;
    });
    if (rowIndex === -1) throw new Error('Почта не найдена');

    console.log(rowIndex);

    // Шаг 3: Обновляем статус домашнего задания
    const column = assignmentColumnMap[assignmentNumber];
    if (!column) throw new Error('Неверный номер домашнего задания');

    const cell = `${column}${rowIndex + 2}`; // Пример: B2, C5 и т.д.
    console.log(cell);

    await this.sheets.spreadsheets.values.update({
      spreadsheetId: this.spreadsheetId,
      range: `Выполнение ДЗ!${cell}`,
      valueInputOption: 'RAW',
      requestBody: {
        values: [[status]],
      },
    });
  }
}
