import { Injectable, OnModuleInit } from '@nestjs/common';
import { google, sheets_v4 } from 'googleapis';
import * as fs from 'fs';
import { UpdateStatusDto } from './dto/update-hw.dto';
import { emailColumnMap } from './emailColumnMap';
import { Interval } from '@nestjs/schedule';

interface CellUpdate {
  range: string;
  values: any[][];
}

const assignmentColumnMap: Record<number, string> = {
  1: 'I', 2: 'J', 3: 'K', 4: 'L', 5: 'M', 6: 'N', 7: 'O', 8: 'P', 9: 'Q', 10: 'R',
  11: 'S', 12: 'T', 13: 'U', 14: 'V', 15: 'W', 16: 'X', 17: 'Y', 18: 'Z', 19: 'AA',
};

@Injectable()
export class AppService implements OnModuleInit {
  private sheets: sheets_v4.Sheets;
  private spreadsheetId = '1RTbBj4QhR1hK_M9NOJigZyYJqpUoNSNy1huzouhT-Yk';
  private updateBatch: CellUpdate[] = [];
  private logFilePath = 'updated_cells.txt';

  constructor() {
    const auth = this.authenticate();
    this.sheets = google.sheets({ version: 'v4', auth });
  }

  onModuleInit() {
    this.scheduleBatchUpdate();
  }

  private authenticate() {
    const credentials = JSON.parse(fs.readFileSync('credentials.json', 'utf-8'));
    return new google.auth.JWT({
      email: credentials.client_email,
      key: credentials.private_key,
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });
  }

  async updateAssignmentStatus(dto: UpdateStatusDto): Promise<void> {
    const { email, assignmentNumber, status } = dto;
  
    // Приводим email к нижнему регистру
    const lowerCaseEmail = email.toLowerCase();
  
    const rowNumber = Object.entries(emailColumnMap).find(([_, e]) => e.toLowerCase() === lowerCaseEmail)?.[0];
    if (!rowNumber) throw new Error('Почта не найдена');
  
    const column = assignmentColumnMap[assignmentNumber];
    if (!column) throw new Error('Неверный номер домашнего задания');
  
    const cell = `${column}${rowNumber}`;
    
    this.addToBatch(`Выполнение ДЗ!${cell}`, [[status]]);
    this.logUpdatedCell(cell, status);
  
    if (this.updateBatch.length >= 500) {
      await this.executeBatchUpdate();
    }
  }

  private addToBatch(range: string, values: any[][]) {
    this.updateBatch.push({ range, values });
  }

  private logUpdatedCell(cell: string, status: string) {
    const logMessage = `${new Date().toISOString()} - Updated cell: ${cell}, New status: ${status}\n`;
    fs.appendFileSync(this.logFilePath, logMessage);
  }

  @Interval(300000) // 5 минут
  private async scheduleBatchUpdate() {
    if (this.updateBatch.length > 0) {
      await this.executeBatchUpdate();
    }
  }

  private async executeBatchUpdate() {
    try {
      const result = await this.sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: this.spreadsheetId,
        requestBody: {
          valueInputOption: 'RAW',
          data: this.updateBatch,
        },
      });
      console.log(`Batch update completed: ${result.statusText}`);
      this.logBatchUpdate();
      this.updateBatch = []; // Очищаем пакет после обновления
    } catch (error) {
      console.error('Error executing batch update:', error);
    }
  }

  private logBatchUpdate() {
    const logMessage = `${new Date().toISOString()} - Batch update executed. Updated cells:\n`;
    const cellsUpdated = this.updateBatch.map(update => update.range).join(', ');
    fs.appendFileSync(this.logFilePath, logMessage + cellsUpdated + '\n\n');
  }
}