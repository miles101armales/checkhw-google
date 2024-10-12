import { Controller, Get, Query } from '@nestjs/common';
import { AppService } from './app.service';

@Controller()
export class AppController {
  constructor(private readonly appService: AppService) {}

  @Get('homework')
  async setHomeworkStatus(
    @Query('email') email: string,
    @Query('assignmentNumber') assignmentNumber: number,
    @Query('status') status: string,
  ): Promise<string> {
    try {
      await this.appService.updateAssignmentStatus({
        email,
        assignmentNumber,
        status,
      });
      return 'OK';
    } catch (error) {
      return `${error}`;
    }
  }
}
