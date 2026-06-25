import {
  SQSClient,
  SendMessageCommand,
} from '@aws-sdk/client-sqs';

import { Injectable } from '@nestjs/common';

@Injectable()
export class AuditEngineService {
  private readonly sqs = new SQSClient({
      region: process.env.AWS_REGION,
  });
  async enqueueAudit( auditJobId: string ) {
    const command = new SendMessageCommand({
        QueueUrl: process.env.AWS_SQS_AUDIT_QUEUE_URL,
        MessageBody: JSON.stringify({
            auditJobId,
            requestedAt: new Date().toISOString()
        }),
    });
    const result = await this.sqs.send(command);
    console.log('[SQS]',result.MessageId);
    return result;
  }

}