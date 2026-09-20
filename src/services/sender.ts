import type { ActionItem } from '../models/action-item.js';
import { Hey365Error } from '../utils/errors.js';
import { logger } from '../utils/logger.js';
import { shortHash } from '../utils/text.js';
import { createEntity, doAction } from './workiq.js';
import { recordSend } from './store.js';

export interface SendOutcome {
  index: number;
  id: string;
  ok: boolean;
  target: string;
  error?: { code: string; message: string; nextStep: string };
}

/**
 * Sends the stored draft as a reply inside the original thread.
 * Hey365 never creates a new thread implicitly (spec section 10).
 */
export async function sendItem(item: ActionItem, expectedHash?: string): Promise<SendOutcome> {
  const draft = item.draft;
  if (!draft) {
    return failure(item, new Hey365Error('DRAFT_NOT_FOUND'), describeTarget(item));
  }
  // Guard against sending text the user never saw.
  if (expectedHash && expectedHash !== draft.hash) {
    return failure(item, new Hey365Error('DRAFT_MISMATCH'), describeTarget(item));
  }
  if (shortHash(draft.text) !== draft.hash) {
    return failure(item, new Hey365Error('DRAFT_MISMATCH'), describeTarget(item));
  }

  const target = describeTarget(item);

  try {
    switch (item.routing.kind) {
      case 'outlook': {
        // `reply` keeps the conversation id and recipients of the original.
        await doAction(`/me/messages/${item.routing.messageId}/reply`, {
          comment: draft.text,
        });
        break;
      }
      case 'teams-chat': {
        await createEntity(`/chats/${encodeURIComponent(item.routing.chatId)}/messages`, {
          body: { contentType: 'text', content: draft.text },
        });
        break;
      }
      case 'teams-channel': {
        const { teamId, channelId, rootMessageId } = item.routing;
        await createEntity(
          `/teams/${teamId}/channels/${encodeURIComponent(channelId)}/messages/${rootMessageId}/replies`,
          { body: { contentType: 'text', content: draft.text } },
        );
        break;
      }
      default: {
        return failure(item, new Hey365Error('SEND_FAILED', 'Unsupported reply target'), target);
      }
    }
  } catch (error) {
    const hey =
      error instanceof Hey365Error
        ? error.code === 'WORKIQ_PERMISSION_DENIED'
          ? new Hey365Error('WORKIQ_WRITE_DISABLED', error.message)
          : error
        : new Hey365Error('SEND_FAILED', (error as Error).message);
    logger.warn('send failed', { index: item.index, code: hey.code });
    return failure(item, hey, target);
  }

  recordSend(item, target);
  return { index: item.index, id: item.id, ok: true, target };
}

export function describeTarget(item: ActionItem): string {
  switch (item.routing.kind) {
    case 'outlook':
      return `Outlook / ${item.sender.name} / ${item.subject}`;
    case 'teams-chat':
      return `Teams chat / ${item.sender.name}`;
    case 'teams-channel':
      return `Teams channel / ${item.subject || item.sender.name}`;
    default:
      return item.subject;
  }
}

function failure(item: ActionItem, error: Hey365Error, target: string): SendOutcome {
  return {
    index: item.index,
    id: item.id,
    ok: false,
    target,
    error: { code: error.code, message: error.message, nextStep: error.nextStep('ja') },
  };
}
