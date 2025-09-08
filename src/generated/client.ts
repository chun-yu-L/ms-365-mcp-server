import { makeApi, Zodios, type ZodiosOptions } from './hack.js';
import { z } from 'zod';

type microsoft_graph_mailFolder = {
  id?: string | undefined;
  childFolderCount?: (number | null) | undefined;
  displayName?: (string | null) | undefined;
  isHidden?: (boolean | null) | undefined;
  parentFolderId?: (string | null) | undefined;
  totalItemCount?: (number | null) | undefined;
  unreadItemCount?: (number | null) | undefined;
  childFolders?: Array<microsoft_graph_mailFolder> | undefined;
  messageRules?: Array<microsoft_graph_messageRule> | undefined;
  messages?: Array<microsoft_graph_message> | undefined;
  multiValueExtendedProperties?:
    | Array<microsoft_graph_multiValueLegacyExtendedProperty>
    | undefined;
  singleValueExtendedProperties?:
    | Array<microsoft_graph_singleValueLegacyExtendedProperty>
    | undefined;
};
type microsoft_graph_messageRule = {
  id?: string | undefined;
  actions?: microsoft_graph_messageRuleActions | undefined;
  conditions?: microsoft_graph_messageRulePredicates | undefined;
  displayName?: (string | null) | undefined;
  exceptions?: microsoft_graph_messageRulePredicates | undefined;
  hasError?: (boolean | null) | undefined;
  isEnabled?: (boolean | null) | undefined;
  isReadOnly?: (boolean | null) | undefined;
  sequence?: (number | null) | undefined;
};
type microsoft_graph_messageRuleActions = {
  assignCategories?: Array<string | null> | undefined;
  copyToFolder?: (string | null) | undefined;
  delete?: (boolean | null) | undefined;
  forwardAsAttachmentTo?: Array<microsoft_graph_recipient> | undefined;
  forwardTo?: Array<microsoft_graph_recipient> | undefined;
  markAsRead?: (boolean | null) | undefined;
  markImportance?: microsoft_graph_importance | undefined;
  moveToFolder?: (string | null) | undefined;
  permanentDelete?: (boolean | null) | undefined;
  redirectTo?: Array<microsoft_graph_recipient> | undefined;
  stopProcessingRules?: (boolean | null) | undefined;
};
type microsoft_graph_recipient = {
  emailAddress?: microsoft_graph_emailAddress | undefined;
};
type microsoft_graph_emailAddress = {
  address?: (string | null) | undefined;
  name?: (string | null) | undefined;
};
type microsoft_graph_importance = 'low' | 'normal' | 'high';
type microsoft_graph_messageRulePredicates = {
  bodyContains?: Array<string | null> | undefined;
  bodyOrSubjectContains?: Array<string | null> | undefined;
  categories?: Array<string | null> | undefined;
  fromAddresses?: Array<microsoft_graph_recipient> | undefined;
  hasAttachments?: (boolean | null) | undefined;
  headerContains?: Array<string | null> | undefined;
  importance?: microsoft_graph_importance | undefined;
  isApprovalRequest?: (boolean | null) | undefined;
  isAutomaticForward?: (boolean | null) | undefined;
  isAutomaticReply?: (boolean | null) | undefined;
  isEncrypted?: (boolean | null) | undefined;
  isMeetingRequest?: (boolean | null) | undefined;
  isMeetingResponse?: (boolean | null) | undefined;
  isNonDeliveryReport?: (boolean | null) | undefined;
  isPermissionControlled?: (boolean | null) | undefined;
  isReadReceipt?: (boolean | null) | undefined;
  isSigned?: (boolean | null) | undefined;
  isVoicemail?: (boolean | null) | undefined;
  messageActionFlag?: microsoft_graph_messageActionFlag | undefined;
  notSentToMe?: (boolean | null) | undefined;
  recipientContains?: Array<string | null> | undefined;
  senderContains?: Array<string | null> | undefined;
  sensitivity?: microsoft_graph_sensitivity | undefined;
  sentCcMe?: (boolean | null) | undefined;
  sentOnlyToMe?: (boolean | null) | undefined;
} & {
  [key: string]: any;
};
type microsoft_graph_messageActionFlag =
  | 'any'
  | 'call'
  | 'doNotForward'
  | 'followUp'
  | 'fyi'
  | 'forward'
  | 'noResponseNecessary'
  | 'read'
  | 'reply'
  | 'replyToAll'
  | 'review';
type microsoft_graph_sensitivity = 'normal' | 'personal' | 'private' | 'confidential';
type microsoft_graph_message = {
  body?: microsoft_graph_itemBody | undefined;
  subject?: (string | null) | undefined;
  attachments?: Array<microsoft_graph_attachment> | undefined;
  bccRecipients?: Array<microsoft_graph_recipient> | undefined;
  bodyPreview?: (string | null) | undefined;
  ccRecipients?: Array<microsoft_graph_recipient> | undefined;
  conversationId?: (string | null) | undefined;
  conversationIndex?: (string | null) | undefined;
  flag?: microsoft_graph_followupFlag | undefined;
  from?: microsoft_graph_recipient | undefined;
  hasAttachments?: (boolean | null) | undefined;
  importance?: microsoft_graph_importance | undefined;
  inferenceClassification?: microsoft_graph_inferenceClassificationType | undefined;
  internetMessageHeaders?: Array<microsoft_graph_internetMessageHeader> | undefined;
  internetMessageId?: (string | null) | undefined;
  isDeliveryReceiptRequested?: (boolean | null) | undefined;
  isDraft?: (boolean | null) | undefined;
  isRead?: (boolean | null) | undefined;
  isReadReceiptRequested?: (boolean | null) | undefined;
  parentFolderId?: (string | null) | undefined;
  receivedDateTime?: (string | null) | undefined;
  replyTo?: Array<microsoft_graph_recipient> | undefined;
  sender?: microsoft_graph_recipient | undefined;
  sentDateTime?: (string | null) | undefined;
  toRecipients?: Array<microsoft_graph_recipient> | undefined;
} & {
  [key: string]: any;
};
type microsoft_graph_itemBody = {
  content?: (string | null) | undefined;
  contentType?: microsoft_graph_bodyType | undefined;
};
type microsoft_graph_bodyType = 'text' | 'html';
type microsoft_graph_attachment = {
  id?: string | undefined;
  contentType?: (string | null) | undefined;
  isInline?: boolean | undefined;
  lastModifiedDateTime?: (string | null) | undefined;
  name?: (string | null) | undefined;
  size?: number | undefined;
};
type microsoft_graph_followupFlag = {
  completedDateTime?: microsoft_graph_dateTimeTimeZone | undefined;
  dueDateTime?: microsoft_graph_dateTimeTimeZone | undefined;
  flagStatus?: microsoft_graph_followupFlagStatus | undefined;
  startDateTime?: microsoft_graph_dateTimeTimeZone | undefined;
};
type microsoft_graph_dateTimeTimeZone = {
  dateTime?: string | undefined;
  timeZone?: (string | null) | undefined;
};
type microsoft_graph_followupFlagStatus = 'notFlagged' | 'complete' | 'flagged';
type microsoft_graph_inferenceClassificationType = 'focused' | 'other';
type microsoft_graph_internetMessageHeader = {
  name?: (string | null) | undefined;
  value?: (string | null) | undefined;
};
type microsoft_graph_multiValueLegacyExtendedProperty = {
  id?: string | undefined;
  value?: Array<string | null> | undefined;
};
type microsoft_graph_singleValueLegacyExtendedProperty = {
  id?: string | undefined;
  value?: (string | null) | undefined;
};

const microsoft_graph_emailAddress = z
  .object({
    address: z.string().describe('The email address of the person or entity.').nullish(),
    name: z.string().describe('The display name of the person or entity.').nullish(),
  })
  .strict();
const microsoft_graph_recipient = z
  .object({ emailAddress: microsoft_graph_emailAddress.optional() })
  .strict();
const microsoft_graph_importance = z.enum(['low', 'normal', 'high']);
const microsoft_graph_messageRuleActions = z
  .object({
    assignCategories: z
      .array(z.string().nullable())
      .describe('A list of categories to be assigned to a message.')
      .optional(),
    copyToFolder: z
      .string()
      .describe('The ID of a folder that a message is to be copied to.')
      .nullish(),
    delete: z
      .boolean()
      .describe('Indicates whether a message should be moved to the Deleted Items folder.')
      .nullish(),
    forwardAsAttachmentTo: z
      .array(microsoft_graph_recipient)
      .describe(
        'The email addresses of the recipients to which a message should be forwarded as an attachment.'
      )
      .optional(),
    forwardTo: z
      .array(microsoft_graph_recipient)
      .describe('The email addresses of the recipients to which a message should be forwarded.')
      .optional(),
    markAsRead: z
      .boolean()
      .describe('Indicates whether a message should be marked as read.')
      .nullish(),
    markImportance: microsoft_graph_importance.optional(),
    moveToFolder: z
      .string()
      .describe('The ID of the folder that a message will be moved to.')
      .nullish(),
    permanentDelete: z
      .boolean()
      .describe(
        'Indicates whether a message should be permanently deleted and not saved to the Deleted Items folder.'
      )
      .nullish(),
    redirectTo: z
      .array(microsoft_graph_recipient)
      .describe('The email addresses to which a message should be redirected.')
      .optional(),
    stopProcessingRules: z
      .boolean()
      .describe('Indicates whether subsequent rules should be evaluated.')
      .nullish(),
  })
  .strict();
const microsoft_graph_messageActionFlag = z.enum([
  'any',
  'call',
  'doNotForward',
  'followUp',
  'fyi',
  'forward',
  'noResponseNecessary',
  'read',
  'reply',
  'replyToAll',
  'review',
]);
const microsoft_graph_sensitivity = z.enum(['normal', 'personal', 'private', 'confidential']);
const microsoft_graph_messageRulePredicates = z
  .object({
    bodyContains: z
      .array(z.string().nullable())
      .describe(
        'Represents the strings that should appear in the body of an incoming message in order for the condition or exception to apply.'
      )
      .optional(),
    bodyOrSubjectContains: z
      .array(z.string().nullable())
      .describe(
        'Represents the strings that should appear in the body or subject of an incoming message in order for the condition or exception to apply.'
      )
      .optional(),
    categories: z
      .array(z.string().nullable())
      .describe(
        'Represents the categories that an incoming message should be labeled with in order for the condition or exception to apply.'
      )
      .optional(),
    fromAddresses: z
      .array(microsoft_graph_recipient)
      .describe(
        'Represents the specific sender email addresses of an incoming message in order for the condition or exception to apply.'
      )
      .optional(),
    hasAttachments: z
      .boolean()
      .describe(
        'Indicates whether an incoming message must have attachments in order for the condition or exception to apply.'
      )
      .nullish(),
    headerContains: z
      .array(z.string().nullable())
      .describe(
        'Represents the strings that appear in the headers of an incoming message in order for the condition or exception to apply.'
      )
      .optional(),
    importance: microsoft_graph_importance.optional(),
    isApprovalRequest: z
      .boolean()
      .describe(
        'Indicates whether an incoming message must be an approval request in order for the condition or exception to apply.'
      )
      .nullish(),
    isAutomaticForward: z
      .boolean()
      .describe(
        'Indicates whether an incoming message must be automatically forwarded in order for the condition or exception to apply.'
      )
      .nullish(),
    isAutomaticReply: z
      .boolean()
      .describe(
        'Indicates whether an incoming message must be an auto reply in order for the condition or exception to apply.'
      )
      .nullish(),
    isEncrypted: z
      .boolean()
      .describe(
        'Indicates whether an incoming message must be encrypted in order for the condition or exception to apply.'
      )
      .nullish(),
    isMeetingRequest: z
      .boolean()
      .describe(
        'Indicates whether an incoming message must be a meeting request in order for the condition or exception to apply.'
      )
      .nullish(),
    isMeetingResponse: z
      .boolean()
      .describe(
        'Indicates whether an incoming message must be a meeting response in order for the condition or exception to apply.'
      )
      .nullish(),
    isNonDeliveryReport: z
      .boolean()
      .describe(
        'Indicates whether an incoming message must be a non-delivery report in order for the condition or exception to apply.'
      )
      .nullish(),
    isPermissionControlled: z
      .boolean()
      .describe(
        'Indicates whether an incoming message must be permission controlled (RMS-protected) in order for the condition or exception to apply.'
      )
      .nullish(),
    isReadReceipt: z
      .boolean()
      .describe(
        'Indicates whether an incoming message must be a read receipt in order for the condition or exception to apply.'
      )
      .nullish(),
    isSigned: z
      .boolean()
      .describe(
        'Indicates whether an incoming message must be S/MIME-signed in order for the condition or exception to apply.'
      )
      .nullish(),
    isVoicemail: z
      .boolean()
      .describe(
        'Indicates whether an incoming message must be a voice mail in order for the condition or exception to apply.'
      )
      .nullish(),
    messageActionFlag: microsoft_graph_messageActionFlag.optional(),
    notSentToMe: z
      .boolean()
      .describe(
        'Indicates whether the owner of the mailbox must not be a recipient of an incoming message in order for the condition or exception to apply.'
      )
      .nullish(),
    recipientContains: z
      .array(z.string().nullable())
      .describe(
        'Represents the strings that appear in either the toRecipients or ccRecipients properties of an incoming message in order for the condition or exception to apply.'
      )
      .optional(),
    senderContains: z
      .array(z.string().nullable())
      .describe(
        'Represents the strings that appear in the from property of an incoming message in order for the condition or exception to apply.'
      )
      .optional(),
    sensitivity: microsoft_graph_sensitivity.optional(),
    sentCcMe: z
      .boolean()
      .describe(
        'Indicates whether the owner of the mailbox must be in the ccRecipients property of an incoming message in order for the condition or exception to apply.'
      )
      .nullish(),
    sentOnlyToMe: z
      .boolean()
      .describe(
        'Indicates whether the owner of the mailbox must be the only recipient in an incoming message in order for the condition or exception to apply.'
      )
      .nullish(),
  })
  .strict()
  .passthrough();
const microsoft_graph_messageRule = z
  .object({
    id: z.string().describe('The unique identifier for an entity. Read-only.').optional(),
    actions: microsoft_graph_messageRuleActions.optional(),
    conditions: microsoft_graph_messageRulePredicates
      .describe('[Note: Simplified from 30 properties to 25 most common ones]')
      .optional(),
    displayName: z.string().describe('The display name of the rule.').nullish(),
    exceptions: microsoft_graph_messageRulePredicates
      .describe('[Note: Simplified from 30 properties to 25 most common ones]')
      .optional(),
    hasError: z
      .boolean()
      .describe('Indicates whether the rule is in an error condition. Read-only.')
      .nullish(),
    isEnabled: z
      .boolean()
      .describe('Indicates whether the rule is enabled to be applied to messages.')
      .nullish(),
    isReadOnly: z
      .boolean()
      .describe(
        'Indicates if the rule is read-only and cannot be modified or deleted by the rules REST API.'
      )
      .nullish(),
    sequence: z
      .number()
      .gte(-2147483648)
      .lte(2147483647)
      .describe('Indicates the order in which the rule is executed, among other rules.')
      .nullish(),
  })
  .strict();
const microsoft_graph_bodyType = z.enum(['text', 'html']);
const microsoft_graph_itemBody = z
  .object({
    content: z.string().describe('The content of the item.').nullish(),
    contentType: microsoft_graph_bodyType.optional(),
  })
  .strict();
const microsoft_graph_attachment = z
  .object({
    id: z.string().describe('The unique identifier for an entity. Read-only.').optional(),
    contentType: z.string().describe('The MIME type.').nullish(),
    isInline: z
      .boolean()
      .describe('true if the attachment is an inline attachment; otherwise, false.')
      .optional(),
    lastModifiedDateTime: z
      .string()
      .regex(
        /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
      )
      .datetime({ offset: true })
      .describe(
        'The Timestamp type represents date and time information using ISO 8601 format and is always in UTC time. For example, midnight UTC on Jan 1, 2014 is 2014-01-01T00:00:00Z'
      )
      .nullish(),
    name: z.string().describe("The attachment's file name.").nullish(),
    size: z
      .number()
      .gte(-2147483648)
      .lte(2147483647)
      .describe('The length of the attachment in bytes.')
      .optional(),
  })
  .strict();
const microsoft_graph_dateTimeTimeZone = z
  .object({
    dateTime: z
      .string()
      .describe(
        'A single point of time in a combined date and time representation ({date}T{time}; for example, 2017-08-29T04:00:00.0000000).'
      )
      .optional(),
    timeZone: z
      .string()
      .describe(
        "Represents a time zone, for example, 'Pacific Standard Time'. See below for more possible values."
      )
      .nullish(),
  })
  .strict();
const microsoft_graph_followupFlagStatus = z.enum(['notFlagged', 'complete', 'flagged']);
const microsoft_graph_followupFlag = z
  .object({
    completedDateTime: microsoft_graph_dateTimeTimeZone.optional(),
    dueDateTime: microsoft_graph_dateTimeTimeZone.optional(),
    flagStatus: microsoft_graph_followupFlagStatus.optional(),
    startDateTime: microsoft_graph_dateTimeTimeZone.optional(),
  })
  .strict();
const microsoft_graph_inferenceClassificationType = z.enum(['focused', 'other']);
const microsoft_graph_internetMessageHeader = z
  .object({
    name: z.string().describe('Represents the key in a key-value pair.').nullish(),
    value: z.string().describe('The value in a key-value pair.').nullish(),
  })
  .strict();
const microsoft_graph_message = z
  .object({
    body: microsoft_graph_itemBody.optional(),
    subject: z.string().describe('The subject of the message.').nullish(),
    attachments: z
      .array(microsoft_graph_attachment)
      .describe('The fileAttachment and itemAttachment attachments for the message.')
      .optional(),
    bccRecipients: z
      .array(microsoft_graph_recipient)
      .describe('The Bcc: recipients for the message.')
      .optional(),
    bodyPreview: z
      .string()
      .describe('The first 255 characters of the message body. It is in text format.')
      .nullish(),
    ccRecipients: z
      .array(microsoft_graph_recipient)
      .describe('The Cc: recipients for the message.')
      .optional(),
    conversationId: z
      .string()
      .describe('The ID of the conversation the email belongs to.')
      .nullish(),
    conversationIndex: z
      .string()
      .describe('Indicates the position of the message within the conversation.')
      .nullish(),
    flag: microsoft_graph_followupFlag.optional(),
    from: microsoft_graph_recipient.optional(),
    hasAttachments: z
      .boolean()
      .describe(
        "Indicates whether the message has attachments. This property doesn't include inline attachments, so if a message contains only inline attachments, this property is false. To verify the existence of inline attachments, parse the body property to look for a src attribute, such as <IMG src='cid:image001.jpg@01D26CD8.6C05F070'>."
      )
      .nullish(),
    importance: microsoft_graph_importance.optional(),
    inferenceClassification: microsoft_graph_inferenceClassificationType.optional(),
    internetMessageHeaders: z
      .array(microsoft_graph_internetMessageHeader)
      .describe(
        'A collection of message headers defined by RFC5322. The set includes message headers indicating the network path taken by a message from the sender to the recipient. It can also contain custom message headers that hold app data for the message.  Returned only on applying a $select query option. Read-only.'
      )
      .optional(),
    internetMessageId: z
      .string()
      .describe('The message ID in the format specified by RFC2822.')
      .nullish(),
    isDeliveryReceiptRequested: z
      .boolean()
      .describe('Indicates whether a read receipt is requested for the message.')
      .nullish(),
    isDraft: z
      .boolean()
      .describe(
        "Indicates whether the message is a draft. A message is a draft if it hasn't been sent yet."
      )
      .nullish(),
    isRead: z.boolean().describe('Indicates whether the message has been read.').nullish(),
    isReadReceiptRequested: z
      .boolean()
      .describe('Indicates whether a read receipt is requested for the message.')
      .nullish(),
    parentFolderId: z
      .string()
      .describe("The unique identifier for the message's parent mailFolder.")
      .nullish(),
    receivedDateTime: z
      .string()
      .regex(
        /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
      )
      .datetime({ offset: true })
      .describe(
        'The date and time the message was received.  The date and time information uses ISO 8601 format and is always in UTC time. For example, midnight UTC on Jan 1, 2014 is 2014-01-01T00:00:00Z.'
      )
      .nullish(),
    replyTo: z
      .array(microsoft_graph_recipient)
      .describe('The email addresses to use when replying.')
      .optional(),
    sender: microsoft_graph_recipient.optional(),
    sentDateTime: z
      .string()
      .regex(
        /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
      )
      .datetime({ offset: true })
      .describe(
        'The date and time the message was sent.  The date and time information uses ISO 8601 format and is always in UTC time. For example, midnight UTC on Jan 1, 2014 is 2014-01-01T00:00:00Z.'
      )
      .nullish(),
    toRecipients: z
      .array(microsoft_graph_recipient)
      .describe('The To: recipients for the message.')
      .optional(),
  })
  .strict()
  .passthrough();
const microsoft_graph_multiValueLegacyExtendedProperty = z
  .object({
    id: z.string().describe('The unique identifier for an entity. Read-only.').optional(),
    value: z.array(z.string().nullable()).describe('A collection of property values.').optional(),
  })
  .strict();
const microsoft_graph_singleValueLegacyExtendedProperty = z
  .object({
    id: z.string().describe('The unique identifier for an entity. Read-only.').optional(),
    value: z.string().describe('A property value.').nullish(),
  })
  .strict();
const microsoft_graph_mailFolder: z.ZodType<microsoft_graph_mailFolder> = z.lazy(() =>
  z
    .object({
      id: z.string().describe('The unique identifier for an entity. Read-only.').optional(),
      childFolderCount: z
        .number()
        .gte(-2147483648)
        .lte(2147483647)
        .describe('The number of immediate child mailFolders in the current mailFolder.')
        .nullish(),
      displayName: z.string().describe("The mailFolder's display name.").nullish(),
      isHidden: z
        .boolean()
        .describe(
          'Indicates whether the mailFolder is hidden. This property can be set only when creating the folder. Find more information in Hidden mail folders.'
        )
        .nullish(),
      parentFolderId: z
        .string()
        .describe("The unique identifier for the mailFolder's parent mailFolder.")
        .nullish(),
      totalItemCount: z
        .number()
        .gte(-2147483648)
        .lte(2147483647)
        .describe('The number of items in the mailFolder.')
        .nullish(),
      unreadItemCount: z
        .number()
        .gte(-2147483648)
        .lte(2147483647)
        .describe('The number of items in the mailFolder marked as unread.')
        .nullish(),
      childFolders: z
        .array(microsoft_graph_mailFolder)
        .describe('The collection of child folders in the mailFolder.')
        .optional(),
      messageRules: z
        .array(microsoft_graph_messageRule)
        .describe("The collection of rules that apply to the user's Inbox folder.")
        .optional(),
      messages: z
        .array(microsoft_graph_message)
        .describe('The collection of messages in the mailFolder.')
        .optional(),
      multiValueExtendedProperties: z
        .array(microsoft_graph_multiValueLegacyExtendedProperty)
        .describe(
          'The collection of multi-value extended properties defined for the mailFolder. Read-only. Nullable.'
        )
        .optional(),
      singleValueExtendedProperties: z
        .array(microsoft_graph_singleValueLegacyExtendedProperty)
        .describe(
          'The collection of single-value extended properties defined for the mailFolder. Read-only. Nullable.'
        )
        .optional(),
    })
    .strict()
);
const microsoft_graph_mailFolderCollectionResponse = z
  .object({
    '@odata.count': z.number().int().nullable(),
    '@odata.nextLink': z.string().nullable(),
    value: z.array(microsoft_graph_mailFolder),
  })
  .partial()
  .strict();
const microsoft_graph_ODataErrors_ErrorDetails = z
  .object({ code: z.string(), message: z.string(), target: z.string().nullish() })
  .strict();
const microsoft_graph_ODataErrors_InnerError = z
  .object({
    'request-id': z.string().describe('Request Id as tracked internally by the service').nullish(),
    'client-request-id': z
      .string()
      .describe('Client request Id as sent by the client application.')
      .nullish(),
    date: z
      .string()
      .regex(
        /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
      )
      .datetime({ offset: true })
      .describe('Date when the error occured.')
      .nullish(),
  })
  .strict();
const microsoft_graph_ODataErrors_MainError = z
  .object({
    code: z.string(),
    message: z.string(),
    target: z.string().nullish(),
    details: z.array(microsoft_graph_ODataErrors_ErrorDetails).optional(),
    innerError: microsoft_graph_ODataErrors_InnerError.optional(),
  })
  .strict();
const microsoft_graph_ODataErrors_ODataError = z
  .object({ error: microsoft_graph_ODataErrors_MainError })
  .strict();
const microsoft_graph_messageCollectionResponse = z
  .object({
    '@odata.count': z.number().int().nullable(),
    '@odata.nextLink': z.string().nullable(),
    value: z.array(microsoft_graph_message),
  })
  .partial()
  .strict();
const microsoft_graph_attachmentCollectionResponse = z
  .object({
    '@odata.count': z.number().int().nullable(),
    '@odata.nextLink': z.string().nullable(),
    value: z.array(microsoft_graph_attachment),
  })
  .partial()
  .strict();

export const schemas = {
  microsoft_graph_emailAddress,
  microsoft_graph_recipient,
  microsoft_graph_importance,
  microsoft_graph_messageRuleActions,
  microsoft_graph_messageActionFlag,
  microsoft_graph_sensitivity,
  microsoft_graph_messageRulePredicates,
  microsoft_graph_messageRule,
  microsoft_graph_bodyType,
  microsoft_graph_itemBody,
  microsoft_graph_attachment,
  microsoft_graph_dateTimeTimeZone,
  microsoft_graph_followupFlagStatus,
  microsoft_graph_followupFlag,
  microsoft_graph_inferenceClassificationType,
  microsoft_graph_internetMessageHeader,
  microsoft_graph_message,
  microsoft_graph_multiValueLegacyExtendedProperty,
  microsoft_graph_singleValueLegacyExtendedProperty,
  microsoft_graph_mailFolder,
  microsoft_graph_mailFolderCollectionResponse,
  microsoft_graph_ODataErrors_ErrorDetails,
  microsoft_graph_ODataErrors_InnerError,
  microsoft_graph_ODataErrors_MainError,
  microsoft_graph_ODataErrors_ODataError,
  microsoft_graph_messageCollectionResponse,
  microsoft_graph_attachmentCollectionResponse,
};

const endpoints = makeApi([
  {
    method: 'get',
    path: '/me/mailFolders',
    alias: 'list-mail-folders',
    description: `Get the mail folder collection directly under the root folder of the signed-in user. The returned collection includes any mail search folders directly under the root. By default, this operation does not return hidden folders. Use a query parameter includeHiddenFolders to include them in the response. This operation does not return all mail folders in a mailbox, only the child folders of the root folder. To return all mail folders in a mailbox, each child folder must be traversed separately.`,
    requestFormat: 'json',
    parameters: [
      {
        name: 'includeHiddenFolders',
        type: 'Query',
        schema: z.string().describe('Include Hidden Folders').optional(),
      },
      {
        name: '$top',
        type: 'Query',
        schema: z.number().int().gte(0).describe('Show only the first n items').optional(),
      },
      {
        name: '$skip',
        type: 'Query',
        schema: z.number().int().gte(0).describe('Skip the first n items').optional(),
      },
      {
        name: '$search',
        type: 'Query',
        schema: z.string().describe('Search items by search phrases').optional(),
      },
      {
        name: '$filter',
        type: 'Query',
        schema: z.string().describe('Filter items by property values').optional(),
      },
      {
        name: '$count',
        type: 'Query',
        schema: z.boolean().describe('Include count of items').optional(),
      },
      {
        name: '$orderby',
        type: 'Query',
        schema: z.array(z.string()).describe('Order items by property values').optional(),
      },
      {
        name: '$select',
        type: 'Query',
        schema: z.array(z.string()).describe('Select properties to be returned').optional(),
      },
      {
        name: '$expand',
        type: 'Query',
        schema: z.array(z.string()).describe('Expand related entities').optional(),
      },
    ],
    response: z.void(),
    errors: [
      {
        status: NaN,
        description: `Retrieved collection`,
        schema: microsoft_graph_mailFolderCollectionResponse,
      },
      {
        status: NaN,
        description: `error`,
        schema: microsoft_graph_ODataErrors_ODataError,
      },
      {
        status: NaN,
        description: `error`,
        schema: microsoft_graph_ODataErrors_ODataError,
      },
    ],
  },
  {
    method: 'get',
    path: '/me/mailFolders/:mailFolderId/messages',
    alias: 'list-mail-folder-messages',
    description: `Get all the messages in the specified user&#x27;s mailbox, or those messages in a specified folder in the mailbox.`,
    requestFormat: 'json',
    parameters: [
      {
        name: '$top',
        type: 'Query',
        schema: z.number().int().gte(0).describe('Show only the first n items').optional(),
      },
      {
        name: '$skip',
        type: 'Query',
        schema: z.number().int().gte(0).describe('Skip the first n items').optional(),
      },
      {
        name: '$search',
        type: 'Query',
        schema: z.string().describe('Search items by search phrases').optional(),
      },
      {
        name: '$filter',
        type: 'Query',
        schema: z.string().describe('Filter items by property values').optional(),
      },
      {
        name: '$count',
        type: 'Query',
        schema: z.boolean().describe('Include count of items').optional(),
      },
      {
        name: '$orderby',
        type: 'Query',
        schema: z.array(z.string()).describe('Order items by property values').optional(),
      },
      {
        name: '$select',
        type: 'Query',
        schema: z.array(z.string()).describe('Select properties to be returned').optional(),
      },
      {
        name: '$expand',
        type: 'Query',
        schema: z.array(z.string()).describe('Expand related entities').optional(),
      },
    ],
    response: z.void(),
    errors: [
      {
        status: NaN,
        description: `Retrieved collection`,
        schema: microsoft_graph_messageCollectionResponse,
      },
      {
        status: NaN,
        description: `error`,
        schema: microsoft_graph_ODataErrors_ODataError,
      },
      {
        status: NaN,
        description: `error`,
        schema: microsoft_graph_ODataErrors_ODataError,
      },
    ],
  },
  {
    method: 'get',
    path: '/me/messages',
    alias: 'list-mail-messages',
    description: `Get an open extension (openTypeExtension object) identified by name or fully qualified name. The table in the Permissions section lists the resources that support open extensions. The following table lists the three scenarios where you can get an open extension from a supported resource instance.`,
    requestFormat: 'json',
    parameters: [
      {
        name: 'includeHiddenMessages',
        type: 'Query',
        schema: z.string().describe('Include Hidden Messages').optional(),
      },
      {
        name: '$top',
        type: 'Query',
        schema: z.number().int().gte(0).describe('Show only the first n items').optional(),
      },
      {
        name: '$skip',
        type: 'Query',
        schema: z.number().int().gte(0).describe('Skip the first n items').optional(),
      },
      {
        name: '$search',
        type: 'Query',
        schema: z.string().describe('Search items by search phrases').optional(),
      },
      {
        name: '$filter',
        type: 'Query',
        schema: z.string().describe('Filter items by property values').optional(),
      },
      {
        name: '$count',
        type: 'Query',
        schema: z.boolean().describe('Include count of items').optional(),
      },
      {
        name: '$orderby',
        type: 'Query',
        schema: z.array(z.string()).describe('Order items by property values').optional(),
      },
      {
        name: '$select',
        type: 'Query',
        schema: z.array(z.string()).describe('Select properties to be returned').optional(),
      },
      {
        name: '$expand',
        type: 'Query',
        schema: z.array(z.string()).describe('Expand related entities').optional(),
      },
    ],
    response: z.void(),
    errors: [
      {
        status: NaN,
        description: `Retrieved collection`,
        schema: microsoft_graph_messageCollectionResponse,
      },
      {
        status: NaN,
        description: `error`,
        schema: microsoft_graph_ODataErrors_ODataError,
      },
      {
        status: NaN,
        description: `error`,
        schema: microsoft_graph_ODataErrors_ODataError,
      },
    ],
  },
  {
    method: 'get',
    path: '/me/messages/:messageId',
    alias: 'get-mail-message',
    description: `Get an open extension (openTypeExtension object) identified by name or fully qualified name. The table in the Permissions section lists the resources that support open extensions. The following table lists the three scenarios where you can get an open extension from a supported resource instance.`,
    requestFormat: 'json',
    parameters: [
      {
        name: '$select',
        type: 'Query',
        schema: z.array(z.string()).describe('Select properties to be returned').optional(),
      },
      {
        name: '$expand',
        type: 'Query',
        schema: z.array(z.string()).describe('Expand related entities').optional(),
      },
    ],
    response: z.void(),
    errors: [
      {
        status: NaN,
        description: `Retrieved navigation property`,
        schema: z
          .object({
            body: microsoft_graph_itemBody.optional(),
            subject: z.string().describe('The subject of the message.').nullish(),
            attachments: z
              .array(microsoft_graph_attachment)
              .describe('The fileAttachment and itemAttachment attachments for the message.')
              .optional(),
            bccRecipients: z
              .array(microsoft_graph_recipient)
              .describe('The Bcc: recipients for the message.')
              .optional(),
            bodyPreview: z
              .string()
              .describe('The first 255 characters of the message body. It is in text format.')
              .nullish(),
            ccRecipients: z
              .array(microsoft_graph_recipient)
              .describe('The Cc: recipients for the message.')
              .optional(),
            conversationId: z
              .string()
              .describe('The ID of the conversation the email belongs to.')
              .nullish(),
            conversationIndex: z
              .string()
              .describe('Indicates the position of the message within the conversation.')
              .nullish(),
            flag: microsoft_graph_followupFlag.optional(),
            from: microsoft_graph_recipient.optional(),
            hasAttachments: z
              .boolean()
              .describe(
                "Indicates whether the message has attachments. This property doesn't include inline attachments, so if a message contains only inline attachments, this property is false. To verify the existence of inline attachments, parse the body property to look for a src attribute, such as <IMG src='cid:image001.jpg@01D26CD8.6C05F070'>."
              )
              .nullish(),
            importance: microsoft_graph_importance.optional(),
            inferenceClassification: microsoft_graph_inferenceClassificationType.optional(),
            internetMessageHeaders: z
              .array(microsoft_graph_internetMessageHeader)
              .describe(
                'A collection of message headers defined by RFC5322. The set includes message headers indicating the network path taken by a message from the sender to the recipient. It can also contain custom message headers that hold app data for the message.  Returned only on applying a $select query option. Read-only.'
              )
              .optional(),
            internetMessageId: z
              .string()
              .describe('The message ID in the format specified by RFC2822.')
              .nullish(),
            isDeliveryReceiptRequested: z
              .boolean()
              .describe('Indicates whether a read receipt is requested for the message.')
              .nullish(),
            isDraft: z
              .boolean()
              .describe(
                "Indicates whether the message is a draft. A message is a draft if it hasn't been sent yet."
              )
              .nullish(),
            isRead: z.boolean().describe('Indicates whether the message has been read.').nullish(),
            isReadReceiptRequested: z
              .boolean()
              .describe('Indicates whether a read receipt is requested for the message.')
              .nullish(),
            parentFolderId: z
              .string()
              .describe("The unique identifier for the message's parent mailFolder.")
              .nullish(),
            receivedDateTime: z
              .string()
              .regex(
                /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
              )
              .datetime({ offset: true })
              .describe(
                'The date and time the message was received.  The date and time information uses ISO 8601 format and is always in UTC time. For example, midnight UTC on Jan 1, 2014 is 2014-01-01T00:00:00Z.'
              )
              .nullish(),
            replyTo: z
              .array(microsoft_graph_recipient)
              .describe('The email addresses to use when replying.')
              .optional(),
            sender: microsoft_graph_recipient.optional(),
            sentDateTime: z
              .string()
              .regex(
                /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
              )
              .datetime({ offset: true })
              .describe(
                'The date and time the message was sent.  The date and time information uses ISO 8601 format and is always in UTC time. For example, midnight UTC on Jan 1, 2014 is 2014-01-01T00:00:00Z.'
              )
              .nullish(),
            toRecipients: z
              .array(microsoft_graph_recipient)
              .describe('The To: recipients for the message.')
              .optional(),
          })
          .strict()
          .passthrough(),
      },
      {
        status: NaN,
        description: `error`,
        schema: microsoft_graph_ODataErrors_ODataError,
      },
      {
        status: NaN,
        description: `error`,
        schema: microsoft_graph_ODataErrors_ODataError,
      },
    ],
  },
  {
    method: 'get',
    path: '/me/messages/:messageId/attachments',
    alias: 'list-mail-attachments',
    description: `Retrieve a list of attachment objects.`,
    requestFormat: 'json',
    parameters: [
      {
        name: '$top',
        type: 'Query',
        schema: z.number().int().gte(0).describe('Show only the first n items').optional(),
      },
      {
        name: '$skip',
        type: 'Query',
        schema: z.number().int().gte(0).describe('Skip the first n items').optional(),
      },
      {
        name: '$search',
        type: 'Query',
        schema: z.string().describe('Search items by search phrases').optional(),
      },
      {
        name: '$filter',
        type: 'Query',
        schema: z.string().describe('Filter items by property values').optional(),
      },
      {
        name: '$count',
        type: 'Query',
        schema: z.boolean().describe('Include count of items').optional(),
      },
      {
        name: '$orderby',
        type: 'Query',
        schema: z.array(z.string()).describe('Order items by property values').optional(),
      },
      {
        name: '$select',
        type: 'Query',
        schema: z.array(z.string()).describe('Select properties to be returned').optional(),
      },
      {
        name: '$expand',
        type: 'Query',
        schema: z.array(z.string()).describe('Expand related entities').optional(),
      },
    ],
    response: z.void(),
    errors: [
      {
        status: NaN,
        description: `Retrieved collection`,
        schema: microsoft_graph_attachmentCollectionResponse,
      },
      {
        status: NaN,
        description: `error`,
        schema: microsoft_graph_ODataErrors_ODataError,
      },
      {
        status: NaN,
        description: `error`,
        schema: microsoft_graph_ODataErrors_ODataError,
      },
    ],
  },
  {
    method: 'get',
    path: '/me/messages/:messageId/attachments/:attachmentId',
    alias: 'get-mail-attachment',
    description: `Read the properties, relationships, or raw contents of an attachment that is attached to a user event, message, or group post. An attachment can be one of the following types: All these types of attachments are derived from the attachment resource.`,
    requestFormat: 'json',
    parameters: [
      {
        name: '$select',
        type: 'Query',
        schema: z.array(z.string()).describe('Select properties to be returned').optional(),
      },
      {
        name: '$expand',
        type: 'Query',
        schema: z.array(z.string()).describe('Expand related entities').optional(),
      },
    ],
    response: z.void(),
    errors: [
      {
        status: NaN,
        description: `Retrieved navigation property`,
        schema: microsoft_graph_attachment,
      },
      {
        status: NaN,
        description: `error`,
        schema: microsoft_graph_ODataErrors_ODataError,
      },
      {
        status: NaN,
        description: `error`,
        schema: microsoft_graph_ODataErrors_ODataError,
      },
    ],
  },
]);

export const api = new Zodios(endpoints);

export function createApiClient(baseUrl: string, options?: ZodiosOptions) {
  return new Zodios(baseUrl, endpoints, options);
}
