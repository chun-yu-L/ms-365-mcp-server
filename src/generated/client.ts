import { makeApi, Zodios, type ZodiosOptions } from './hack.js';
import { z } from 'zod';

type microsoft_graph_contentType = {
  id?: string | undefined;
  associatedHubsUrls?: Array<string | null> | undefined;
  description?: (string | null) | undefined;
  documentSet?: microsoft_graph_documentSet | undefined;
  documentTemplate?: microsoft_graph_documentSetContent | undefined;
  group?: (string | null) | undefined;
  hidden?: (boolean | null) | undefined;
  inheritedFrom?: microsoft_graph_itemReference | undefined;
  isBuiltIn?: (boolean | null) | undefined;
  name?: (string | null) | undefined;
  order?: microsoft_graph_contentTypeOrder | undefined;
  parentId?: (string | null) | undefined;
  propagateChanges?: (boolean | null) | undefined;
  readOnly?: (boolean | null) | undefined;
  sealed?: (boolean | null) | undefined;
  base?: microsoft_graph_contentType | undefined;
  baseTypes?: Array<microsoft_graph_contentType> | undefined;
  columnLinks?: Array<microsoft_graph_columnLink> | undefined;
  columnPositions?: Array<microsoft_graph_columnDefinition> | undefined;
  columns?: Array<microsoft_graph_columnDefinition> | undefined;
};
type microsoft_graph_documentSet = {
  allowedContentTypes?: Array<microsoft_graph_contentTypeInfo> | undefined;
  defaultContents?: Array<microsoft_graph_documentSetContent> | undefined;
  propagateWelcomePageChanges?: (boolean | null) | undefined;
  shouldPrefixNameToFile?: (boolean | null) | undefined;
  welcomePageUrl?: (string | null) | undefined;
  sharedColumns?: Array<microsoft_graph_columnDefinition> | undefined;
  welcomePageColumns?: Array<microsoft_graph_columnDefinition> | undefined;
};
type microsoft_graph_contentTypeInfo = {
  id?: (string | null) | undefined;
  name?: (string | null) | undefined;
};
type microsoft_graph_documentSetContent = {
  contentType?: microsoft_graph_contentTypeInfo | undefined;
  fileName?: (string | null) | undefined;
  folderName?: (string | null) | undefined;
};
type microsoft_graph_columnDefinition = {
  id?: string | undefined;
  name?: (string | null) | undefined;
  displayName?: (string | null) | undefined;
  description?: (string | null) | undefined;
  type?: microsoft_graph_columnTypes | undefined;
  boolean?: microsoft_graph_booleanColumn | undefined;
  calculated?: microsoft_graph_calculatedColumn | undefined;
  choice?: microsoft_graph_choiceColumn | undefined;
  columnGroup?: (string | null) | undefined;
  contentApprovalStatus?: microsoft_graph_contentApprovalStatusColumn | undefined;
  currency?: microsoft_graph_currencyColumn | undefined;
  dateTime?: microsoft_graph_dateTimeColumn | undefined;
  defaultValue?: microsoft_graph_defaultColumnValue | undefined;
  enforceUniqueValues?: (boolean | null) | undefined;
  geolocation?: microsoft_graph_geolocationColumn | undefined;
  hidden?: (boolean | null) | undefined;
  hyperlinkOrPicture?: microsoft_graph_hyperlinkOrPictureColumn | undefined;
  indexed?: (boolean | null) | undefined;
  isDeletable?: (boolean | null) | undefined;
  isReorderable?: (boolean | null) | undefined;
  isSealed?: (boolean | null) | undefined;
  lookup?: microsoft_graph_lookupColumn | undefined;
  number?: microsoft_graph_numberColumn | undefined;
  personOrGroup?: microsoft_graph_personOrGroupColumn | undefined;
  propagateChanges?: (boolean | null) | undefined;
} & {
  [key: string]: any;
};
type microsoft_graph_columnTypes =
  | 'note'
  | 'text'
  | 'choice'
  | 'multichoice'
  | 'number'
  | 'currency'
  | 'dateTime'
  | 'lookup'
  | 'boolean'
  | 'user'
  | 'url'
  | 'calculated'
  | 'location'
  | 'geolocation'
  | 'term'
  | 'multiterm'
  | 'thumbnail'
  | 'approvalStatus'
  | 'unknownFutureValue';
type microsoft_graph_booleanColumn = {};
type microsoft_graph_calculatedColumn = {
  format?: (string | null) | undefined;
  formula?: (string | null) | undefined;
  outputType?: (string | null) | undefined;
};
type microsoft_graph_choiceColumn = {
  allowTextEntry?: (boolean | null) | undefined;
  choices?: Array<string | null> | undefined;
  displayAs?: (string | null) | undefined;
};
type microsoft_graph_contentApprovalStatusColumn = {};
type microsoft_graph_currencyColumn = {
  locale?: (string | null) | undefined;
};
type microsoft_graph_dateTimeColumn = {
  displayAs?: (string | null) | undefined;
  format?: (string | null) | undefined;
};
type microsoft_graph_defaultColumnValue = {
  formula?: (string | null) | undefined;
  value?: (string | null) | undefined;
};
type microsoft_graph_geolocationColumn = {};
type microsoft_graph_hyperlinkOrPictureColumn = {
  isPicture?: (boolean | null) | undefined;
};
type microsoft_graph_lookupColumn = {
  allowMultipleValues?: (boolean | null) | undefined;
  allowUnlimitedLength?: (boolean | null) | undefined;
  columnName?: (string | null) | undefined;
  listId?: (string | null) | undefined;
  primaryLookupColumnId?: (string | null) | undefined;
};
type microsoft_graph_numberColumn = {
  decimalPlaces?: (string | null) | undefined;
  displayAs?: (string | null) | undefined;
  maximum?: (number | null) | undefined;
  minimum?: (number | null) | undefined;
};
type microsoft_graph_personOrGroupColumn = {
  allowMultipleSelection?: (boolean | null) | undefined;
  chooseFromType?: (string | null) | undefined;
  displayAs?: (string | null) | undefined;
};
type microsoft_graph_itemReference = {
  driveId?: (string | null) | undefined;
  driveType?: (string | null) | undefined;
  id?: (string | null) | undefined;
  name?: (string | null) | undefined;
  path?: (string | null) | undefined;
  shareId?: (string | null) | undefined;
  sharepointIds?: microsoft_graph_sharepointIds | undefined;
  siteId?: (string | null) | undefined;
};
type microsoft_graph_sharepointIds = {
  listId?: (string | null) | undefined;
  listItemId?: (string | null) | undefined;
  listItemUniqueId?: (string | null) | undefined;
  siteId?: (string | null) | undefined;
  siteUrl?: (string | null) | undefined;
  tenantId?: (string | null) | undefined;
  webId?: (string | null) | undefined;
};
type microsoft_graph_contentTypeOrder = {
  default?: (boolean | null) | undefined;
  position?: (number | null) | undefined;
};
type microsoft_graph_columnLink = {
  id?: string | undefined;
  name?: (string | null) | undefined;
};
type microsoft_graph_drive = {
  id?: string | undefined;
  createdBy?: microsoft_graph_identitySet | undefined;
  createdDateTime?: string | undefined;
  description?: (string | null) | undefined;
  eTag?: (string | null) | undefined;
  lastModifiedBy?: microsoft_graph_identitySet | undefined;
  lastModifiedDateTime?: string | undefined;
  name?: (string | null) | undefined;
  parentReference?: microsoft_graph_itemReference | undefined;
  webUrl?: (string | null) | undefined;
  createdByUser?: microsoft_graph_user | undefined;
  lastModifiedByUser?: microsoft_graph_user | undefined;
  driveType?: (string | null) | undefined;
  owner?: microsoft_graph_identitySet | undefined;
  quota?: microsoft_graph_quota | undefined;
  sharePointIds?: microsoft_graph_sharepointIds | undefined;
  system?: microsoft_graph_systemFacet | undefined;
  bundles?: Array<microsoft_graph_driveItem> | undefined;
  following?: Array<microsoft_graph_driveItem> | undefined;
  items?: Array<microsoft_graph_driveItem> | undefined;
  list?: microsoft_graph_list | undefined;
  root?: microsoft_graph_driveItem | undefined;
  special?: Array<microsoft_graph_driveItem> | undefined;
};
type microsoft_graph_identitySet = {
  application?: microsoft_graph_identity | undefined;
  device?: microsoft_graph_identity | undefined;
  user?: microsoft_graph_identity | undefined;
};
type microsoft_graph_identity = {
  displayName?: (string | null) | undefined;
  id?: (string | null) | undefined;
};
type microsoft_graph_user = {
  id?: string | undefined;
  displayName?: (string | null) | undefined;
  createdDateTime?: (string | null) | undefined;
  state?: (string | null) | undefined;
  userPrincipalName?: (string | null) | undefined;
  deletedDateTime?: (string | null) | undefined;
  aboutMe?: (string | null) | undefined;
  accountEnabled?: (boolean | null) | undefined;
  ageGroup?: (string | null) | undefined;
  assignedLicenses?: Array<microsoft_graph_assignedLicense> | undefined;
  assignedPlans?: Array<microsoft_graph_assignedPlan> | undefined;
  authorizationInfo?: microsoft_graph_authorizationInfo | undefined;
  birthday?: string | undefined;
  businessPhones?: Array<string> | undefined;
  city?: (string | null) | undefined;
  companyName?: (string | null) | undefined;
  consentProvidedForMinor?: (string | null) | undefined;
  country?: (string | null) | undefined;
  creationType?: (string | null) | undefined;
  customSecurityAttributes?: microsoft_graph_customSecurityAttributeValue | undefined;
  department?: (string | null) | undefined;
  deviceEnrollmentLimit?: number | undefined;
  employeeHireDate?: (string | null) | undefined;
  employeeId?: (string | null) | undefined;
  employeeLeaveDateTime?: (string | null) | undefined;
} & {
  [key: string]: any;
};
type microsoft_graph_assignedLicense = {
  disabledPlans?: Array<string> | undefined;
  skuId?: (string | null) | undefined;
};
type microsoft_graph_assignedPlan = {
  assignedDateTime?: (string | null) | undefined;
  capabilityStatus?: (string | null) | undefined;
  service?: (string | null) | undefined;
  servicePlanId?: (string | null) | undefined;
};
type microsoft_graph_authorizationInfo = {
  certificateUserIds?: Array<string | null> | undefined;
};
type microsoft_graph_customSecurityAttributeValue = {};
type microsoft_graph_quota = {
  deleted?: (number | null) | undefined;
  remaining?: (number | null) | undefined;
  state?: (string | null) | undefined;
  storagePlanInformation?: microsoft_graph_storagePlanInformation | undefined;
  total?: (number | null) | undefined;
  used?: (number | null) | undefined;
};
type microsoft_graph_storagePlanInformation = {
  upgradeAvailable?: (boolean | null) | undefined;
};
type microsoft_graph_systemFacet = {};
type microsoft_graph_driveItem = {
  id?: string | undefined;
  name?: (string | null) | undefined;
  description?: (string | null) | undefined;
  createdDateTime?: string | undefined;
  lastModifiedDateTime?: string | undefined;
  content?: (string | null) | undefined;
  createdBy?: microsoft_graph_identitySet | undefined;
  eTag?: (string | null) | undefined;
  lastModifiedBy?: microsoft_graph_identitySet | undefined;
  parentReference?: microsoft_graph_itemReference | undefined;
  webUrl?: (string | null) | undefined;
  createdByUser?: microsoft_graph_user | undefined;
  lastModifiedByUser?: microsoft_graph_user | undefined;
  audio?: microsoft_graph_audio | undefined;
  bundle?: microsoft_graph_bundle | undefined;
  cTag?: (string | null) | undefined;
  deleted?: microsoft_graph_deleted | undefined;
  file?: microsoft_graph_file | undefined;
  fileSystemInfo?: microsoft_graph_fileSystemInfo | undefined;
  folder?: microsoft_graph_folder | undefined;
  image?: microsoft_graph_image | undefined;
  location?: microsoft_graph_geoCoordinates | undefined;
  malware?: microsoft_graph_malware | undefined;
  package?: microsoft_graph_package | undefined;
  pendingOperations?: microsoft_graph_pendingOperations | undefined;
} & {
  [key: string]: any;
};
type microsoft_graph_audio = {
  album?: (string | null) | undefined;
  albumArtist?: (string | null) | undefined;
  artist?: (string | null) | undefined;
  bitrate?: (number | null) | undefined;
  composers?: (string | null) | undefined;
  copyright?: (string | null) | undefined;
  disc?: (number | null) | undefined;
  discCount?: (number | null) | undefined;
  duration?: (number | null) | undefined;
  genre?: (string | null) | undefined;
  hasDrm?: (boolean | null) | undefined;
  isVariableBitrate?: (boolean | null) | undefined;
  title?: (string | null) | undefined;
  track?: (number | null) | undefined;
  trackCount?: (number | null) | undefined;
  year?: (number | null) | undefined;
};
type microsoft_graph_bundle = {
  album?: microsoft_graph_album | undefined;
  childCount?: (number | null) | undefined;
};
type microsoft_graph_album = {
  coverImageItemId?: (string | null) | undefined;
};
type microsoft_graph_deleted = {
  state?: (string | null) | undefined;
};
type microsoft_graph_file = {
  hashes?: microsoft_graph_hashes | undefined;
  mimeType?: (string | null) | undefined;
  processingMetadata?: (boolean | null) | undefined;
};
type microsoft_graph_hashes = {
  crc32Hash?: (string | null) | undefined;
  quickXorHash?: (string | null) | undefined;
  sha1Hash?: (string | null) | undefined;
  sha256Hash?: (string | null) | undefined;
};
type microsoft_graph_fileSystemInfo = {
  createdDateTime?: (string | null) | undefined;
  lastAccessedDateTime?: (string | null) | undefined;
  lastModifiedDateTime?: (string | null) | undefined;
};
type microsoft_graph_folder = {
  childCount?: (number | null) | undefined;
  view?: microsoft_graph_folderView | undefined;
};
type microsoft_graph_folderView = {
  sortBy?: (string | null) | undefined;
  sortOrder?: (string | null) | undefined;
  viewType?: (string | null) | undefined;
};
type microsoft_graph_image = {
  height?: (number | null) | undefined;
  width?: (number | null) | undefined;
};
type microsoft_graph_geoCoordinates = {
  altitude?: (number | null) | undefined;
  latitude?: (number | null) | undefined;
  longitude?: (number | null) | undefined;
};
type microsoft_graph_malware = {
  description?: (string | null) | undefined;
};
type microsoft_graph_package = {
  type?: (string | null) | undefined;
};
type microsoft_graph_pendingOperations = {
  pendingContentUpdate?: microsoft_graph_pendingContentUpdate | undefined;
};
type microsoft_graph_pendingContentUpdate = {
  queuedDateTime?: (string | null) | undefined;
};
type microsoft_graph_listInfo = {
  contentTypesEnabled?: (boolean | null) | undefined;
  hidden?: (boolean | null) | undefined;
  template?: (string | null) | undefined;
};
type microsoft_graph_listItem = {
  id?: string | undefined;
  createdBy?: microsoft_graph_identitySet | undefined;
  createdDateTime?: string | undefined;
  description?: (string | null) | undefined;
  eTag?: (string | null) | undefined;
  lastModifiedBy?: microsoft_graph_identitySet | undefined;
  lastModifiedDateTime?: string | undefined;
  name?: (string | null) | undefined;
  parentReference?: microsoft_graph_itemReference | undefined;
  webUrl?: (string | null) | undefined;
  createdByUser?: microsoft_graph_user | undefined;
  lastModifiedByUser?: microsoft_graph_user | undefined;
  contentType?: microsoft_graph_contentTypeInfo | undefined;
  sharepointIds?: microsoft_graph_sharepointIds | undefined;
  analytics?: microsoft_graph_itemAnalytics | undefined;
  documentSetVersions?: Array<microsoft_graph_documentSetVersion> | undefined;
  driveItem?: microsoft_graph_driveItem | undefined;
  fields?: microsoft_graph_fieldValueSet | undefined;
  versions?: Array<microsoft_graph_listItemVersion> | undefined;
};
type microsoft_graph_itemAnalytics = {
  id?: string | undefined;
  allTime?: microsoft_graph_itemActivityStat | undefined;
  itemActivityStats?: Array<microsoft_graph_itemActivityStat> | undefined;
  lastSevenDays?: microsoft_graph_itemActivityStat | undefined;
};
type microsoft_graph_itemActivityStat = {
  id?: string | undefined;
  access?: microsoft_graph_itemActionStat | undefined;
  create?: microsoft_graph_itemActionStat | undefined;
  delete?: microsoft_graph_itemActionStat | undefined;
  edit?: microsoft_graph_itemActionStat | undefined;
  endDateTime?: (string | null) | undefined;
  incompleteData?: microsoft_graph_incompleteData | undefined;
  isTrending?: (boolean | null) | undefined;
  move?: microsoft_graph_itemActionStat | undefined;
  startDateTime?: (string | null) | undefined;
  activities?: Array<microsoft_graph_itemActivity> | undefined;
};
type microsoft_graph_itemActionStat = {
  actionCount?: (number | null) | undefined;
  actorCount?: (number | null) | undefined;
};
type microsoft_graph_incompleteData = {
  missingDataBeforeDateTime?: (string | null) | undefined;
  wasThrottled?: (boolean | null) | undefined;
};
type microsoft_graph_itemActivity = {
  id?: string | undefined;
  access?: microsoft_graph_accessAction | undefined;
  activityDateTime?: (string | null) | undefined;
  actor?: microsoft_graph_identitySet | undefined;
  driveItem?: microsoft_graph_driveItem | undefined;
};
type microsoft_graph_accessAction = {};
type microsoft_graph_documentSetVersion = {
  comment?: (string | null) | undefined;
  createdBy?: microsoft_graph_identitySet | undefined;
  createdDateTime?: (string | null) | undefined;
  items?: Array<microsoft_graph_documentSetVersionItem> | undefined;
  shouldCaptureMinorVersion?: (boolean | null) | undefined;
};
type microsoft_graph_documentSetVersionItem = {
  itemId?: (string | null) | undefined;
  title?: (string | null) | undefined;
  versionId?: (string | null) | undefined;
};
type microsoft_graph_fieldValueSet = {
  id?: string | undefined;
};
type microsoft_graph_listItemVersion = {
  id?: string | undefined;
  lastModifiedBy?: microsoft_graph_identitySet | undefined;
  lastModifiedDateTime?: (string | null) | undefined;
  publication?: microsoft_graph_publicationFacet | undefined;
  fields?: microsoft_graph_fieldValueSet | undefined;
};
type microsoft_graph_publicationFacet = {
  checkedOutBy?: microsoft_graph_identitySet | undefined;
  level?: (string | null) | undefined;
  versionId?: (string | null) | undefined;
};
type microsoft_graph_richLongRunningOperation = {
  id?: string | undefined;
  createdDateTime?: (string | null) | undefined;
  lastActionDateTime?: (string | null) | undefined;
  resourceLocation?: (string | null) | undefined;
  status?: microsoft_graph_longRunningOperationStatus | undefined;
  statusDetail?: (string | null) | undefined;
  error?: microsoft_graph_publicError | undefined;
  percentageComplete?: (number | null) | undefined;
  resourceId?: (string | null) | undefined;
  type?: (string | null) | undefined;
};
type microsoft_graph_longRunningOperationStatus =
  | 'notStarted'
  | 'running'
  | 'succeeded'
  | 'failed'
  | 'unknownFutureValue';
type microsoft_graph_publicError = {
  code?: (string | null) | undefined;
  details?: Array<microsoft_graph_publicErrorDetail> | undefined;
  innerError?: microsoft_graph_publicInnerError | undefined;
  message?: (string | null) | undefined;
  target?: (string | null) | undefined;
};
type microsoft_graph_publicErrorDetail = {
  code?: (string | null) | undefined;
  message?: (string | null) | undefined;
  target?: (string | null) | undefined;
};
type microsoft_graph_publicInnerError = {
  code?: (string | null) | undefined;
  details?: Array<microsoft_graph_publicErrorDetail> | undefined;
  message?: (string | null) | undefined;
  target?: (string | null) | undefined;
};
type microsoft_graph_subscription = {
  id?: string | undefined;
  applicationId?: (string | null) | undefined;
  changeType?: string | undefined;
  clientState?: (string | null) | undefined;
  creatorId?: (string | null) | undefined;
  encryptionCertificate?: (string | null) | undefined;
  encryptionCertificateId?: (string | null) | undefined;
  expirationDateTime?: string | undefined;
  includeResourceData?: (boolean | null) | undefined;
  latestSupportedTlsVersion?: (string | null) | undefined;
  lifecycleNotificationUrl?: (string | null) | undefined;
  notificationQueryOptions?: (string | null) | undefined;
  notificationUrl?: string | undefined;
  notificationUrlAppId?: (string | null) | undefined;
  resource?: string | undefined;
};
type microsoft_graph_list = {
  id?: string | undefined;
  createdBy?: microsoft_graph_identitySet | undefined;
  createdDateTime?: string | undefined;
  description?: (string | null) | undefined;
  eTag?: (string | null) | undefined;
  lastModifiedBy?: microsoft_graph_identitySet | undefined;
  lastModifiedDateTime?: string | undefined;
  name?: (string | null) | undefined;
  parentReference?: microsoft_graph_itemReference | undefined;
  webUrl?: (string | null) | undefined;
  createdByUser?: microsoft_graph_user | undefined;
  lastModifiedByUser?: microsoft_graph_user | undefined;
  displayName?: (string | null) | undefined;
  list?: microsoft_graph_listInfo | undefined;
  sharepointIds?: microsoft_graph_sharepointIds | undefined;
  system?: microsoft_graph_systemFacet | undefined;
  columns?: Array<microsoft_graph_columnDefinition> | undefined;
  contentTypes?: Array<microsoft_graph_contentType> | undefined;
  drive?: microsoft_graph_drive | undefined;
  items?: Array<microsoft_graph_listItem> | undefined;
  operations?: Array<microsoft_graph_richLongRunningOperation> | undefined;
  subscriptions?: Array<microsoft_graph_subscription> | undefined;
};
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
const microsoft_graph_publicErrorDetail = z
  .object({
    code: z.string().describe('The error code.').nullish(),
    message: z.string().describe('The error message.').nullish(),
    target: z.string().describe('The target of the error.').nullish(),
  })
  .strict();
const microsoft_graph_publicInnerError = z
  .object({
    code: z.string().describe('The error code.').nullish(),
    details: z
      .array(microsoft_graph_publicErrorDetail)
      .describe('A collection of error details.')
      .optional(),
    message: z.string().describe('The error message.').nullish(),
    target: z.string().describe('The target of the error.').nullish(),
  })
  .strict();
const microsoft_graph_publicError = z
  .object({
    code: z.string().describe('Represents the error code.').nullish(),
    details: z
      .array(microsoft_graph_publicErrorDetail)
      .describe('Details of the error.')
      .optional(),
    innerError: microsoft_graph_publicInnerError.optional(),
    message: z.string().describe('A non-localized message for the developer.').nullish(),
    target: z.string().describe('The target of the error.').nullish(),
  })
  .strict();
const microsoft_graph_identity = z
  .object({
    displayName: z
      .string()
      .describe(
        "The display name of the identity.For drive items, the display name might not always be available or up to date. For example, if a user changes their display name the API might show the new value in a future response, but the items associated with the user don't show up as changed when using delta."
      )
      .nullish(),
    id: z
      .string()
      .describe(
        "Unique identifier for the identity or actor. For example, in the access reviews decisions API, this property might record the id of the principal, that is, the group, user, or application that's subject to review."
      )
      .nullish(),
  })
  .strict();
const microsoft_graph_identitySet = z
  .object({
    application: microsoft_graph_identity.optional(),
    device: microsoft_graph_identity.optional(),
    user: microsoft_graph_identity.optional(),
  })
  .strict();
const microsoft_graph_sharepointIds = z
  .object({
    listId: z
      .string()
      .describe("The unique identifier (guid) for the item's list in SharePoint.")
      .nullish(),
    listItemId: z
      .string()
      .describe('An integer identifier for the item within the containing list.')
      .nullish(),
    listItemUniqueId: z
      .string()
      .describe(
        'The unique identifier (guid) for the item within OneDrive for Business or a SharePoint site.'
      )
      .nullish(),
    siteId: z
      .string()
      .describe("The unique identifier (guid) for the item's site collection (SPSite).")
      .nullish(),
    siteUrl: z
      .string()
      .describe('The SharePoint URL for the site that contains the item.')
      .nullish(),
    tenantId: z.string().describe('The unique identifier (guid) for the tenancy.').nullish(),
    webId: z
      .string()
      .describe("The unique identifier (guid) for the item's site (SPWeb).")
      .nullish(),
  })
  .strict();
const microsoft_graph_itemReference = z
  .object({
    driveId: z
      .string()
      .describe(
        'Unique identifier of the drive instance that contains the driveItem. Only returned if the item is located in a drive. Read-only.'
      )
      .nullish(),
    driveType: z
      .string()
      .describe(
        'Identifies the type of drive. Only returned if the item is located in a drive. See drive resource for values.'
      )
      .nullish(),
    id: z
      .string()
      .describe(
        'Unique identifier of the driveItem in the drive or a listItem in a list. Read-only.'
      )
      .nullish(),
    name: z.string().describe('The name of the item being referenced. Read-only.').nullish(),
    path: z
      .string()
      .describe('Percent-encoded path that can be used to navigate to the item. Read-only.')
      .nullish(),
    shareId: z
      .string()
      .describe(
        'A unique identifier for a shared resource that can be accessed via the Shares API.'
      )
      .nullish(),
    sharepointIds: microsoft_graph_sharepointIds.optional(),
    siteId: z
      .string()
      .describe(
        'For OneDrive for Business and SharePoint, this property represents the ID of the site that contains the parent document library of the driveItem resource or the parent list of the listItem resource. The value is the same as the id property of that site resource. It is an opaque string that consists of three identifiers of the site. For OneDrive, this property is not populated.'
      )
      .nullish(),
  })
  .strict();
const microsoft_graph_assignedLicense = z
  .object({
    disabledPlans: z
      .array(
        z
          .string()
          .regex(/^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/)
          .uuid()
      )
      .describe(
        "A collection of the unique identifiers for plans that have been disabled. IDs are available in servicePlans > servicePlanId in the tenant's subscribedSkus or serviceStatus > servicePlanId in the tenant's companySubscription."
      )
      .optional(),
    skuId: z
      .string()
      .regex(/^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/)
      .uuid()
      .describe(
        'The unique identifier for the SKU. Corresponds to the skuId from subscribedSkus or companySubscription.'
      )
      .nullish(),
  })
  .strict();
const microsoft_graph_assignedPlan = z
  .object({
    assignedDateTime: z
      .string()
      .regex(
        /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
      )
      .datetime({ offset: true })
      .describe(
        'The date and time at which the plan was assigned. The Timestamp type represents date and time information using ISO 8601 format and is always in UTC time. For example, midnight UTC on Jan 1, 2014 is 2014-01-01T00:00:00Z.'
      )
      .nullish(),
    capabilityStatus: z
      .string()
      .describe(
        'Condition of the capability assignment. The possible values are Enabled, Warning, Suspended, Deleted, LockedOut. See a detailed description of each value.'
      )
      .nullish(),
    service: z.string().describe('The name of the service; for example, exchange.').nullish(),
    servicePlanId: z
      .string()
      .regex(/^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/)
      .uuid()
      .describe(
        'A GUID that identifies the service plan. For a complete list of GUIDs and their equivalent friendly service names, see Product names and service plan identifiers for licensing.'
      )
      .nullish(),
  })
  .strict();
const microsoft_graph_authorizationInfo = z
  .object({ certificateUserIds: z.array(z.string().nullable()).optional() })
  .strict();
const microsoft_graph_customSecurityAttributeValue = z.object({}).strict();
const microsoft_graph_user = z
  .object({
    id: z.string().describe('The unique identifier for an entity. Read-only.').optional(),
    displayName: z
      .string()
      .describe(
        "The name displayed in the address book for the user. This value is usually the combination of the user's first name, middle initial, and family name. This property is required when a user is created and it can't be cleared during updates. Maximum length is 256 characters. Returned by default. Supports $filter (eq, ne, not , ge, le, in, startsWith, and eq on null values), $orderby, and $search."
      )
      .nullish(),
    createdDateTime: z
      .string()
      .regex(
        /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
      )
      .datetime({ offset: true })
      .describe(
        "The date and time the user was created, in ISO 8601 format and UTC. The value can't be modified and is automatically populated when the entity is created. Nullable. For on-premises users, the value represents when they were first created in Microsoft Entra ID. Property is null for some users created before June 2018 and on-premises users that were synced to Microsoft Entra ID before June 2018. Read-only. Returned only on $select. Supports $filter (eq, ne, not , ge, le, in)."
      )
      .nullish(),
    state: z
      .string()
      .describe(
        "The state or province in the user's address. Maximum length is 128 characters. Returned only on $select. Supports $filter (eq, ne, not, ge, le, in, startsWith, and eq on null values)."
      )
      .nullish(),
    userPrincipalName: z
      .string()
      .describe(
        "The user principal name (UPN) of the user. The UPN is an Internet-style sign-in name for the user based on the Internet standard RFC 822. By convention, this value should map to the user's email name. The general format is alias@domain, where the domain must be present in the tenant's collection of verified domains. This property is required when a user is created. The verified domains for the tenant can be accessed from the verifiedDomains property of organization.NOTE: This property can't contain accent characters. Only the following characters are allowed A - Z, a - z, 0 - 9, ' . - _ ! # ^ ~. For the complete list of allowed characters, see username policies. Returned by default. Supports $filter (eq, ne, not, ge, le, in, startsWith, endsWith) and $orderby."
      )
      .nullish(),
    deletedDateTime: z
      .string()
      .regex(
        /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
      )
      .datetime({ offset: true })
      .describe(
        "Date and time when this object was deleted. Always null when the object hasn't been deleted."
      )
      .nullish(),
    aboutMe: z
      .string()
      .describe(
        'A freeform text entry field for the user to describe themselves. Returned only on $select.'
      )
      .nullish(),
    accountEnabled: z
      .boolean()
      .describe(
        'true if the account is enabled; otherwise, false. This property is required when a user is created. Returned only on $select. Supports $filter (eq, ne, not, and in).'
      )
      .nullish(),
    ageGroup: z
      .string()
      .describe(
        'Sets the age group of the user. Allowed values: null, Minor, NotAdult, and Adult. For more information, see legal age group property definitions. Returned only on $select. Supports $filter (eq, ne, not, and in).'
      )
      .nullish(),
    assignedLicenses: z
      .array(microsoft_graph_assignedLicense)
      .describe(
        "The licenses that are assigned to the user, including inherited (group-based) licenses. This property doesn't differentiate between directly assigned and inherited licenses. Use the licenseAssignmentStates property to identify the directly assigned and inherited licenses. Not nullable. Returned only on $select. Supports $filter (eq, not, /$count eq 0, /$count ne 0)."
      )
      .optional(),
    assignedPlans: z
      .array(microsoft_graph_assignedPlan)
      .describe(
        'The plans that are assigned to the user. Read-only. Not nullable. Returned only on $select. Supports $filter (eq and not).'
      )
      .optional(),
    authorizationInfo: microsoft_graph_authorizationInfo.optional(),
    birthday: z
      .string()
      .regex(
        /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
      )
      .datetime({ offset: true })
      .describe(
        'The birthday of the user. The Timestamp type represents date and time information using ISO 8601 format and is always in UTC. For example, midnight UTC on Jan 1, 2014, is 2014-01-01T00:00:00Z. Returned only on $select.'
      )
      .optional(),
    businessPhones: z
      .array(z.string())
      .describe(
        "The telephone numbers for the user. NOTE: Although it's a string collection, only one number can be set for this property. Read-only for users synced from the on-premises directory. Returned by default. Supports $filter (eq, not, ge, le, startsWith)."
      )
      .optional(),
    city: z
      .string()
      .describe(
        'The city where the user is located. Maximum length is 128 characters. Returned only on $select. Supports $filter (eq, ne, not, ge, le, in, startsWith, and eq on null values).'
      )
      .nullish(),
    companyName: z
      .string()
      .describe(
        'The name of the company that the user is associated with. This property can be useful for describing the company that a guest comes from. The maximum length is 64 characters.Returned only on $select. Supports $filter (eq, ne, not, ge, le, in, startsWith, and eq on null values).'
      )
      .nullish(),
    consentProvidedForMinor: z
      .string()
      .describe(
        'Sets whether consent was obtained for minors. Allowed values: null, Granted, Denied, and NotRequired. For more information, see legal age group property definitions. Returned only on $select. Supports $filter (eq, ne, not, and in).'
      )
      .nullish(),
    country: z
      .string()
      .describe(
        'The country or region where the user is located; for example, US or UK. Maximum length is 128 characters. Returned only on $select. Supports $filter (eq, ne, not, ge, le, in, startsWith, and eq on null values).'
      )
      .nullish(),
    creationType: z
      .string()
      .describe(
        'Indicates whether the user account was created through one of the following methods:  As a regular school or work account (null). As an external account (Invitation). As a local account for an Azure Active Directory B2C tenant (LocalAccount). Through self-service sign-up by an internal user using email verification (EmailVerified). Through self-service sign-up by a guest signing up through a link that is part of a user flow (SelfServiceSignUp). Read-only.Returned only on $select. Supports $filter (eq, ne, not, in).'
      )
      .nullish(),
    customSecurityAttributes: microsoft_graph_customSecurityAttributeValue.optional(),
    department: z
      .string()
      .describe(
        'The name of the department in which the user works. Maximum length is 64 characters. Returned only on $select. Supports $filter (eq, ne, not , ge, le, in, and eq on null values).'
      )
      .nullish(),
    deviceEnrollmentLimit: z
      .number()
      .gte(-2147483648)
      .lte(2147483647)
      .describe(
        'The limit on the maximum number of devices that the user is permitted to enroll. Allowed values are 5 or 1000.'
      )
      .optional(),
    employeeHireDate: z
      .string()
      .regex(
        /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
      )
      .datetime({ offset: true })
      .describe(
        'The date and time when the user was hired or will start work in a future hire. Returned only on $select. Supports $filter (eq, ne, not , ge, le, in).'
      )
      .nullish(),
    employeeId: z
      .string()
      .describe(
        'The employee identifier assigned to the user by the organization. The maximum length is 16 characters. Returned only on $select. Supports $filter (eq, ne, not , ge, le, in, startsWith, and eq on null values).'
      )
      .nullish(),
    employeeLeaveDateTime: z
      .string()
      .regex(
        /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
      )
      .datetime({ offset: true })
      .describe(
        'The date and time when the user left or will leave the organization. To read this property, the calling app must be assigned the User-LifeCycleInfo.Read.All permission. To write this property, the calling app must be assigned the User.Read.All and User-LifeCycleInfo.ReadWrite.All permissions. To read this property in delegated scenarios, the admin needs at least one of the following Microsoft Entra roles: Lifecycle Workflows Administrator (least privilege), Global Reader. To write this property in delegated scenarios, the admin needs the Global Administrator role. Supports $filter (eq, ne, not , ge, le, in). For more information, see Configure the employeeLeaveDateTime property for a user.'
      )
      .nullish(),
  })
  .strict()
  .passthrough();
const microsoft_graph_root = z.object({}).strict();
const microsoft_graph_siteArchiveStatus = z.enum([
  'recentlyArchived',
  'fullyArchived',
  'reactivating',
  'unknownFutureValue',
]);
const microsoft_graph_siteArchivalDetails = z
  .object({ archiveStatus: microsoft_graph_siteArchiveStatus.optional() })
  .strict();
const microsoft_graph_siteCollection = z
  .object({
    archivalDetails: microsoft_graph_siteArchivalDetails.optional(),
    dataLocationCode: z
      .string()
      .describe(
        'The geographic region code for where this site collection resides. Only present for multi-geo tenants. Read-only.'
      )
      .nullish(),
    hostname: z.string().describe('The hostname for the site collection. Read-only.').nullish(),
    root: microsoft_graph_root.optional(),
  })
  .strict();
const microsoft_graph_itemActionStat = z
  .object({
    actionCount: z
      .number()
      .gte(-2147483648)
      .lte(2147483647)
      .describe('The number of times the action took place. Read-only.')
      .nullish(),
    actorCount: z
      .number()
      .gte(-2147483648)
      .lte(2147483647)
      .describe('The number of distinct actors that performed the action. Read-only.')
      .nullish(),
  })
  .strict();
const microsoft_graph_incompleteData = z
  .object({
    missingDataBeforeDateTime: z
      .string()
      .regex(
        /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
      )
      .datetime({ offset: true })
      .describe('The service does not have source data before the specified time.')
      .nullish(),
    wasThrottled: z
      .boolean()
      .describe('Some data was not recorded due to excessive activity.')
      .nullish(),
  })
  .strict();
const microsoft_graph_accessAction = z.object({}).strict();
const microsoft_graph_audio = z
  .object({
    album: z.string().describe('The title of the album for this audio file.').nullish(),
    albumArtist: z.string().describe('The artist named on the album for the audio file.').nullish(),
    artist: z.string().describe('The performing artist for the audio file.').nullish(),
    bitrate: z.number().describe('Bitrate expressed in kbps.').nullish(),
    composers: z.string().describe('The name of the composer of the audio file.').nullish(),
    copyright: z.string().describe('Copyright information for the audio file.').nullish(),
    disc: z
      .number()
      .gte(-32768)
      .lte(32767)
      .describe('The number of the disc this audio file came from.')
      .nullish(),
    discCount: z
      .number()
      .gte(-32768)
      .lte(32767)
      .describe('The total number of discs in this album.')
      .nullish(),
    duration: z
      .number()
      .describe('Duration of the audio file, expressed in milliseconds')
      .nullish(),
    genre: z.string().describe('The genre of this audio file.').nullish(),
    hasDrm: z
      .boolean()
      .describe('Indicates if the file is protected with digital rights management.')
      .nullish(),
    isVariableBitrate: z
      .boolean()
      .describe('Indicates if the file is encoded with a variable bitrate.')
      .nullish(),
    title: z.string().describe('The title of the audio file.').nullish(),
    track: z
      .number()
      .gte(-2147483648)
      .lte(2147483647)
      .describe('The number of the track on the original disc for this audio file.')
      .nullish(),
    trackCount: z
      .number()
      .gte(-2147483648)
      .lte(2147483647)
      .describe('The total number of tracks on the original disc for this audio file.')
      .nullish(),
    year: z
      .number()
      .gte(-2147483648)
      .lte(2147483647)
      .describe('The year the audio file was recorded.')
      .nullish(),
  })
  .strict();
const microsoft_graph_album = z
  .object({
    coverImageItemId: z
      .string()
      .describe('Unique identifier of the driveItem that is the cover of the album.')
      .nullish(),
  })
  .strict();
const microsoft_graph_bundle = z
  .object({
    album: microsoft_graph_album.optional(),
    childCount: z
      .number()
      .gte(-2147483648)
      .lte(2147483647)
      .describe('Number of children contained immediately within this container.')
      .nullish(),
  })
  .strict();
const microsoft_graph_deleted = z
  .object({ state: z.string().describe('Represents the state of the deleted item.').nullish() })
  .strict();
const microsoft_graph_hashes = z
  .object({
    crc32Hash: z
      .string()
      .describe('The CRC32 value of the file (if available). Read-only.')
      .nullish(),
    quickXorHash: z
      .string()
      .describe(
        'A proprietary hash of the file that can be used to determine if the contents of the file change (if available). Read-only.'
      )
      .nullish(),
    sha1Hash: z
      .string()
      .describe('SHA1 hash for the contents of the file (if available). Read-only.')
      .nullish(),
    sha256Hash: z.string().describe("This property isn't supported. Don't use.").nullish(),
  })
  .strict();
const microsoft_graph_file = z
  .object({
    hashes: microsoft_graph_hashes.optional(),
    mimeType: z
      .string()
      .describe(
        'The MIME type for the file. This is determined by logic on the server and might not be the value provided when the file was uploaded. Read-only.'
      )
      .nullish(),
    processingMetadata: z.boolean().nullish(),
  })
  .strict();
const microsoft_graph_fileSystemInfo = z
  .object({
    createdDateTime: z
      .string()
      .regex(
        /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
      )
      .datetime({ offset: true })
      .describe('The UTC date and time the file was created on a client.')
      .nullish(),
    lastAccessedDateTime: z
      .string()
      .regex(
        /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
      )
      .datetime({ offset: true })
      .describe(
        'The UTC date and time the file was last accessed. Available for the recent file list only.'
      )
      .nullish(),
    lastModifiedDateTime: z
      .string()
      .regex(
        /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
      )
      .datetime({ offset: true })
      .describe('The UTC date and time the file was last modified on a client.')
      .nullish(),
  })
  .strict();
const microsoft_graph_folderView = z
  .object({
    sortBy: z.string().describe('The method by which the folder should be sorted.').nullish(),
    sortOrder: z
      .string()
      .describe(
        'If true, indicates that items should be sorted in descending order. Otherwise, items should be sorted ascending.'
      )
      .nullish(),
    viewType: z
      .string()
      .describe('The type of view that should be used to represent the folder.')
      .nullish(),
  })
  .strict();
const microsoft_graph_folder = z
  .object({
    childCount: z
      .number()
      .gte(-2147483648)
      .lte(2147483647)
      .describe('Number of children contained immediately within this container.')
      .nullish(),
    view: microsoft_graph_folderView.optional(),
  })
  .strict();
const microsoft_graph_image = z
  .object({
    height: z
      .number()
      .gte(-2147483648)
      .lte(2147483647)
      .describe('Optional. Height of the image, in pixels. Read-only.')
      .nullish(),
    width: z
      .number()
      .gte(-2147483648)
      .lte(2147483647)
      .describe('Optional. Width of the image, in pixels. Read-only.')
      .nullish(),
  })
  .strict();
const microsoft_graph_geoCoordinates = z
  .object({
    altitude: z
      .number()
      .describe(
        'Optional. The altitude (height), in feet,  above sea level for the item. Read-only. [Simplified from 3 options]'
      )
      .nullish(),
    latitude: z
      .number()
      .describe(
        'Optional. The latitude, in decimal, for the item. Read-only. [Simplified from 3 options]'
      )
      .nullish(),
    longitude: z
      .number()
      .describe(
        'Optional. The longitude, in decimal, for the item. Read-only. [Simplified from 3 options]'
      )
      .nullish(),
  })
  .strict();
const microsoft_graph_malware = z
  .object({
    description: z.string().describe('Contains the virus details for the malware facet.').nullish(),
  })
  .strict();
const microsoft_graph_package = z
  .object({
    type: z
      .string()
      .describe(
        'A string indicating the type of package. While oneNote is the only currently defined value, you should expect other package types to be returned and handle them accordingly.'
      )
      .nullish(),
  })
  .strict();
const microsoft_graph_pendingContentUpdate = z
  .object({
    queuedDateTime: z
      .string()
      .regex(
        /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
      )
      .datetime({ offset: true })
      .describe('Date and time the pending binary operation was queued in UTC time. Read-only.')
      .nullish(),
  })
  .strict();
const microsoft_graph_pendingOperations = z
  .object({ pendingContentUpdate: microsoft_graph_pendingContentUpdate.optional() })
  .strict();
const microsoft_graph_driveItem = z
  .object({
    id: z.string().describe('The unique identifier for an entity. Read-only.').optional(),
    name: z.string().describe('The name of the item. Read-write.').nullish(),
    description: z
      .string()
      .describe('Provides a user-visible description of the item. Optional.')
      .nullish(),
    createdDateTime: z
      .string()
      .regex(
        /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
      )
      .datetime({ offset: true })
      .describe('Date and time of item creation. Read-only.')
      .optional(),
    lastModifiedDateTime: z
      .string()
      .regex(
        /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
      )
      .datetime({ offset: true })
      .describe('Date and time the item was last modified. Read-only.')
      .optional(),
    content: z.string().describe('The content stream, if the item represents a file.').nullish(),
    createdBy: microsoft_graph_identitySet.optional(),
    eTag: z.string().describe('ETag for the item. Read-only.').nullish(),
    lastModifiedBy: microsoft_graph_identitySet.optional(),
    parentReference: microsoft_graph_itemReference.optional(),
    webUrl: z
      .string()
      .describe(
        'URL that either displays the resource in the browser (for Office file formats), or is a direct link to the file (for other formats). Read-only.'
      )
      .nullish(),
    createdByUser: microsoft_graph_user
      .describe('[Note: Simplified from 132 properties to 25 most common ones]')
      .optional(),
    lastModifiedByUser: microsoft_graph_user
      .describe('[Note: Simplified from 132 properties to 25 most common ones]')
      .optional(),
    audio: microsoft_graph_audio.optional(),
    bundle: microsoft_graph_bundle.optional(),
    cTag: z
      .string()
      .describe(
        "An eTag for the content of the item. This eTag isn't changed if only the metadata is changed. Note This property isn't returned if the item is a folder. Read-only."
      )
      .nullish(),
    deleted: microsoft_graph_deleted.optional(),
    file: microsoft_graph_file.optional(),
    fileSystemInfo: microsoft_graph_fileSystemInfo.optional(),
    folder: microsoft_graph_folder.optional(),
    image: microsoft_graph_image.optional(),
    location: microsoft_graph_geoCoordinates.optional(),
    malware: microsoft_graph_malware.optional(),
    package: microsoft_graph_package.optional(),
    pendingOperations: microsoft_graph_pendingOperations.optional(),
  })
  .strict()
  .passthrough();
const microsoft_graph_itemActivity = z
  .object({
    id: z.string().describe('The unique identifier for an entity. Read-only.').optional(),
    access: microsoft_graph_accessAction.optional(),
    activityDateTime: z
      .string()
      .regex(
        /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
      )
      .datetime({ offset: true })
      .describe('Details about when the activity took place. Read-only.')
      .nullish(),
    actor: microsoft_graph_identitySet.optional(),
    driveItem: microsoft_graph_driveItem
      .describe('[Note: Simplified from 45 properties to 25 most common ones]')
      .optional(),
  })
  .strict();
const microsoft_graph_itemActivityStat = z
  .object({
    id: z.string().describe('The unique identifier for an entity. Read-only.').optional(),
    access: microsoft_graph_itemActionStat.optional(),
    create: microsoft_graph_itemActionStat.optional(),
    delete: microsoft_graph_itemActionStat.optional(),
    edit: microsoft_graph_itemActionStat.optional(),
    endDateTime: z
      .string()
      .regex(
        /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
      )
      .datetime({ offset: true })
      .describe('When the interval ends. Read-only.')
      .nullish(),
    incompleteData: microsoft_graph_incompleteData.optional(),
    isTrending: z
      .boolean()
      .describe("Indicates whether the item is 'trending.' Read-only.")
      .nullish(),
    move: microsoft_graph_itemActionStat.optional(),
    startDateTime: z
      .string()
      .regex(
        /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
      )
      .datetime({ offset: true })
      .describe('When the interval starts. Read-only.')
      .nullish(),
    activities: z
      .array(microsoft_graph_itemActivity)
      .describe('Exposes the itemActivities represented in this itemActivityStat resource.')
      .optional(),
  })
  .strict();
const microsoft_graph_itemAnalytics = z
  .object({
    id: z.string().describe('The unique identifier for an entity. Read-only.').optional(),
    allTime: microsoft_graph_itemActivityStat.optional(),
    itemActivityStats: z.array(microsoft_graph_itemActivityStat).optional(),
    lastSevenDays: microsoft_graph_itemActivityStat.optional(),
  })
  .strict();
const microsoft_graph_columnTypes = z.enum([
  'note',
  'text',
  'choice',
  'multichoice',
  'number',
  'currency',
  'dateTime',
  'lookup',
  'boolean',
  'user',
  'url',
  'calculated',
  'location',
  'geolocation',
  'term',
  'multiterm',
  'thumbnail',
  'approvalStatus',
  'unknownFutureValue',
]);
const microsoft_graph_booleanColumn = z.object({}).strict();
const microsoft_graph_calculatedColumn = z
  .object({
    format: z
      .string()
      .describe(
        'For dateTime output types, the format of the value. Possible values are: dateOnly or dateTime.'
      )
      .nullish(),
    formula: z
      .string()
      .describe('The formula used to compute the value for this column.')
      .nullish(),
    outputType: z
      .string()
      .describe(
        'The output type used to format values in this column. Possible values are: boolean, currency, dateTime, number, or text.'
      )
      .nullish(),
  })
  .strict();
const microsoft_graph_choiceColumn = z
  .object({
    allowTextEntry: z
      .boolean()
      .describe("If true, allows custom values that aren't in the configured choices.")
      .nullish(),
    choices: z
      .array(z.string().nullable())
      .describe('The list of values available for this column.')
      .optional(),
    displayAs: z
      .string()
      .describe(
        'How the choices are to be presented in the UX. Must be one of checkBoxes, dropDownMenu, or radioButtons'
      )
      .nullish(),
  })
  .strict();
const microsoft_graph_contentApprovalStatusColumn = z.object({}).strict();
const microsoft_graph_currencyColumn = z
  .object({
    locale: z
      .string()
      .describe('Specifies the locale from which to infer the currency symbol.')
      .nullish(),
  })
  .strict();
const microsoft_graph_dateTimeColumn = z
  .object({
    displayAs: z
      .string()
      .describe(
        'How the value should be presented in the UX. Must be one of default, friendly, or standard. See below for more details. If unspecified, treated as default.'
      )
      .nullish(),
    format: z
      .string()
      .describe(
        'Indicates whether the value should be presented as a date only or a date and time. Must be one of dateOnly or dateTime'
      )
      .nullish(),
  })
  .strict();
const microsoft_graph_defaultColumnValue = z
  .object({
    formula: z
      .string()
      .describe('The formula used to compute the default value for the column.')
      .nullish(),
    value: z
      .string()
      .describe('The direct value to use as the default value for the column.')
      .nullish(),
  })
  .strict();
const microsoft_graph_geolocationColumn = z.object({}).strict();
const microsoft_graph_hyperlinkOrPictureColumn = z
  .object({
    isPicture: z
      .boolean()
      .describe(
        'Specifies whether the display format used for URL columns is an image or a hyperlink.'
      )
      .nullish(),
  })
  .strict();
const microsoft_graph_lookupColumn = z
  .object({
    allowMultipleValues: z
      .boolean()
      .describe('Indicates whether multiple values can be selected from the source.')
      .nullish(),
    allowUnlimitedLength: z
      .boolean()
      .describe(
        'Indicates whether values in the column should be able to exceed the standard limit of 255 characters.'
      )
      .nullish(),
    columnName: z.string().describe('The name of the lookup source column.').nullish(),
    listId: z.string().describe('The unique identifier of the lookup source list.').nullish(),
    primaryLookupColumnId: z
      .string()
      .describe(
        'If specified, this column is a secondary lookup, pulling an additional field from the list item looked up by the primary lookup. Use the list item looked up by the primary as the source for the column named here.'
      )
      .nullish(),
  })
  .strict();
const microsoft_graph_numberColumn = z
  .object({
    decimalPlaces: z
      .string()
      .describe(
        'How many decimal places to display. See below for information about the possible values.'
      )
      .nullish(),
    displayAs: z
      .string()
      .describe(
        'How the value should be presented in the UX. Must be one of number or percentage. If unspecified, treated as number.'
      )
      .nullish(),
    maximum: z
      .number()
      .describe('The maximum permitted value. [Simplified from 3 options]')
      .nullish(),
    minimum: z
      .number()
      .describe('The minimum permitted value. [Simplified from 3 options]')
      .nullish(),
  })
  .strict();
const microsoft_graph_personOrGroupColumn = z
  .object({
    allowMultipleSelection: z
      .boolean()
      .describe('Indicates whether multiple values can be selected from the source.')
      .nullish(),
    chooseFromType: z
      .string()
      .describe(
        'Whether to allow selection of people only, or people and groups. Must be one of peopleAndGroups or peopleOnly.'
      )
      .nullish(),
    displayAs: z
      .string()
      .describe('How to display the information about the person or group chosen. See below.')
      .nullish(),
  })
  .strict();
const microsoft_graph_columnDefinition = z
  .object({
    id: z.string().describe('The unique identifier for an entity. Read-only.').optional(),
    name: z
      .string()
      .describe(
        'The API-facing name of the column as it appears in the fields on a listItem. For the user-facing name, see displayName.'
      )
      .nullish(),
    displayName: z.string().describe('The user-facing name of the column.').nullish(),
    description: z.string().describe('The user-facing description of the column.').nullish(),
    type: microsoft_graph_columnTypes.optional(),
    boolean: microsoft_graph_booleanColumn.optional(),
    calculated: microsoft_graph_calculatedColumn.optional(),
    choice: microsoft_graph_choiceColumn.optional(),
    columnGroup: z
      .string()
      .describe(
        'For site columns, the name of the group this column belongs to. Helps organize related columns.'
      )
      .nullish(),
    contentApprovalStatus: microsoft_graph_contentApprovalStatusColumn.optional(),
    currency: microsoft_graph_currencyColumn.optional(),
    dateTime: microsoft_graph_dateTimeColumn.optional(),
    defaultValue: microsoft_graph_defaultColumnValue.optional(),
    enforceUniqueValues: z
      .boolean()
      .describe('If true, no two list items may have the same value for this column.')
      .nullish(),
    geolocation: microsoft_graph_geolocationColumn.optional(),
    hidden: z
      .boolean()
      .describe('Specifies whether the column is displayed in the user interface.')
      .nullish(),
    hyperlinkOrPicture: microsoft_graph_hyperlinkOrPictureColumn.optional(),
    indexed: z
      .boolean()
      .describe('Specifies whether the column values can be used for sorting and searching.')
      .nullish(),
    isDeletable: z.boolean().describe('Indicates whether this column can be deleted.').nullish(),
    isReorderable: z
      .boolean()
      .describe('Indicates whether values in the column can be reordered. Read-only.')
      .nullish(),
    isSealed: z.boolean().describe('Specifies whether the column can be changed.').nullish(),
    lookup: microsoft_graph_lookupColumn.optional(),
    number: microsoft_graph_numberColumn.optional(),
    personOrGroup: microsoft_graph_personOrGroupColumn.optional(),
    propagateChanges: z
      .boolean()
      .describe(
        "If 'true', changes to this column will be propagated to lists that implement the column."
      )
      .nullish(),
  })
  .strict()
  .passthrough();
const microsoft_graph_contentTypeInfo = z
  .object({
    id: z.string().describe('The ID of the content type.').nullish(),
    name: z.string().describe('The name of the content type.').nullish(),
  })
  .strict();
const microsoft_graph_documentSetContent = z
  .object({
    contentType: microsoft_graph_contentTypeInfo.optional(),
    fileName: z
      .string()
      .describe(
        'Name of the file in resource folder that should be added as a default content or a template in the document set.'
      )
      .nullish(),
    folderName: z
      .string()
      .describe(
        'Folder name in which the file will be placed when a new document set is created in the library.'
      )
      .nullish(),
  })
  .strict();
const microsoft_graph_documentSet = z
  .object({
    allowedContentTypes: z
      .array(microsoft_graph_contentTypeInfo)
      .describe('Content types allowed in document set.')
      .optional(),
    defaultContents: z
      .array(microsoft_graph_documentSetContent)
      .describe('Default contents of document set.')
      .optional(),
    propagateWelcomePageChanges: z
      .boolean()
      .describe('Specifies whether to push welcome page changes to inherited content types.')
      .nullish(),
    shouldPrefixNameToFile: z
      .boolean()
      .describe('Indicates whether to add the name of the document set to each file name.')
      .nullish(),
    welcomePageUrl: z.string().describe('Welcome page absolute URL.').nullish(),
    sharedColumns: z.array(microsoft_graph_columnDefinition).optional(),
    welcomePageColumns: z.array(microsoft_graph_columnDefinition).optional(),
  })
  .strict();
const microsoft_graph_contentTypeOrder = z
  .object({
    default: z.boolean().describe('Indicates whether this is the default content type').nullish(),
    position: z
      .number()
      .gte(-2147483648)
      .lte(2147483647)
      .describe('Specifies the position in which the content type appears in the selection UI.')
      .nullish(),
  })
  .strict();
const microsoft_graph_columnLink = z
  .object({
    id: z.string().describe('The unique identifier for an entity. Read-only.').optional(),
    name: z.string().describe('The name of the column  in this content type.').nullish(),
  })
  .strict();
const microsoft_graph_contentType: z.ZodType<microsoft_graph_contentType> = z.lazy(() =>
  z
    .object({
      id: z.string().describe('The unique identifier for an entity. Read-only.').optional(),
      associatedHubsUrls: z
        .array(z.string().nullable())
        .describe(
          'List of canonical URLs for hub sites with which this content type is associated to. This will contain all hub sites where this content type is queued to be enforced or is already enforced. Enforcing a content type means that the content type is applied to the lists in the enforced sites.'
        )
        .optional(),
      description: z.string().describe('The descriptive text for the item.').nullish(),
      documentSet: microsoft_graph_documentSet.optional(),
      documentTemplate: microsoft_graph_documentSetContent.optional(),
      group: z
        .string()
        .describe(
          'The name of the group this content type belongs to. Helps organize related content types.'
        )
        .nullish(),
      hidden: z
        .boolean()
        .describe("Indicates whether the content type is hidden in the list's 'New' menu.")
        .nullish(),
      inheritedFrom: microsoft_graph_itemReference.optional(),
      isBuiltIn: z
        .boolean()
        .describe('Specifies if a content type is a built-in content type.')
        .nullish(),
      name: z.string().describe('The name of the content type.').nullish(),
      order: microsoft_graph_contentTypeOrder.optional(),
      parentId: z.string().describe('The unique identifier of the content type.').nullish(),
      propagateChanges: z
        .boolean()
        .describe(
          'If true, any changes made to the content type are pushed to inherited content types and lists that implement the content type.'
        )
        .nullish(),
      readOnly: z
        .boolean()
        .describe(
          "If true, the content type can't be modified unless this value is first set to false."
        )
        .nullish(),
      sealed: z
        .boolean()
        .describe(
          "If true, the content type can't be modified by users or through push-down operations. Only site collection administrators can seal or unseal content types."
        )
        .nullish(),
      base: microsoft_graph_contentType.optional(),
      baseTypes: z
        .array(microsoft_graph_contentType)
        .describe('The collection of content types that are ancestors of this content type.')
        .optional(),
      columnLinks: z
        .array(microsoft_graph_columnLink)
        .describe('The collection of columns that are required by this content type.')
        .optional(),
      columnPositions: z
        .array(microsoft_graph_columnDefinition)
        .describe('Column order information in a content type.')
        .optional(),
      columns: z
        .array(microsoft_graph_columnDefinition)
        .describe('The collection of column definitions for this content type.')
        .optional(),
    })
    .strict()
);
const microsoft_graph_storagePlanInformation = z
  .object({
    upgradeAvailable: z
      .boolean()
      .describe('Indicates whether there are higher storage quota plans available. Read-only.')
      .nullish(),
  })
  .strict();
const microsoft_graph_quota = z
  .object({
    deleted: z
      .number()
      .describe('Total space consumed by files in the recycle bin, in bytes. Read-only.')
      .nullish(),
    remaining: z
      .number()
      .describe('Total space remaining before reaching the capacity limit, in bytes. Read-only.')
      .nullish(),
    state: z
      .string()
      .describe('Enumeration value that indicates the state of the storage space. Read-only.')
      .nullish(),
    storagePlanInformation: microsoft_graph_storagePlanInformation.optional(),
    total: z.number().describe('Total allowed storage space, in bytes. Read-only.').nullish(),
    used: z.number().describe('Total space used, in bytes. Read-only.').nullish(),
  })
  .strict();
const microsoft_graph_systemFacet = z.object({}).strict();
const microsoft_graph_listInfo = z
  .object({
    contentTypesEnabled: z
      .boolean()
      .describe('If true, indicates that content types are enabled for this list.')
      .nullish(),
    hidden: z
      .boolean()
      .describe(
        "If true, indicates that the list isn't normally visible in the SharePoint user experience."
      )
      .nullish(),
    template: z
      .string()
      .describe(
        'An enumerated value that represents the base list template used in creating the list. Possible values include documentLibrary, genericList, task, survey, announcements, contacts, and more.'
      )
      .nullish(),
  })
  .strict();
const microsoft_graph_documentSetVersionItem = z
  .object({
    itemId: z.string().describe('The unique identifier for the item.').nullish(),
    title: z.string().describe('The title of the item.').nullish(),
    versionId: z.string().describe('The version ID of the item.').nullish(),
  })
  .strict();
const microsoft_graph_documentSetVersion = z
  .object({
    comment: z.string().describe('Comment about the captured version.').nullish(),
    createdBy: microsoft_graph_identitySet.optional(),
    createdDateTime: z
      .string()
      .regex(
        /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
      )
      .datetime({ offset: true })
      .describe('Date and time when this version was created.')
      .nullish(),
    items: z
      .array(microsoft_graph_documentSetVersionItem)
      .describe('Items within the document set that are captured as part of this version.')
      .optional(),
    shouldCaptureMinorVersion: z
      .boolean()
      .describe(
        'If true, minor versions of items are also captured; otherwise, only major versions are captured. The default value is false.'
      )
      .nullish(),
  })
  .strict();
const microsoft_graph_fieldValueSet = z
  .object({ id: z.string().describe('The unique identifier for an entity. Read-only.').optional() })
  .strict();
const microsoft_graph_publicationFacet = z
  .object({
    checkedOutBy: microsoft_graph_identitySet.optional(),
    level: z
      .string()
      .describe(
        'The state of publication for this document. Either published or checkout. Read-only.'
      )
      .nullish(),
    versionId: z
      .string()
      .describe(
        'The unique identifier for the version that is visible to the current caller. Read-only.'
      )
      .nullish(),
  })
  .strict();
const microsoft_graph_listItemVersion = z
  .object({
    id: z.string().describe('The unique identifier for an entity. Read-only.').optional(),
    lastModifiedBy: microsoft_graph_identitySet.optional(),
    lastModifiedDateTime: z
      .string()
      .regex(
        /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
      )
      .datetime({ offset: true })
      .describe('Date and time the version was last modified. Read-only.')
      .nullish(),
    publication: microsoft_graph_publicationFacet.optional(),
    fields: microsoft_graph_fieldValueSet.optional(),
  })
  .strict();
const microsoft_graph_listItem = z
  .object({
    id: z.string().describe('The unique identifier for an entity. Read-only.').optional(),
    createdBy: microsoft_graph_identitySet.optional(),
    createdDateTime: z
      .string()
      .regex(
        /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
      )
      .datetime({ offset: true })
      .describe('Date and time of item creation. Read-only.')
      .optional(),
    description: z
      .string()
      .describe('Provides a user-visible description of the item. Optional.')
      .nullish(),
    eTag: z.string().describe('ETag for the item. Read-only.').nullish(),
    lastModifiedBy: microsoft_graph_identitySet.optional(),
    lastModifiedDateTime: z
      .string()
      .regex(
        /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
      )
      .datetime({ offset: true })
      .describe('Date and time the item was last modified. Read-only.')
      .optional(),
    name: z.string().describe('The name of the item. Read-write.').nullish(),
    parentReference: microsoft_graph_itemReference.optional(),
    webUrl: z
      .string()
      .describe(
        'URL that either displays the resource in the browser (for Office file formats), or is a direct link to the file (for other formats). Read-only.'
      )
      .nullish(),
    createdByUser: microsoft_graph_user
      .describe('[Note: Simplified from 132 properties to 25 most common ones]')
      .optional(),
    lastModifiedByUser: microsoft_graph_user
      .describe('[Note: Simplified from 132 properties to 25 most common ones]')
      .optional(),
    contentType: microsoft_graph_contentTypeInfo.optional(),
    sharepointIds: microsoft_graph_sharepointIds.optional(),
    analytics: microsoft_graph_itemAnalytics.optional(),
    documentSetVersions: z
      .array(microsoft_graph_documentSetVersion)
      .describe('Version information for a document set version created by a user.')
      .optional(),
    driveItem: microsoft_graph_driveItem
      .describe('[Note: Simplified from 45 properties to 25 most common ones]')
      .optional(),
    fields: microsoft_graph_fieldValueSet.optional(),
    versions: z
      .array(microsoft_graph_listItemVersion)
      .describe('The list of previous versions of the list item.')
      .optional(),
  })
  .strict();
const microsoft_graph_longRunningOperationStatus = z.enum([
  'notStarted',
  'running',
  'succeeded',
  'failed',
  'unknownFutureValue',
]);
const microsoft_graph_richLongRunningOperation = z
  .object({
    id: z.string().describe('The unique identifier for an entity. Read-only.').optional(),
    createdDateTime: z
      .string()
      .regex(
        /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
      )
      .datetime({ offset: true })
      .describe(
        'The start time of the operation. The timestamp type represents date and time information using ISO 8601 format and is always in UTC. For example, midnight UTC on Jan 1, 2014 is 2014-01-01T00:00:00Z.'
      )
      .nullish(),
    lastActionDateTime: z
      .string()
      .regex(
        /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
      )
      .datetime({ offset: true })
      .describe(
        'The time of the last action in the operation. The timestamp type represents date and time information using ISO 8601 format and is always in UTC. For example, midnight UTC on Jan 1, 2014 is 2014-01-01T00:00:00Z.'
      )
      .nullish(),
    resourceLocation: z
      .string()
      .describe('URI of the resource that the operation is performed on.')
      .nullish(),
    status: microsoft_graph_longRunningOperationStatus.optional(),
    statusDetail: z.string().describe('Details about the status of the operation.').nullish(),
    error: microsoft_graph_publicError.optional(),
    percentageComplete: z
      .number()
      .gte(-2147483648)
      .lte(2147483647)
      .describe('A value between 0 and 100 that indicates the progress of the operation.')
      .nullish(),
    resourceId: z.string().describe('The unique identifier for the result.').nullish(),
    type: z.string().describe('The type of the operation.').nullish(),
  })
  .strict();
const microsoft_graph_subscription = z
  .object({
    id: z.string().describe('The unique identifier for an entity. Read-only.').optional(),
    applicationId: z
      .string()
      .describe(
        'Optional. Identifier of the application used to create the subscription. Read-only.'
      )
      .nullish(),
    changeType: z
      .string()
      .describe(
        'Required. Indicates the type of change in the subscribed resource that raises a change notification. The supported values are: created, updated, deleted. Multiple values can be combined using a comma-separated list. Note:  Drive root item and list change notifications support only the updated changeType. User and group change notifications support updated and deleted changeType. Use updated to receive notifications when user or group is created, updated, or soft deleted. Use deleted to receive notifications when user or group is permanently deleted.'
      )
      .optional(),
    clientState: z
      .string()
      .describe(
        'Optional. Specifies the value of the clientState property sent by the service in each change notification. The maximum length is 128 characters. The client can check that the change notification came from the service by comparing the value of the clientState property sent with the subscription with the value of the clientState property received with each change notification.'
      )
      .nullish(),
    creatorId: z
      .string()
      .describe(
        'Optional. Identifier of the user or service principal that created the subscription. If the app used delegated permissions to create the subscription, this field contains the ID of the signed-in user the app called on behalf of. If the app used application permissions, this field contains the ID of the service principal corresponding to the app. Read-only.'
      )
      .nullish(),
    encryptionCertificate: z
      .string()
      .describe(
        'Optional. A base64-encoded representation of a certificate with a public key used to encrypt resource data in change notifications. Optional but required when includeResourceData is true.'
      )
      .nullish(),
    encryptionCertificateId: z
      .string()
      .describe(
        'Optional. A custom app-provided identifier to help identify the certificate needed to decrypt resource data.'
      )
      .nullish(),
    expirationDateTime: z
      .string()
      .regex(
        /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
      )
      .datetime({ offset: true })
      .describe(
        'Required. Specifies the date and time when the webhook subscription expires. The time is in UTC, and can be an amount of time from subscription creation that varies for the resource subscribed to. Any value under 45 minutes after the time of the request is automatically set to 45 minutes after the request time. For the maximum supported subscription length of time, see Subscription lifetime.'
      )
      .optional(),
    includeResourceData: z
      .boolean()
      .describe(
        'Optional. When set to true, change notifications include resource data (such as content of a chat message).'
      )
      .nullish(),
    latestSupportedTlsVersion: z
      .string()
      .describe(
        'Optional. Specifies the latest version of Transport Layer Security (TLS) that the notification endpoint, specified by notificationUrl, supports. The possible values are: v10, v11, v12, v13. For subscribers whose notification endpoint supports a version lower than the currently recommended version (TLS 1.2), specifying this property by a set timeline allows them to temporarily use their deprecated version of TLS before completing their upgrade to TLS 1.2. For these subscribers, not setting this property per the timeline would result in subscription operations failing. For subscribers whose notification endpoint already supports TLS 1.2, setting this property is optional. In such cases, Microsoft Graph defaults the property to v1_2.'
      )
      .nullish(),
    lifecycleNotificationUrl: z
      .string()
      .describe(
        'Required for Teams resources if  the expirationDateTime value is more than 1 hour from now; optional otherwise. The URL of the endpoint that receives lifecycle notifications, including subscriptionRemoved, reauthorizationRequired, and missed notifications. This URL must make use of the HTTPS protocol. For more information, see Reduce missing subscriptions and change notifications.'
      )
      .nullish(),
    notificationQueryOptions: z
      .string()
      .describe(
        'Optional. OData query options for specifying value for the targeting resource. Clients receive notifications when resource reaches the state matching the query options provided here. With this new property in the subscription creation payload along with all existing properties, Webhooks deliver notifications whenever a resource reaches the desired state mentioned in the notificationQueryOptions property. For example, when the print job is completed or when a print job resource isFetchable property value becomes true etc.  Supported only for Universal Print Service. For more information, see Subscribe to change notifications from cloud printing APIs using Microsoft Graph.'
      )
      .nullish(),
    notificationUrl: z
      .string()
      .describe(
        'Required. The URL of the endpoint that receives the change notifications. This URL must make use of the HTTPS protocol. Any query string parameter included in the notificationUrl property is included in the HTTP POST request when Microsoft Graph sends the change notifications.'
      )
      .optional(),
    notificationUrlAppId: z
      .string()
      .describe(
        'Optional. The app ID that the subscription service can use to generate the validation token. The value allows the client to validate the authenticity of the notification received.'
      )
      .nullish(),
    resource: z
      .string()
      .describe(
        "Required. Specifies the resource that is monitored for changes. Don't include the base URL (https://graph.microsoft.com/v1.0/). See the possible resource path values for each supported resource."
      )
      .optional(),
  })
  .strict();
const microsoft_graph_list: z.ZodType<microsoft_graph_list> = z.lazy(() =>
  z
    .object({
      id: z.string().describe('The unique identifier for an entity. Read-only.').optional(),
      createdBy: microsoft_graph_identitySet.optional(),
      createdDateTime: z
        .string()
        .regex(
          /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
        )
        .datetime({ offset: true })
        .describe('Date and time of item creation. Read-only.')
        .optional(),
      description: z
        .string()
        .describe('Provides a user-visible description of the item. Optional.')
        .nullish(),
      eTag: z.string().describe('ETag for the item. Read-only.').nullish(),
      lastModifiedBy: microsoft_graph_identitySet.optional(),
      lastModifiedDateTime: z
        .string()
        .regex(
          /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
        )
        .datetime({ offset: true })
        .describe('Date and time the item was last modified. Read-only.')
        .optional(),
      name: z.string().describe('The name of the item. Read-write.').nullish(),
      parentReference: microsoft_graph_itemReference.optional(),
      webUrl: z
        .string()
        .describe(
          'URL that either displays the resource in the browser (for Office file formats), or is a direct link to the file (for other formats). Read-only.'
        )
        .nullish(),
      createdByUser: microsoft_graph_user
        .describe('[Note: Simplified from 132 properties to 25 most common ones]')
        .optional(),
      lastModifiedByUser: microsoft_graph_user
        .describe('[Note: Simplified from 132 properties to 25 most common ones]')
        .optional(),
      displayName: z.string().describe('The displayable title of the list.').nullish(),
      list: microsoft_graph_listInfo.optional(),
      sharepointIds: microsoft_graph_sharepointIds.optional(),
      system: microsoft_graph_systemFacet.optional(),
      columns: z
        .array(microsoft_graph_columnDefinition)
        .describe('The collection of field definitions for this list.')
        .optional(),
      contentTypes: z
        .array(microsoft_graph_contentType)
        .describe('The collection of content types present in this list.')
        .optional(),
      drive: microsoft_graph_drive.optional(),
      items: z
        .array(microsoft_graph_listItem)
        .describe('All items contained in the list.')
        .optional(),
      operations: z
        .array(microsoft_graph_richLongRunningOperation)
        .describe('The collection of long-running operations on the list.')
        .optional(),
      subscriptions: z
        .array(microsoft_graph_subscription)
        .describe('The set of subscriptions on the list.')
        .optional(),
    })
    .strict()
);
const microsoft_graph_drive: z.ZodType<microsoft_graph_drive> = z.lazy(() =>
  z
    .object({
      id: z.string().describe('The unique identifier for an entity. Read-only.').optional(),
      createdBy: microsoft_graph_identitySet.optional(),
      createdDateTime: z
        .string()
        .regex(
          /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
        )
        .datetime({ offset: true })
        .describe('Date and time of item creation. Read-only.')
        .optional(),
      description: z
        .string()
        .describe('Provides a user-visible description of the item. Optional.')
        .nullish(),
      eTag: z.string().describe('ETag for the item. Read-only.').nullish(),
      lastModifiedBy: microsoft_graph_identitySet.optional(),
      lastModifiedDateTime: z
        .string()
        .regex(
          /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
        )
        .datetime({ offset: true })
        .describe('Date and time the item was last modified. Read-only.')
        .optional(),
      name: z.string().describe('The name of the item. Read-write.').nullish(),
      parentReference: microsoft_graph_itemReference.optional(),
      webUrl: z
        .string()
        .describe(
          'URL that either displays the resource in the browser (for Office file formats), or is a direct link to the file (for other formats). Read-only.'
        )
        .nullish(),
      createdByUser: microsoft_graph_user
        .describe('[Note: Simplified from 132 properties to 25 most common ones]')
        .optional(),
      lastModifiedByUser: microsoft_graph_user
        .describe('[Note: Simplified from 132 properties to 25 most common ones]')
        .optional(),
      driveType: z
        .string()
        .describe(
          'Describes the type of drive represented by this resource. OneDrive personal drives return personal. OneDrive for Business returns business. SharePoint document libraries return documentLibrary. Read-only.'
        )
        .nullish(),
      owner: microsoft_graph_identitySet.optional(),
      quota: microsoft_graph_quota.optional(),
      sharePointIds: microsoft_graph_sharepointIds.optional(),
      system: microsoft_graph_systemFacet.optional(),
      bundles: z
        .array(microsoft_graph_driveItem)
        .describe(
          'Collection of bundles (albums and multi-select-shared sets of items). Only in personal OneDrive.'
        )
        .optional(),
      following: z
        .array(microsoft_graph_driveItem)
        .describe('The list of items the user is following. Only in OneDrive for Business.')
        .optional(),
      items: z
        .array(microsoft_graph_driveItem)
        .describe('All items contained in the drive. Read-only. Nullable.')
        .optional(),
      list: microsoft_graph_list.optional(),
      root: microsoft_graph_driveItem
        .describe('[Note: Simplified from 45 properties to 25 most common ones]')
        .optional(),
      special: z
        .array(microsoft_graph_driveItem)
        .describe('Collection of common folders available in OneDrive. Read-only. Nullable.')
        .optional(),
    })
    .strict()
);
const microsoft_graph_baseItem = z
  .object({
    id: z.string().describe('The unique identifier for an entity. Read-only.').optional(),
    createdBy: microsoft_graph_identitySet.optional(),
    createdDateTime: z
      .string()
      .regex(
        /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
      )
      .datetime({ offset: true })
      .describe('Date and time of item creation. Read-only.')
      .optional(),
    description: z
      .string()
      .describe('Provides a user-visible description of the item. Optional.')
      .nullish(),
    eTag: z.string().describe('ETag for the item. Read-only.').nullish(),
    lastModifiedBy: microsoft_graph_identitySet.optional(),
    lastModifiedDateTime: z
      .string()
      .regex(
        /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
      )
      .datetime({ offset: true })
      .describe('Date and time the item was last modified. Read-only.')
      .optional(),
    name: z.string().describe('The name of the item. Read-write.').nullish(),
    parentReference: microsoft_graph_itemReference.optional(),
    webUrl: z
      .string()
      .describe(
        'URL that either displays the resource in the browser (for Office file formats), or is a direct link to the file (for other formats). Read-only.'
      )
      .nullish(),
    createdByUser: microsoft_graph_user
      .describe('[Note: Simplified from 132 properties to 25 most common ones]')
      .optional(),
    lastModifiedByUser: microsoft_graph_user
      .describe('[Note: Simplified from 132 properties to 25 most common ones]')
      .optional(),
  })
  .strict();
const microsoft_graph_site = z
  .object({
    id: z.string().describe('The unique identifier for an entity. Read-only.').optional(),
    name: z.string().describe('The name of the item. Read-write.').nullish(),
    displayName: z.string().describe('The full title for the site. Read-only.').nullish(),
    description: z
      .string()
      .describe('Provides a user-visible description of the item. Optional.')
      .nullish(),
    createdDateTime: z
      .string()
      .regex(
        /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
      )
      .datetime({ offset: true })
      .describe('Date and time of item creation. Read-only.')
      .optional(),
    lastModifiedDateTime: z
      .string()
      .regex(
        /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
      )
      .datetime({ offset: true })
      .describe('Date and time the item was last modified. Read-only.')
      .optional(),
    error: microsoft_graph_publicError.optional(),
    createdBy: microsoft_graph_identitySet.optional(),
    eTag: z.string().describe('ETag for the item. Read-only.').nullish(),
    lastModifiedBy: microsoft_graph_identitySet.optional(),
    parentReference: microsoft_graph_itemReference.optional(),
    webUrl: z
      .string()
      .describe(
        'URL that either displays the resource in the browser (for Office file formats), or is a direct link to the file (for other formats). Read-only.'
      )
      .nullish(),
    createdByUser: microsoft_graph_user
      .describe('[Note: Simplified from 132 properties to 25 most common ones]')
      .optional(),
    lastModifiedByUser: microsoft_graph_user
      .describe('[Note: Simplified from 132 properties to 25 most common ones]')
      .optional(),
    isPersonalSite: z
      .boolean()
      .describe('Identifies whether the site is personal or not. Read-only.')
      .nullish(),
    root: microsoft_graph_root.optional(),
    sharepointIds: microsoft_graph_sharepointIds.optional(),
    siteCollection: microsoft_graph_siteCollection.optional(),
    analytics: microsoft_graph_itemAnalytics.optional(),
    columns: z
      .array(microsoft_graph_columnDefinition)
      .describe('The collection of column definitions reusable across lists under this site.')
      .optional(),
    contentTypes: z
      .array(microsoft_graph_contentType)
      .describe('The collection of content types defined for this site.')
      .optional(),
    drive: microsoft_graph_drive.optional(),
    drives: z
      .array(microsoft_graph_drive)
      .describe('The collection of drives (document libraries) under this site.')
      .optional(),
    externalColumns: z.array(microsoft_graph_columnDefinition).optional(),
    items: z
      .array(microsoft_graph_baseItem)
      .describe(
        "Used to address any item contained in this site. This collection can't be enumerated."
      )
      .optional(),
  })
  .strict()
  .passthrough();
const microsoft_graph_siteCollectionResponse = z
  .object({
    '@odata.count': z.number().int().nullable(),
    '@odata.nextLink': z.string().nullable(),
    value: z.array(microsoft_graph_site),
  })
  .partial()
  .strict();
const microsoft_graph_driveCollectionResponse = z
  .object({
    '@odata.count': z.number().int().nullable(),
    '@odata.nextLink': z.string().nullable(),
    value: z.array(microsoft_graph_drive),
  })
  .partial()
  .strict();
const microsoft_graph_baseItemCollectionResponse = z
  .object({
    '@odata.count': z.number().int().nullable(),
    '@odata.nextLink': z.string().nullable(),
    value: z.array(microsoft_graph_baseItem),
  })
  .partial()
  .strict();
const microsoft_graph_listCollectionResponse = z
  .object({
    '@odata.count': z.number().int().nullable(),
    '@odata.nextLink': z.string().nullable(),
    value: z.array(microsoft_graph_list),
  })
  .partial()
  .strict();
const microsoft_graph_listItemCollectionResponse = z
  .object({
    '@odata.count': z.number().int().nullable(),
    '@odata.nextLink': z.string().nullable(),
    value: z.array(microsoft_graph_listItem),
  })
  .partial()
  .strict();
const BaseDeltaFunctionResponse = z
  .object({ '@odata.nextLink': z.string().nullable(), '@odata.deltaLink': z.string().nullable() })
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
  microsoft_graph_publicErrorDetail,
  microsoft_graph_publicInnerError,
  microsoft_graph_publicError,
  microsoft_graph_identity,
  microsoft_graph_identitySet,
  microsoft_graph_sharepointIds,
  microsoft_graph_itemReference,
  microsoft_graph_assignedLicense,
  microsoft_graph_assignedPlan,
  microsoft_graph_authorizationInfo,
  microsoft_graph_customSecurityAttributeValue,
  microsoft_graph_user,
  microsoft_graph_root,
  microsoft_graph_siteArchiveStatus,
  microsoft_graph_siteArchivalDetails,
  microsoft_graph_siteCollection,
  microsoft_graph_itemActionStat,
  microsoft_graph_incompleteData,
  microsoft_graph_accessAction,
  microsoft_graph_audio,
  microsoft_graph_album,
  microsoft_graph_bundle,
  microsoft_graph_deleted,
  microsoft_graph_hashes,
  microsoft_graph_file,
  microsoft_graph_fileSystemInfo,
  microsoft_graph_folderView,
  microsoft_graph_folder,
  microsoft_graph_image,
  microsoft_graph_geoCoordinates,
  microsoft_graph_malware,
  microsoft_graph_package,
  microsoft_graph_pendingContentUpdate,
  microsoft_graph_pendingOperations,
  microsoft_graph_driveItem,
  microsoft_graph_itemActivity,
  microsoft_graph_itemActivityStat,
  microsoft_graph_itemAnalytics,
  microsoft_graph_columnTypes,
  microsoft_graph_booleanColumn,
  microsoft_graph_calculatedColumn,
  microsoft_graph_choiceColumn,
  microsoft_graph_contentApprovalStatusColumn,
  microsoft_graph_currencyColumn,
  microsoft_graph_dateTimeColumn,
  microsoft_graph_defaultColumnValue,
  microsoft_graph_geolocationColumn,
  microsoft_graph_hyperlinkOrPictureColumn,
  microsoft_graph_lookupColumn,
  microsoft_graph_numberColumn,
  microsoft_graph_personOrGroupColumn,
  microsoft_graph_columnDefinition,
  microsoft_graph_contentTypeInfo,
  microsoft_graph_documentSetContent,
  microsoft_graph_documentSet,
  microsoft_graph_contentTypeOrder,
  microsoft_graph_columnLink,
  microsoft_graph_contentType,
  microsoft_graph_storagePlanInformation,
  microsoft_graph_quota,
  microsoft_graph_systemFacet,
  microsoft_graph_listInfo,
  microsoft_graph_documentSetVersionItem,
  microsoft_graph_documentSetVersion,
  microsoft_graph_fieldValueSet,
  microsoft_graph_publicationFacet,
  microsoft_graph_listItemVersion,
  microsoft_graph_listItem,
  microsoft_graph_longRunningOperationStatus,
  microsoft_graph_richLongRunningOperation,
  microsoft_graph_subscription,
  microsoft_graph_list,
  microsoft_graph_drive,
  microsoft_graph_baseItem,
  microsoft_graph_site,
  microsoft_graph_siteCollectionResponse,
  microsoft_graph_driveCollectionResponse,
  microsoft_graph_baseItemCollectionResponse,
  microsoft_graph_listCollectionResponse,
  microsoft_graph_listItemCollectionResponse,
  BaseDeltaFunctionResponse,
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
  {
    method: 'get',
    path: '/sites',
    alias: 'search-sharepoint-sites',
    description: `Search across a SharePoint tenant for sites that match keywords provided. The only property that works for sorting is createdDateTime. The search filter is a free text search that uses multiple properties when retrieving the search results.`,
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
        schema: microsoft_graph_siteCollectionResponse,
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
    path: '/sites/:siteId',
    alias: 'get-sharepoint-site',
    description: `Retrieve properties and relationships for a site resource.
A site resource represents a team site in SharePoint.`,
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
        description: `Retrieved entity`,
        schema: z
          .object({
            id: z.string().describe('The unique identifier for an entity. Read-only.').optional(),
            name: z.string().describe('The name of the item. Read-write.').nullish(),
            displayName: z.string().describe('The full title for the site. Read-only.').nullish(),
            description: z
              .string()
              .describe('Provides a user-visible description of the item. Optional.')
              .nullish(),
            createdDateTime: z
              .string()
              .regex(
                /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
              )
              .datetime({ offset: true })
              .describe('Date and time of item creation. Read-only.')
              .optional(),
            lastModifiedDateTime: z
              .string()
              .regex(
                /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
              )
              .datetime({ offset: true })
              .describe('Date and time the item was last modified. Read-only.')
              .optional(),
            error: microsoft_graph_publicError.optional(),
            createdBy: microsoft_graph_identitySet.optional(),
            eTag: z.string().describe('ETag for the item. Read-only.').nullish(),
            lastModifiedBy: microsoft_graph_identitySet.optional(),
            parentReference: microsoft_graph_itemReference.optional(),
            webUrl: z
              .string()
              .describe(
                'URL that either displays the resource in the browser (for Office file formats), or is a direct link to the file (for other formats). Read-only.'
              )
              .nullish(),
            createdByUser: microsoft_graph_user
              .describe('[Note: Simplified from 132 properties to 25 most common ones]')
              .optional(),
            lastModifiedByUser: microsoft_graph_user
              .describe('[Note: Simplified from 132 properties to 25 most common ones]')
              .optional(),
            isPersonalSite: z
              .boolean()
              .describe('Identifies whether the site is personal or not. Read-only.')
              .nullish(),
            root: microsoft_graph_root.optional(),
            sharepointIds: microsoft_graph_sharepointIds.optional(),
            siteCollection: microsoft_graph_siteCollection.optional(),
            analytics: microsoft_graph_itemAnalytics.optional(),
            columns: z
              .array(microsoft_graph_columnDefinition)
              .describe(
                'The collection of column definitions reusable across lists under this site.'
              )
              .optional(),
            contentTypes: z
              .array(microsoft_graph_contentType)
              .describe('The collection of content types defined for this site.')
              .optional(),
            drive: microsoft_graph_drive.optional(),
            drives: z
              .array(microsoft_graph_drive)
              .describe('The collection of drives (document libraries) under this site.')
              .optional(),
            externalColumns: z.array(microsoft_graph_columnDefinition).optional(),
            items: z
              .array(microsoft_graph_baseItem)
              .describe(
                "Used to address any item contained in this site. This collection can't be enumerated."
              )
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
    path: '/sites/:siteId/drives',
    alias: 'list-sharepoint-site-drives',
    description: `The collection of drives (document libraries) under this site.`,
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
        schema: microsoft_graph_driveCollectionResponse,
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
    path: '/sites/:siteId/drives/:driveId',
    alias: 'get-sharepoint-site-drive-by-id',
    description: `The collection of drives (document libraries) under this site.`,
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
        schema: microsoft_graph_drive,
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
    path: '/sites/:siteId/getByPath(path&#x3D;&#x27;:path&#x27;)',
    alias: 'get-sharepoint-site-by-path',
    description: `Invoke function getByPath`,
    requestFormat: 'json',
    response: z.void(),
    errors: [
      {
        status: NaN,
        description: `Success`,
        schema: z
          .object({
            id: z.string().describe('The unique identifier for an entity. Read-only.').optional(),
            name: z.string().describe('The name of the item. Read-write.').nullish(),
            displayName: z.string().describe('The full title for the site. Read-only.').nullish(),
            description: z
              .string()
              .describe('Provides a user-visible description of the item. Optional.')
              .nullish(),
            createdDateTime: z
              .string()
              .regex(
                /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
              )
              .datetime({ offset: true })
              .describe('Date and time of item creation. Read-only.')
              .optional(),
            lastModifiedDateTime: z
              .string()
              .regex(
                /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
              )
              .datetime({ offset: true })
              .describe('Date and time the item was last modified. Read-only.')
              .optional(),
            error: microsoft_graph_publicError.optional(),
            createdBy: microsoft_graph_identitySet.optional(),
            eTag: z.string().describe('ETag for the item. Read-only.').nullish(),
            lastModifiedBy: microsoft_graph_identitySet.optional(),
            parentReference: microsoft_graph_itemReference.optional(),
            webUrl: z
              .string()
              .describe(
                'URL that either displays the resource in the browser (for Office file formats), or is a direct link to the file (for other formats). Read-only.'
              )
              .nullish(),
            createdByUser: microsoft_graph_user
              .describe('[Note: Simplified from 132 properties to 25 most common ones]')
              .optional(),
            lastModifiedByUser: microsoft_graph_user
              .describe('[Note: Simplified from 132 properties to 25 most common ones]')
              .optional(),
            isPersonalSite: z
              .boolean()
              .describe('Identifies whether the site is personal or not. Read-only.')
              .nullish(),
            root: microsoft_graph_root.optional(),
            sharepointIds: microsoft_graph_sharepointIds.optional(),
            siteCollection: microsoft_graph_siteCollection.optional(),
            analytics: microsoft_graph_itemAnalytics.optional(),
            columns: z
              .array(microsoft_graph_columnDefinition)
              .describe(
                'The collection of column definitions reusable across lists under this site.'
              )
              .optional(),
            contentTypes: z
              .array(microsoft_graph_contentType)
              .describe('The collection of content types defined for this site.')
              .optional(),
            drive: microsoft_graph_drive.optional(),
            drives: z
              .array(microsoft_graph_drive)
              .describe('The collection of drives (document libraries) under this site.')
              .optional(),
            externalColumns: z.array(microsoft_graph_columnDefinition).optional(),
            items: z
              .array(microsoft_graph_baseItem)
              .describe(
                "Used to address any item contained in this site. This collection can't be enumerated."
              )
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
    path: '/sites/:siteId/items',
    alias: 'list-sharepoint-site-items',
    description: `Used to address any item contained in this site. This collection can&#x27;t be enumerated.`,
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
        schema: microsoft_graph_baseItemCollectionResponse,
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
    path: '/sites/:siteId/items/:baseItemId',
    alias: 'get-sharepoint-site-item',
    description: `Used to address any item contained in this site. This collection can&#x27;t be enumerated.`,
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
        schema: microsoft_graph_baseItem,
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
    path: '/sites/:siteId/lists',
    alias: 'list-sharepoint-site-lists',
    description: `Get the collection of lists for a site. Lists with the system facet are hidden by default.
To list them, include system in your $select statement.`,
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
        schema: microsoft_graph_listCollectionResponse,
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
    path: '/sites/:siteId/lists/:listId',
    alias: 'get-sharepoint-site-list',
    description: `Get a list of rich long-running operations associated with a list.`,
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
        schema: microsoft_graph_list,
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
    path: '/sites/:siteId/lists/:listId/items',
    alias: 'list-sharepoint-site-list-items',
    description: `Get the collection of items in a list.`,
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
        schema: microsoft_graph_listItemCollectionResponse,
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
    path: '/sites/:siteId/lists/:listId/items/:listItemId',
    alias: 'get-sharepoint-site-list-item',
    description: `Returns the metadata for an item in a list.`,
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
        schema: microsoft_graph_listItem,
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
    path: '/sites/delta()',
    alias: 'get-sharepoint-sites-delta',
    description: `Get newly created, updated, or deleted sites without having to perform a full read of the entire sites collection. A delta function call for sites is similar to a GET request, except that by appropriately applying state tokens in one or more of these calls,
you can query for incremental changes in the sites. It allows you to maintain and synchronize a local store of a user&#x27;s sites without having to fetch all the sites from the server every time.
The application calls the API without specifying any parameters.
The service begins enumerating sites and returns pages of changes to these sites, accompanied by either an @odata.nextLink or an @odata.deltaLink.
Your application should continue making calls using the @odata.nextLink until there&#x27;s an @odata.deltaLink  in the response. After you receive all the changes, you can apply them to your local state.
To monitor future changes, call the delta API by using the @odata.deltaLink in the previous response. Any resources marked as deleted should be removed from your local state.`,
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
        name: '$select',
        type: 'Query',
        schema: z.array(z.string()).describe('Select properties to be returned').optional(),
      },
      {
        name: '$orderby',
        type: 'Query',
        schema: z.array(z.string()).describe('Order items by property values').optional(),
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
        description: `Success`,
        schema: BaseDeltaFunctionResponse.and(
          z
            .object({ value: z.array(microsoft_graph_site) })
            .partial()
            .strict()
        ),
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
