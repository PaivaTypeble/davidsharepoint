$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$outputPath = Join-Path $repoRoot 'sharepoint-route-document-generic-n8n-workflow.json'

function New-Node {
    param(
        [string] $Id,
        [string] $Name,
        [string] $Type,
        [double] $TypeVersion,
        [object] $Parameters,
        [int] $X,
        [int] $Y
    )

    return [ordered]@{
        parameters  = $Parameters
        id          = $Id
        name        = $Name
        type        = $Type
        typeVersion = $TypeVersion
        position    = @($X, $Y)
    }
}

    function New-N8nExpression {
      param(
        [string] $Expression
      )

      return "={{ $Expression }}"
    }

    function New-Connection {
      param(
        [string] $Node,
        [int] $Index = 0
      )

      return [ordered]@{
        node  = $Node
        type  = 'main'
        index = $Index
      }
    }

$configCode = @'
return [{
  json: {
    tenantId: 'YOUR_TENANT_ID',
    clientId: 'YOUR_CLIENT_ID',
    clientSecret: 'YOUR_CLIENT_SECRET',
    sourceFileUrl: '',
    destinationRootFolderUrl: '',
    mappingWorkbookUrl: '',
    mappingWorkbookFileName: 'Company Acronyms and NIPC.xlsx',
    targetFileName: '',
    dryRun: true,
  },
}];
'@

$buildResolveQueueCode = @'
const config = $('Config').first().json;

const firstNonEmpty = (...values) => {
  for (const value of values) {
    const normalized = String(value ?? '').trim();
    if (normalized) {
      return normalized;
    }
  }

  return '';
};

const normalizeConfiguredSharePointUrl = (value, fallbackOrigin = '') => {
  const raw = String(value ?? '')
    .trim()
    .replace(/^['"]+|['"]+$/g, '');

  if (!raw) {
    return '';
  }

  if (/^https?:\/\//i.test(raw)) {
    return raw;
  }

  if (/^\/\//.test(raw)) {
    return 'https:' + raw;
  }

  if (/^[^/\s]+\.sharepoint\.com(\/|$)/i.test(raw)) {
    return 'https://' + raw;
  }

  if (fallbackOrigin && (/^\//.test(raw) || /^(sites|teams)\//i.test(raw) || /^:[^/]+:\/r\//i.test(raw))) {
    const normalizedPath = raw.startsWith('/') ? raw : '/' + raw;
    return fallbackOrigin + normalizedPath;
  }

  return raw;
};

const parseAbsoluteUrl = (value) => {
  const raw = String(value ?? '').trim();
  const match = /^(https?):\/\/([^/?#]+)([^?#]*)?(\?[^#]*)?(#.*)?$/i.exec(raw);

  if (!match) {
    return null;
  }

  return {
    protocol: match[1].toLowerCase(),
    hostname: String(match[2]).split(':')[0].toLowerCase(),
    origin: match[1].toLowerCase() + '://' + match[2],
  };
};

const getOrigin = (value) => {
  return parseAbsoluteUrl(value)?.origin || '';
};

const combineUrl = (baseUrl, fileName) => {
  const normalizedBaseUrl = String(baseUrl || '').trim();
  const separator = normalizedBaseUrl.endsWith('/') ? '' : '/';
  return normalizedBaseUrl + separator + encodeURIComponent(fileName);
};

const required = ['tenantId', 'clientId', 'clientSecret', 'sourceFileUrl', 'destinationRootFolderUrl'];
for (const key of required) {
  const value = String(config[key] ?? '').trim();
  if (!value || value.startsWith('YOUR_')) {
    throw new Error('Set ' + key + ' in the Config node before running the workflow.');
  }

  if ((key === 'sourceFileUrl' || key === 'destinationRootFolderUrl') && /(^https?:\/\/)?contoso\.sharepoint\.com\b/i.test(value)) {
    throw new Error('Replace the demo ' + key + ' in the Config node before running the workflow.');
  }
}

const destinationRootFolderUrl = normalizeConfiguredSharePointUrl(config.destinationRootFolderUrl);
const destinationOrigin = getOrigin(destinationRootFolderUrl);
const sourceFileUrl = normalizeConfiguredSharePointUrl(config.sourceFileUrl, destinationOrigin);
const mappingWorkbookFileName = firstNonEmpty(config.mappingWorkbookFileName, 'Company Acronyms and NIPC.xlsx');
const mappingWorkbookUrl = normalizeConfiguredSharePointUrl(
  firstNonEmpty(
    config.mappingWorkbookUrl,
    combineUrl(destinationRootFolderUrl, mappingWorkbookFileName),
  ),
  destinationOrigin,
);

return [
  {
    json: {
      kind: 'source',
      sharePointUrl: sourceFileUrl,
      originHint: destinationOrigin,
      expectedType: 'file',
    },
  },
  {
    json: {
      kind: 'destinationRoot',
      sharePointUrl: destinationRootFolderUrl,
      originHint: destinationOrigin,
      expectedType: 'folder',
    },
  },
  {
    json: {
      kind: 'workbook',
      sharePointUrl: mappingWorkbookUrl,
      originHint: destinationOrigin,
      expectedType: 'file',
    },
  },
];
'@

$parseSharePointUrlCode = @'
const inputs = $input.all();
const output = [];

const normalizeSharePointUrl = (value, originHint = '') => {
  const raw = String(value ?? '')
    .trim()
    .replace(/^['"]+|['"]+$/g, '');

  if (!raw) {
    return '';
  }

  if (/^https?:\/\//i.test(raw)) {
    return raw;
  }

  if (/^\/\//.test(raw)) {
    return 'https:' + raw;
  }

  if (/^[^/\s]+\.sharepoint\.com(\/|$)/i.test(raw)) {
    return 'https://' + raw;
  }

  if (originHint && (/^\//.test(raw) || /^(sites|teams)\//i.test(raw) || /^:[^/]+:\/r\//i.test(raw))) {
    const normalizedPath = raw.startsWith('/') ? raw : '/' + raw;
    return originHint + normalizedPath;
  }

  return raw;
};

const parseAbsoluteUrl = (value) => {
  const raw = String(value ?? '').trim();
  const match = /^(https?):\/\/([^/?#]+)([^?#]*)?(\?[^#]*)?(#.*)?$/i.exec(raw);

  if (!match) {
    return null;
  }

  const query = {};
  const search = match[4] || '';

  if (search.startsWith('?')) {
    for (const part of search.slice(1).split('&').filter(Boolean)) {
      const separatorIndex = part.indexOf('=');
      const keyPart = separatorIndex >= 0 ? part.slice(0, separatorIndex) : part;
      const valuePart = separatorIndex >= 0 ? part.slice(separatorIndex + 1) : '';
      const key = decodeURIComponent(keyPart.replace(/\+/g, ' '));

      if (!(key in query)) {
        query[key] = decodeURIComponent(valuePart.replace(/\+/g, ' '));
      }
    }
  }

  return {
    protocol: match[1].toLowerCase(),
    hostname: String(match[2]).split(':')[0].toLowerCase(),
    pathname: match[3] || '/',
    query,
  };
};

const stripSpecialShareLinkPrefix = (path) => {
  const segments = String(path ?? '')
    .split('/')
    .filter(Boolean)
    .map((segment) => decodeURIComponent(segment));

  if (segments.length >= 3 && segments[0].startsWith(':') && segments[1].toLowerCase() === 'r') {
    return '/' + segments.slice(2).join('/');
  }

  return String(path ?? '');
};

const cleanItemSegments = (segments) => {
  let normalized = [...segments];
  const aspxIndex = normalized.findIndex((segment) => /\.aspx$/i.test(segment));

  if (aspxIndex >= 0) {
    normalized = normalized.slice(0, aspxIndex);
  }

  if (normalized.length > 0 && /^forms$/i.test(normalized[normalized.length - 1])) {
    normalized = normalized.slice(0, -1);
  }

  if (normalized.length > 0 && /^forms$/i.test(normalized[0])) {
    normalized = [];
  }

  return normalized;
};

for (const inputItem of inputs) {
  const source = inputItem.json;
  const normalizedUrlText = normalizeSharePointUrl(source.sharePointUrl, source.originHint);
  const inputUrl = parseAbsoluteUrl(normalizedUrlText);

  if (!inputUrl) {
    throw new Error('sharePointUrl is not a valid URL for ' + source.kind + '. Received: ' + JSON.stringify(normalizedUrlText || String(source.sharePointUrl ?? '')) + '.');
  }

  const hostname = inputUrl.hostname;
  if (!/\.sharepoint\.com$/i.test(hostname)) {
    throw new Error('sharePointUrl for ' + source.kind + ' must point to a SharePoint Online host. Received host: ' + JSON.stringify(hostname) + '.');
  }

  let serverRelativePath = '';

  for (const key of ['id', 'RootFolder']) {
    const value = inputUrl.query[key];
    if (value) {
      serverRelativePath = stripSpecialShareLinkPrefix(decodeURIComponent(value));
      break;
    }
  }

  if (!serverRelativePath) {
    serverRelativePath = stripSpecialShareLinkPrefix(decodeURIComponent(inputUrl.pathname || '/'));
  }

  serverRelativePath = serverRelativePath
    .replace(/\\/g, '/')
    .replace(/\/+/g, '/')
    .replace(/\/$/, '');

  if (!serverRelativePath.startsWith('/')) {
    serverRelativePath = '/' + serverRelativePath;
  }

  const segments = cleanItemSegments(serverRelativePath.split('/').filter(Boolean));
  if (!segments.length) {
    throw new Error('Could not determine the SharePoint path for ' + source.kind + '.');
  }

  const seen = new Set();

  const addCandidate = (siteParts, withinSite) => {
    if (!withinSite.length) {
      return;
    }

    const libraryName = decodeURIComponent(withinSite[0]);
    const folderRelativePath = withinSite
      .slice(1)
      .map((segment) => decodeURIComponent(segment))
      .join('/');

    const candidateSitePath = siteParts.length ? '/' + siteParts.join('/') : '';
    const siteResolveUrl = candidateSitePath
      ? 'https://graph.microsoft.com/v1.0/sites/' + hostname + ':' + candidateSitePath
      : 'https://graph.microsoft.com/v1.0/sites/' + hostname;

    const candidateKey = [source.kind, candidateSitePath, libraryName, folderRelativePath].join('|');
    if (seen.has(candidateKey)) {
      return;
    }

    seen.add(candidateKey);
    output.push({
      json: {
        ...source,
        hostname,
        serverRelativePath,
        candidateSitePath,
        siteResolveUrl,
        libraryName,
        folderRelativePath,
      },
    });
  };

  if (/^(sites|teams)$/i.test(segments[0])) {
    if (segments.length < 3) {
      throw new Error('The SharePoint URL for ' + source.kind + ' must include a site path and a document library.');
    }

    for (let siteLength = segments.length - 1; siteLength >= 2; siteLength--) {
      addCandidate(segments.slice(0, siteLength), segments.slice(siteLength));
    }
  } else {
    addCandidate([], segments);
  }
}

if (!output.length) {
  throw new Error('No candidate SharePoint paths could be built from the configured URLs.');
}

return output;
'@

$selectResolvedTargetsCode = @'
const responses = $input.all();
const candidates = $('Parse SharePoint URL').all();
const resolvedByKind = new Map();

for (let index = 0; index < responses.length; index++) {
  const response = responses[index].json;
  const candidate = candidates[index].json;

  if (resolvedByKind.has(candidate.kind)) {
    continue;
  }

  if (response.statusCode >= 200 && response.statusCode < 300 && response.body && response.body.id) {
    resolvedByKind.set(candidate.kind, {
      json: {
        ...candidate,
        siteId: response.body.id,
        siteName: response.body.displayName || response.body.name || '',
        siteWebUrl: response.body.webUrl || '',
      },
    });
  }
}

for (const kind of ['source', 'destinationRoot', 'workbook']) {
  if (!resolvedByKind.has(kind)) {
    throw new Error('None of the candidate SharePoint site paths resolved for ' + kind + '.');
  }
}

return ['source', 'destinationRoot', 'workbook'].map((kind) => resolvedByKind.get(kind));
'@

$pickDriveCode = @'
const sites = $('Select Resolved Targets').all();
const payloads = $input.all();

const normalize = (value = '') => decodeURIComponent(String(value))
  .replace(/%20/gi, ' ')
  .toLowerCase()
  .replace(/\s+/g, ' ')
  .trim();

const getUrlPathname = (value = '') => {
  const match = /^(https?):\/\/([^/?#]+)([^?#]*)?(\?[^#]*)?(#.*)?$/i.exec(String(value ?? '').trim());
  return match ? (match[3] || '/') : '';
};

return payloads.map((payloadItem, index) => {
  const site = sites[index].json;
  const drives = Array.isArray(payloadItem.json.value) ? payloadItem.json.value : [];

  if (!drives.length) {
    throw new Error('No document libraries were returned for ' + site.kind + '.');
  }

  const targetLibrary = normalize(site.libraryName);
  const serverRelativePath = normalize(site.serverRelativePath);

  const scoreDrive = (drive) => {
    const driveName = normalize(drive.name);
    const pathSegments = getUrlPathname(drive.webUrl || '').split('/').filter(Boolean);
    const lastSegment = normalize(pathSegments[pathSegments.length - 1] || '');

    let score = 0;
    if (driveName === targetLibrary) score = 100;
    if (lastSegment === targetLibrary) score = Math.max(score, 90);
    if (serverRelativePath.includes('/' + driveName + '/')) score = Math.max(score, 80);
    if (serverRelativePath.endsWith('/' + driveName)) score = Math.max(score, 80);

    return score;
  };

  const ranked = drives
    .map((drive) => ({ drive, score: scoreDrive(drive) }))
    .sort((left, right) => right.score - left.score);

  if (!ranked.length || ranked[0].score <= 0) {
    throw new Error('Could not match document library ' + JSON.stringify(site.libraryName) + ' for ' + site.kind + '.');
  }

  const selected = ranked[0].drive;

  return {
    json: {
      ...site,
      driveId: selected.id,
      driveName: selected.name,
      driveWebUrl: selected.webUrl || '',
    },
  };
});
'@

$buildRuntimeCode = @'
const resolvedItems = $input.all();
const contexts = $('Pick Drive').all();
const config = $('Config').first().json;

const firstNonEmpty = (...values) => {
  for (const value of values) {
    const normalized = String(value ?? '').trim();
    if (normalized) {
      return normalized;
    }
  }

  return '';
};

const result = {};

for (let index = 0; index < resolvedItems.length; index++) {
  const context = contexts[index].json;
  const item = resolvedItems[index].json;
  const isFile = !!item.file;
  const isFolder = !!item.folder;

  if (context.expectedType === 'file' && !isFile) {
    throw new Error(context.kind + ' must resolve to a file.');
  }

  if (context.expectedType === 'folder' && !isFolder) {
    throw new Error(context.kind + ' must resolve to a folder.');
  }

  result[context.kind] = {
    sourceUrl: context.sharePointUrl,
    siteId: context.siteId,
    siteName: context.siteName,
    driveId: context.driveId,
    driveName: context.driveName,
    itemId: item.id,
    itemName: item.name,
    webUrl: item.webUrl || '',
    isFile,
    isFolder,
  };
}

return [{
  json: {
    source: result.source,
    destinationRoot: result.destinationRoot,
    workbook: result.workbook,
    dryRun: Boolean(config.dryRun),
    targetFileNameOverride: firstNonEmpty(config.targetFileName) || null,
  },
}];
'@

$prepareExtractionCode = @'
const item = $input.first().json;
const fileName = String(item.source.itemName || '');
const extensionMatch = /\.[^.]+$/.exec(fileName);
const sourceExtension = extensionMatch ? extensionMatch[0].toLowerCase() : '';

const textLike = new Set(['.txt', '.csv', '.json', '.xml']);
const imageLike = new Set(['.png', '.jpg', '.jpeg', '.tif', '.tiff', '.bmp', '.gif', '.webp']);

let extractOperation = null;
let extractionStatus = 'native';
let extractionMessage = null;
let earlyStatus = null;
let ocrPending = false;

if (textLike.has(sourceExtension)) {
  extractOperation = 'text';
} else if (sourceExtension === '.pdf') {
  extractOperation = 'pdf';
} else if (sourceExtension === '.docx') {
  extractOperation = 'pdf';
} else if (imageLike.has(sourceExtension)) {
  extractionStatus = 'ocr_required';
  extractionMessage = 'The source file is an image and needs OCR.';
  earlyStatus = 'ocr_required';
  ocrPending = true;
} else {
  extractionStatus = 'unsupported';
  extractionMessage = 'The file extension ' + JSON.stringify(sourceExtension || '<none>') + ' is not supported without OCR.';
  earlyStatus = 'unsupported';
}

return [{
  json: {
    ...item,
    sourceExtension,
    extractOperation,
    extractionStatus,
    extractionMessage,
    earlyStatus,
    ocrPending,
  },
}];
'@

$passIfEarlyExitCode = @'
const item = $input.first();
return item.json.earlyStatus ? [item] : [];
'@

$finalizeEarlyExitCode = @'
const item = $input.first().json;

return [{
  json: {
    status: item.earlyStatus,
    sourceFileUrl: item.source.sourceUrl,
    sourceFileName: item.source.itemName,
    sourceExtension: item.sourceExtension,
    dryRun: item.dryRun,
    extractionStatus: item.extractionStatus,
    extractionMessage: item.extractionMessage,
    extractedTextPreview: null,
    ocrPending: item.ocrPending,
    matchType: null,
    matchValue: null,
    matchMessage: null,
    company: null,
    destination: null,
  },
}];
'@

$passIfExtractableCode = @'
const item = $input.first();
return item.json.earlyStatus ? [] : [item];
'@

$attachOriginalSourceContextCode = @'
const runtime = $('Pass Extractable').first().json;
const downloaded = $input.first();

return [{
  json: {
    ...runtime,
  },
  binary: downloaded.binary,
}];
'@

$passDirectExtractionFilesCode = @'
const item = $input.first();
return ['.txt', '.csv', '.json', '.xml', '.pdf'].includes(item.json.sourceExtension) ? [item] : [];
'@

$passDocxFilesCode = @'
const item = $input.first();
return item.json.sourceExtension === '.docx' ? [item] : [];
'@

$attachDocxPdfContextCode = @'
const runtime = $('Attach Original Source Context').first().json;
const downloaded = $input.first();

return [{
  json: {
    ...runtime,
  },
  binary: downloaded.binary,
}];
'@

$normalizeExtractionCode = @'
const item = $input.first();
const json = item.json;

const candidateText = typeof json.text === 'string'
  ? json.text
  : typeof json.extractedText === 'string'
    ? json.extractedText
    : '';

const extractedText = candidateText.trim();

let extractionStatus = 'native';
let extractionMessage = null;
let ocrPending = false;
let postExtractionStatus = null;

if (!extractedText) {
  if (json.extractOperation === 'pdf') {
    extractionStatus = 'ocr_required';
    extractionMessage = 'The PDF does not expose any native text and needs OCR.';
    ocrPending = true;
    postExtractionStatus = 'ocr_required';
  } else {
    extractionStatus = 'unsupported';
    extractionMessage = 'The source text file is empty.';
    postExtractionStatus = 'unsupported';
  }
}

return [{
  json: {
    ...json,
    extractedText: extractedText || null,
    extractedTextPreview: extractedText ? extractedText.slice(0, 500) : null,
    extractionStatus,
    extractionMessage,
    ocrPending,
    postExtractionStatus,
  },
}];
'@

$passIfPostExtractionExitCode = @'
const item = $input.first();
return item.json.postExtractionStatus ? [item] : [];
'@

$finalizePostExtractionExitCode = @'
const item = $input.first().json;

return [{
  json: {
    status: item.postExtractionStatus,
    sourceFileUrl: item.source.sourceUrl,
    sourceFileName: item.source.itemName,
    sourceExtension: item.sourceExtension,
    dryRun: item.dryRun,
    extractionStatus: item.extractionStatus,
    extractionMessage: item.extractionMessage,
    extractedTextPreview: item.extractedTextPreview,
    ocrPending: item.ocrPending,
    matchType: null,
    matchValue: null,
    matchMessage: null,
    company: null,
    destination: null,
  },
}];
'@

$passIfPostExtractionContinueCode = @'
const item = $input.first();
return item.json.postExtractionStatus ? [] : [item];
'@

$attachWorkbookContextCode = @'
const runtime = $('Pass Post Extraction Continue').first().json;
const downloaded = $input.first();

return [{
  json: {
    ...runtime,
  },
  binary: downloaded.binary,
}];
'@

$normalizeWorkbookCode = @'
const rows = $input.all();
if (!rows.length) {
  throw new Error('The mapping workbook did not produce any rows.');
}

const source = rows[0].json;

const normalizeHeader = (value = '') => String(value)
  .normalize('NFD')
  .replace(/[\u0300-\u036f]/g, '')
  .replace(/[^0-9A-Za-z\s]/g, ' ')
  .replace(/\s+/g, ' ')
  .trim()
  .toUpperCase();

const parseBoolean = (value) => ['SIM', 'YES', 'TRUE', '1'].includes(normalizeHeader(value));
const nullIfWhiteSpace = (value) => {
  const normalized = String(value ?? '').trim();
  return normalized || null;
};

const folderHeaders = [
  'Nome da pasta',
  'Nome pasta',
  'Pasta',
  'Pasta destino',
  'Nome da pasta SharePoint',
  'Nome da pasta no SharePoint',
  'Folder name',
].map(normalizeHeader);

const entries = [];

for (const rowItem of rows) {
  const lookup = {};
  for (const [key, value] of Object.entries(rowItem.json)) {
    lookup[normalizeHeader(key)] = value;
  }

  const acronym = String(lookup['SIGLA'] ?? '').trim();
  const clientName = String(lookup['NOME DO CLIENTE'] ?? '').trim();
  const nipc = String(lookup['NIPC NOSSO CLIENTE'] ?? '').trim();

  if (!acronym || !clientName || !nipc) {
    continue;
  }

  let folderName = null;
  for (const header of folderHeaders) {
    const candidate = nullIfWhiteSpace(lookup[header]);
    if (candidate) {
      folderName = candidate;
      break;
    }
  }

  entries.push({
    acronym,
    clientName,
    nipc,
    hasFolder: parseBoolean(lookup['CLIENTE TEM PASTA']),
    email: nullIfWhiteSpace(lookup['EMAIL']),
    folderName,
  });
}

if (!entries.length) {
  throw new Error('The mapping workbook does not contain any valid company rows.');
}

return [{
  json: {
    source: source.source,
    destinationRoot: source.destinationRoot,
    workbook: source.workbook,
    dryRun: source.dryRun,
    targetFileNameOverride: source.targetFileNameOverride,
    sourceExtension: source.sourceExtension,
    extractOperation: source.extractOperation,
    extractionStatus: source.extractionStatus,
    extractionMessage: source.extractionMessage,
    ocrPending: source.ocrPending,
    extractedText: source.extractedText,
    extractedTextPreview: source.extractedTextPreview,
    companyMappings: entries,
  },
}];
'@

$matchCompanyCode = @'
const item = $input.first().json;
const entries = Array.isArray(item.companyMappings) ? item.companyMappings : [];

if (!entries.length) {
  throw new Error('The mapping workbook does not contain any company rows.');
}

const normalize = (value = '') => String(value)
  .normalize('NFD')
  .replace(/[\u0300-\u036f]/g, '')
  .replace(/\s+/g, ' ')
  .trim()
  .toUpperCase();

const escapeRegex = (value = '') => String(value).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
const fileNameWithoutExtension = String(item.source.itemName || '').replace(/\.[^.]+$/, '');
const candidateText = [item.extractedText || '', fileNameWithoutExtension].filter(Boolean).join('\n');
const normalizedText = normalize(candidateText);

let company = null;
let matchType = 'none';
let matchValue = null;
let matchMessage = null;
let matchStatus = null;

if (!normalizedText) {
  matchMessage = 'The source document did not produce any searchable text.';
  matchStatus = 'unmatched';
} else {
  const nipcMatches = entries.filter((entry) => new RegExp('(^|[^0-9])' + escapeRegex(entry.nipc) + '(?=[^0-9]|$)').test(normalizedText));
  if (nipcMatches.length === 1) {
    company = nipcMatches[0];
    matchType = 'nipc';
    matchValue = company.nipc;
  } else if (nipcMatches.length > 1) {
    matchType = 'nipc';
    matchMessage = 'The document matched multiple NIPC values.';
    matchStatus = 'unmatched';
  }

  if (!company && !matchStatus) {
    const nameMatches = entries
      .filter((entry) => normalizedText.includes(normalize(entry.clientName)))
      .sort((left, right) => right.clientName.length - left.clientName.length);

    if (nameMatches.length === 1) {
      company = nameMatches[0];
      matchType = 'client_name';
      matchValue = company.clientName;
    } else if (nameMatches.length > 1 && nameMatches[0].clientName.length !== nameMatches[1].clientName.length) {
      company = nameMatches[0];
      matchType = 'client_name';
      matchValue = company.clientName;
    } else if (nameMatches.length > 1) {
      matchType = 'client_name';
      matchMessage = 'The document matched multiple client names.';
      matchStatus = 'unmatched';
    }
  }

  if (!company && !matchStatus) {
    const acronymMatches = entries.filter((entry) => new RegExp('(^|[^A-Z0-9])' + escapeRegex(normalize(entry.acronym)) + '(?=[^A-Z0-9]|$)').test(normalizedText));

    if (acronymMatches.length === 1) {
      company = acronymMatches[0];
      matchType = 'acronym';
      matchValue = company.acronym;
    } else if (acronymMatches.length > 1) {
      matchType = 'acronym';
      matchMessage = 'The document matched multiple acronyms.';
      matchStatus = 'unmatched';
    }
  }

  if (!company && !matchStatus) {
    matchMessage = 'The document text did not match any acronym, client name, or NIPC from the workbook.';
    matchStatus = 'unmatched';
  }

  if (company && !company.hasFolder) {
    matchStatus = 'company_has_no_folder';
  }
}

return [{
  json: {
    source: item.source,
    destinationRoot: item.destinationRoot,
    workbook: item.workbook,
    dryRun: item.dryRun,
    targetFileNameOverride: item.targetFileNameOverride,
    sourceExtension: item.sourceExtension,
    extractionStatus: item.extractionStatus,
    extractionMessage: item.extractionMessage,
    ocrPending: item.ocrPending,
    extractedTextPreview: item.extractedTextPreview,
    matchType,
    matchValue,
    matchMessage,
    matchStatus,
    company,
  },
}];
'@

$passIfMatchExitCode = @'
const item = $input.first();
return item.json.matchStatus ? [item] : [];
'@

$finalizeMatchExitCode = @'
const item = $input.first().json;

return [{
  json: {
    status: item.matchStatus,
    sourceFileUrl: item.source.sourceUrl,
    sourceFileName: item.source.itemName,
    sourceExtension: item.sourceExtension,
    dryRun: item.dryRun,
    extractionStatus: item.extractionStatus,
    extractionMessage: item.extractionMessage,
    extractedTextPreview: item.extractedTextPreview,
    ocrPending: item.ocrPending,
    matchType: item.matchType,
    matchValue: item.matchValue,
    matchMessage: item.matchMessage,
    company: item.company,
    destination: null,
  },
}];
'@

$passIfMatchContinueCode = @'
const item = $input.first();
return item.json.matchStatus ? [] : [item];
'@

$resolveDestinationFolderCode = @'
const payload = $input.first().json;
const item = $('Pass Match Continue').first().json;
const children = Array.isArray(payload.value) ? payload.value : [];
const expectedPrefix = item.company.acronym + '_' + item.company.nipc + '_';

const destinationFolder = children
  .filter((child) => !!child.folder)
  .find((child) => String(child.name || '').toLowerCase().startsWith(expectedPrefix.toLowerCase())) || null;

return [{
  json: {
    ...item,
    destinationFolder,
    folderStatus: destinationFolder ? null : 'folder_not_found',
  },
}];
'@

$passIfFolderExitCode = @'
const item = $input.first();
return item.json.folderStatus ? [item] : [];
'@

$finalizeFolderExitCode = @'
const item = $input.first().json;

return [{
  json: {
    status: item.folderStatus,
    sourceFileUrl: item.source.sourceUrl,
    sourceFileName: item.source.itemName,
    sourceExtension: item.sourceExtension,
    dryRun: item.dryRun,
    extractionStatus: item.extractionStatus,
    extractionMessage: item.extractionMessage,
    extractedTextPreview: item.extractedTextPreview,
    ocrPending: item.ocrPending,
    matchType: item.matchType,
    matchValue: item.matchValue,
    matchMessage: item.matchMessage,
    company: item.company,
    destination: null,
  },
}];
'@

$passIfFolderContinueCode = @'
const item = $input.first();
return item.json.folderStatus ? [] : [item];
'@

$buildOutputPlanCode = @'
const item = $input.first().json;
const originalBinaryItem = $('Attach Original Source Context').first();

const firstNonEmpty = (...values) => {
  for (const value of values) {
    const normalized = String(value ?? '').trim();
    if (normalized) {
      return normalized;
    }
  }

  return '';
};

const sanitizeCharacter = (character) => {
  return ['~', '"', '#', '%', '&', '*', ':', '<', '>', '?', '/', '\\', '{', '|', '}'].includes(character)
    ? '_'
    : character;
};

const getExtension = (value = '') => {
  const match = /\.[^.]+$/.exec(String(value));
  return match ? match[0] : '';
};

const getFileNameWithoutExtension = (value = '') => {
  const extension = getExtension(value);
  return extension ? String(value).slice(0, -extension.length) : String(value);
};

const buildDefaultTargetPrefix = (company) => {
  const preferredPrefix = firstNonEmpty(company.folderName, company.clientName, company.acronym + '_' + company.nipc);
  const withUnderscores = preferredPrefix.split(/\s+/).filter(Boolean).join('_');
  const sanitized = [...withUnderscores].map(sanitizeCharacter).join('').replace(/^_+|_+$/g, '');
  return sanitized || (company.acronym + '_' + company.nipc);
};

const buildTargetFileName = (company, sourceFileName, overrideFileName) => {
  const requestedFileName = firstNonEmpty(overrideFileName)
    ? String(overrideFileName).trim()
    : buildDefaultTargetPrefix(company) + '_' + String(sourceFileName);

  const extension = getExtension(requestedFileName);
  let baseName = [...getFileNameWithoutExtension(requestedFileName)].map(sanitizeCharacter).join('').trim();

  if (!baseName) {
    baseName = buildDefaultTargetPrefix(company);
  }

  return baseName + extension;
};

const contentTypeMap = {
  '.pdf': 'application/pdf',
  '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  '.txt': 'text/plain',
  '.csv': 'text/csv',
};

const targetFileName = buildTargetFileName(item.company, item.source.itemName, item.targetFileNameOverride);

return [{
  json: {
    ...item,
    targetFileName,
    sourceContentType: contentTypeMap[item.sourceExtension] || 'application/octet-stream',
    destination: {
      rootFolderUrl: item.destinationRoot.sourceUrl,
      folderName: item.destinationFolder.name,
      folderUrl: item.destinationFolder.webUrl || '',
      targetFileName,
    },
  },
  binary: originalBinaryItem.binary,
}];
'@

$passIfPreviewCode = @'
const item = $input.first();
return item.json.dryRun ? [item] : [];
'@

$finalizePreviewCode = @'
const item = $input.first().json;

return [{
  json: {
    status: 'preview_ready',
    sourceFileUrl: item.source.sourceUrl,
    sourceFileName: item.source.itemName,
    sourceExtension: item.sourceExtension,
    dryRun: item.dryRun,
    extractionStatus: item.extractionStatus,
    extractionMessage: item.extractionMessage,
    extractedTextPreview: item.extractedTextPreview,
    ocrPending: item.ocrPending,
    matchType: item.matchType,
    matchValue: item.matchValue,
    matchMessage: item.matchMessage,
    company: item.company,
    destination: {
      ...item.destination,
      createdFileUrl: null,
    },
  },
}];
'@

$passIfCopyCode = @'
const item = $input.first();
return item.json.dryRun ? [] : [item];
'@

$finalizeCopiedCode = @'
const plan = $('Pass Copy').first().json;
const created = $input.first().json;

return [{
  json: {
    status: 'copied',
    sourceFileUrl: plan.source.sourceUrl,
    sourceFileName: plan.source.itemName,
    sourceExtension: plan.sourceExtension,
    dryRun: plan.dryRun,
    extractionStatus: plan.extractionStatus,
    extractionMessage: plan.extractionMessage,
    extractedTextPreview: plan.extractedTextPreview,
    ocrPending: plan.ocrPending,
    matchType: plan.matchType,
    matchValue: plan.matchValue,
    matchMessage: plan.matchMessage,
    company: plan.company,
    destination: {
      ...plan.destination,
      createdFileUrl: created.webUrl || null,
    },
  },
}];
'@

$nodes = @(
    (New-Node -Id '1' -Name 'Manual Trigger' -Type 'n8n-nodes-base.manualTrigger' -TypeVersion 1 -Parameters @{} -X -1560 -Y 200),
    (New-Node -Id '2' -Name 'Config' -Type 'n8n-nodes-base.code' -TypeVersion 2 -Parameters @{ mode = 'runOnceForAllItems'; language = 'javaScript'; jsCode = $configCode } -X -1320 -Y 200),
    (New-Node -Id '3' -Name 'Get Microsoft Token' -Type 'n8n-nodes-base.httpRequest' -TypeVersion 4.2 -Parameters @{
        method = 'POST'
      url = (New-N8nExpression '''https://login.microsoftonline.com/'' + $json.tenantId + ''/oauth2/v2.0/token''')
        sendBody = $true
        contentType = 'form-urlencoded'
        specifyBody = 'keypair'
        bodyParameters = @{ parameters = @(
            @{ name = 'grant_type'; value = 'client_credentials' },
            @{ name = 'client_id'; value = '={{ $json.clientId }}' },
            @{ name = 'client_secret'; value = '={{ $json.clientSecret }}' },
            @{ name = 'scope'; value = 'https://graph.microsoft.com/.default' }
        ) }
        options = @{ response = @{ response = @{ responseFormat = 'json' } } }
    } -X -1080 -Y 200),
    (New-Node -Id '4' -Name 'Build Resolve Queue' -Type 'n8n-nodes-base.code' -TypeVersion 2 -Parameters @{ mode = 'runOnceForAllItems'; language = 'javaScript'; jsCode = $buildResolveQueueCode } -X -840 -Y 200),
    (New-Node -Id '5' -Name 'Parse SharePoint URL' -Type 'n8n-nodes-base.code' -TypeVersion 2 -Parameters @{ mode = 'runOnceForAllItems'; language = 'javaScript'; jsCode = $parseSharePointUrlCode } -X -600 -Y 200),
    (New-Node -Id '6' -Name 'Resolve Site Candidates' -Type 'n8n-nodes-base.httpRequest' -TypeVersion 4.2 -Parameters @{
        method = 'GET'
      url = (New-N8nExpression '$json.siteResolveUrl')
        sendHeaders = $true
        specifyHeaders = 'keypair'
        headerParameters = @{ parameters = @(
        @{ name = 'Authorization'; value = (New-N8nExpression '''Bearer '' + $(''Get Microsoft Token'').first().json.access_token') }
        ) }
        options = @{ response = @{ response = @{ fullResponse = $true; neverError = $true; responseFormat = 'json' } } }
    } -X -360 -Y 200),
    (New-Node -Id '7' -Name 'Select Resolved Targets' -Type 'n8n-nodes-base.code' -TypeVersion 2 -Parameters @{ mode = 'runOnceForAllItems'; language = 'javaScript'; jsCode = $selectResolvedTargetsCode } -X -120 -Y 200),
    (New-Node -Id '8' -Name 'List Drives' -Type 'n8n-nodes-base.httpRequest' -TypeVersion 4.2 -Parameters @{
        method = 'GET'
      url = (New-N8nExpression '''https://graph.microsoft.com/v1.0/sites/'' + $json.siteId + ''/drives?$select=id,name,webUrl,driveType''')
        sendHeaders = $true
        specifyHeaders = 'keypair'
        headerParameters = @{ parameters = @(
        @{ name = 'Authorization'; value = (New-N8nExpression '''Bearer '' + $(''Get Microsoft Token'').first().json.access_token') }
        ) }
        options = @{ response = @{ response = @{ responseFormat = 'json' } } }
    } -X 120 -Y 200),
    (New-Node -Id '9' -Name 'Pick Drive' -Type 'n8n-nodes-base.code' -TypeVersion 2 -Parameters @{ mode = 'runOnceForAllItems'; language = 'javaScript'; jsCode = $pickDriveCode } -X 360 -Y 200),
    (New-Node -Id '10' -Name 'Resolve Target Items' -Type 'n8n-nodes-base.httpRequest' -TypeVersion 4.2 -Parameters @{
        method = 'GET'
      url = (New-N8nExpression '$json.folderRelativePath ? (''https://graph.microsoft.com/v1.0/drives/'' + $json.driveId + ''/root:/'' + $json.folderRelativePath.split(''/'').map((segment) => encodeURIComponent(segment)).join(''/'')) : (''https://graph.microsoft.com/v1.0/drives/'' + $json.driveId + ''/root'')')
        sendHeaders = $true
        specifyHeaders = 'keypair'
        headerParameters = @{ parameters = @(
        @{ name = 'Authorization'; value = (New-N8nExpression '''Bearer '' + $(''Get Microsoft Token'').first().json.access_token') }
        ) }
        options = @{ response = @{ response = @{ responseFormat = 'json' } } }
    } -X 600 -Y 200),
    (New-Node -Id '11' -Name 'Build Runtime' -Type 'n8n-nodes-base.code' -TypeVersion 2 -Parameters @{ mode = 'runOnceForAllItems'; language = 'javaScript'; jsCode = $buildRuntimeCode } -X 840 -Y 200),
    (New-Node -Id '12' -Name 'Prepare Extraction' -Type 'n8n-nodes-base.code' -TypeVersion 2 -Parameters @{ mode = 'runOnceForAllItems'; language = 'javaScript'; jsCode = $prepareExtractionCode } -X 1080 -Y 200),
    (New-Node -Id '13' -Name 'Pass Early Exit' -Type 'n8n-nodes-base.code' -TypeVersion 2 -Parameters @{ mode = 'runOnceForAllItems'; language = 'javaScript'; jsCode = $passIfEarlyExitCode } -X 1320 -Y 40),
    (New-Node -Id '14' -Name 'Finalize Early Exit' -Type 'n8n-nodes-base.code' -TypeVersion 2 -Parameters @{ mode = 'runOnceForAllItems'; language = 'javaScript'; jsCode = $finalizeEarlyExitCode } -X 1560 -Y 40),
    (New-Node -Id '15' -Name 'Pass Extractable' -Type 'n8n-nodes-base.code' -TypeVersion 2 -Parameters @{ mode = 'runOnceForAllItems'; language = 'javaScript'; jsCode = $passIfExtractableCode } -X 1320 -Y 240),
    (New-Node -Id '16' -Name 'Download Original Source Content' -Type 'n8n-nodes-base.httpRequest' -TypeVersion 4.2 -Parameters @{
        method = 'GET'
      url = (New-N8nExpression '''https://graph.microsoft.com/v1.0/drives/'' + $json.source.driveId + ''/items/'' + $json.source.itemId + ''/content''')
        sendHeaders = $true
        specifyHeaders = 'keypair'
        headerParameters = @{ parameters = @(
        @{ name = 'Authorization'; value = (New-N8nExpression '''Bearer '' + $(''Get Microsoft Token'').first().json.access_token') }
        ) }
        options = @{ response = @{ response = @{ responseFormat = 'file'; outputPropertyName = 'data' } } }
    } -X 1560 -Y 240),
    (New-Node -Id '17' -Name 'Attach Original Source Context' -Type 'n8n-nodes-base.code' -TypeVersion 2 -Parameters @{ mode = 'runOnceForAllItems'; language = 'javaScript'; jsCode = $attachOriginalSourceContextCode } -X 1800 -Y 240),
    (New-Node -Id '18' -Name 'Pass Direct Extraction Files' -Type 'n8n-nodes-base.code' -TypeVersion 2 -Parameters @{ mode = 'runOnceForAllItems'; language = 'javaScript'; jsCode = $passDirectExtractionFilesCode } -X 2040 -Y 120),
    (New-Node -Id '19' -Name 'Extract Direct Native Text' -Type 'n8n-nodes-base.extractFromFile' -TypeVersion 1.1 -Parameters @{
      operation = (New-N8nExpression '$json.extractOperation')
        binaryPropertyName = 'data'
        destinationKey = 'extractedText'
        options = @{ keepSource = 'both'; joinPages = $true; maxPages = 0 }
    } -X 2280 -Y 120),
    (New-Node -Id '20' -Name 'Pass Docx Files' -Type 'n8n-nodes-base.code' -TypeVersion 2 -Parameters @{ mode = 'runOnceForAllItems'; language = 'javaScript'; jsCode = $passDocxFilesCode } -X 2040 -Y 360),
    (New-Node -Id '21' -Name 'Download Source As Pdf' -Type 'n8n-nodes-base.httpRequest' -TypeVersion 4.2 -Parameters @{
        method = 'GET'
      url = (New-N8nExpression '''https://graph.microsoft.com/v1.0/drives/'' + $json.source.driveId + ''/items/'' + $json.source.itemId + ''/content?format=pdf''')
        sendHeaders = $true
        specifyHeaders = 'keypair'
        headerParameters = @{ parameters = @(
        @{ name = 'Authorization'; value = (New-N8nExpression '''Bearer '' + $(''Get Microsoft Token'').first().json.access_token') }
        ) }
        options = @{ response = @{ response = @{ responseFormat = 'file'; outputPropertyName = 'data' } } }
    } -X 2280 -Y 360),
    (New-Node -Id '22' -Name 'Attach Docx Pdf Context' -Type 'n8n-nodes-base.code' -TypeVersion 2 -Parameters @{ mode = 'runOnceForAllItems'; language = 'javaScript'; jsCode = $attachDocxPdfContextCode } -X 2520 -Y 360),
    (New-Node -Id '23' -Name 'Extract Docx Pdf Text' -Type 'n8n-nodes-base.extractFromFile' -TypeVersion 1.1 -Parameters @{
        operation = 'pdf'
        binaryPropertyName = 'data'
        options = @{ keepSource = 'both'; joinPages = $true; maxPages = 0 }
    } -X 2760 -Y 360),
    (New-Node -Id '24' -Name 'Normalize Extraction' -Type 'n8n-nodes-base.code' -TypeVersion 2 -Parameters @{ mode = 'runOnceForAllItems'; language = 'javaScript'; jsCode = $normalizeExtractionCode } -X 3000 -Y 240),
    (New-Node -Id '25' -Name 'Pass Post Extraction Exit' -Type 'n8n-nodes-base.code' -TypeVersion 2 -Parameters @{ mode = 'runOnceForAllItems'; language = 'javaScript'; jsCode = $passIfPostExtractionExitCode } -X 3240 -Y 80),
    (New-Node -Id '26' -Name 'Finalize Post Extraction Exit' -Type 'n8n-nodes-base.code' -TypeVersion 2 -Parameters @{ mode = 'runOnceForAllItems'; language = 'javaScript'; jsCode = $finalizePostExtractionExitCode } -X 3480 -Y 80),
    (New-Node -Id '27' -Name 'Pass Post Extraction Continue' -Type 'n8n-nodes-base.code' -TypeVersion 2 -Parameters @{ mode = 'runOnceForAllItems'; language = 'javaScript'; jsCode = $passIfPostExtractionContinueCode } -X 3240 -Y 240),
    (New-Node -Id '28' -Name 'Download Workbook Binary' -Type 'n8n-nodes-base.httpRequest' -TypeVersion 4.2 -Parameters @{
        method = 'GET'
      url = (New-N8nExpression '''https://graph.microsoft.com/v1.0/drives/'' + $json.workbook.driveId + ''/items/'' + $json.workbook.itemId + ''/content''')
        sendHeaders = $true
        specifyHeaders = 'keypair'
        headerParameters = @{ parameters = @(
        @{ name = 'Authorization'; value = (New-N8nExpression '''Bearer '' + $(''Get Microsoft Token'').first().json.access_token') }
        ) }
        options = @{ response = @{ response = @{ responseFormat = 'file'; outputPropertyName = 'data' } } }
    } -X 3480 -Y 240),
    (New-Node -Id '29' -Name 'Attach Workbook Context' -Type 'n8n-nodes-base.code' -TypeVersion 2 -Parameters @{ mode = 'runOnceForAllItems'; language = 'javaScript'; jsCode = $attachWorkbookContextCode } -X 3720 -Y 240),
    (New-Node -Id '30' -Name 'Extract Workbook Rows' -Type 'n8n-nodes-base.extractFromFile' -TypeVersion 1.1 -Parameters @{
        operation = 'xlsx'
        binaryPropertyName = 'data'
        options = @{ headerRow = $true; includeEmptyCells = $false; keepSource = 'json'; rawData = $false; sheetName = 'Folha1' }
    } -X 3960 -Y 240),
    (New-Node -Id '31' -Name 'Normalize Workbook' -Type 'n8n-nodes-base.code' -TypeVersion 2 -Parameters @{ mode = 'runOnceForAllItems'; language = 'javaScript'; jsCode = $normalizeWorkbookCode } -X 4200 -Y 240),
    (New-Node -Id '32' -Name 'Match Company' -Type 'n8n-nodes-base.code' -TypeVersion 2 -Parameters @{ mode = 'runOnceForAllItems'; language = 'javaScript'; jsCode = $matchCompanyCode } -X 4440 -Y 240),
    (New-Node -Id '33' -Name 'Pass Match Exit' -Type 'n8n-nodes-base.code' -TypeVersion 2 -Parameters @{ mode = 'runOnceForAllItems'; language = 'javaScript'; jsCode = $passIfMatchExitCode } -X 4680 -Y 80),
    (New-Node -Id '34' -Name 'Finalize Match Exit' -Type 'n8n-nodes-base.code' -TypeVersion 2 -Parameters @{ mode = 'runOnceForAllItems'; language = 'javaScript'; jsCode = $finalizeMatchExitCode } -X 4920 -Y 80),
    (New-Node -Id '35' -Name 'Pass Match Continue' -Type 'n8n-nodes-base.code' -TypeVersion 2 -Parameters @{ mode = 'runOnceForAllItems'; language = 'javaScript'; jsCode = $passIfMatchContinueCode } -X 4680 -Y 240),
    (New-Node -Id '36' -Name 'List Destination Children' -Type 'n8n-nodes-base.httpRequest' -TypeVersion 4.2 -Parameters @{
        method = 'GET'
      url = (New-N8nExpression '''https://graph.microsoft.com/v1.0/drives/'' + $json.destinationRoot.driveId + ''/items/'' + $json.destinationRoot.itemId + ''/children?$select=id,name,webUrl,folder,file''')
        sendHeaders = $true
        specifyHeaders = 'keypair'
        headerParameters = @{ parameters = @(
        @{ name = 'Authorization'; value = (New-N8nExpression '''Bearer '' + $(''Get Microsoft Token'').first().json.access_token') }
        ) }
        options = @{ response = @{ response = @{ responseFormat = 'json' } } }
    } -X 4920 -Y 240),
    (New-Node -Id '37' -Name 'Resolve Destination Folder' -Type 'n8n-nodes-base.code' -TypeVersion 2 -Parameters @{ mode = 'runOnceForAllItems'; language = 'javaScript'; jsCode = $resolveDestinationFolderCode } -X 5160 -Y 240),
    (New-Node -Id '38' -Name 'Pass Folder Exit' -Type 'n8n-nodes-base.code' -TypeVersion 2 -Parameters @{ mode = 'runOnceForAllItems'; language = 'javaScript'; jsCode = $passIfFolderExitCode } -X 5400 -Y 80),
    (New-Node -Id '39' -Name 'Finalize Folder Exit' -Type 'n8n-nodes-base.code' -TypeVersion 2 -Parameters @{ mode = 'runOnceForAllItems'; language = 'javaScript'; jsCode = $finalizeFolderExitCode } -X 5640 -Y 80),
    (New-Node -Id '40' -Name 'Pass Folder Continue' -Type 'n8n-nodes-base.code' -TypeVersion 2 -Parameters @{ mode = 'runOnceForAllItems'; language = 'javaScript'; jsCode = $passIfFolderContinueCode } -X 5400 -Y 240),
    (New-Node -Id '41' -Name 'Build Output Plan' -Type 'n8n-nodes-base.code' -TypeVersion 2 -Parameters @{ mode = 'runOnceForAllItems'; language = 'javaScript'; jsCode = $buildOutputPlanCode } -X 5640 -Y 240),
    (New-Node -Id '42' -Name 'Pass Preview' -Type 'n8n-nodes-base.code' -TypeVersion 2 -Parameters @{ mode = 'runOnceForAllItems'; language = 'javaScript'; jsCode = $passIfPreviewCode } -X 5880 -Y 120),
    (New-Node -Id '43' -Name 'Finalize Preview' -Type 'n8n-nodes-base.code' -TypeVersion 2 -Parameters @{ mode = 'runOnceForAllItems'; language = 'javaScript'; jsCode = $finalizePreviewCode } -X 6120 -Y 120),
    (New-Node -Id '44' -Name 'Pass Copy' -Type 'n8n-nodes-base.code' -TypeVersion 2 -Parameters @{ mode = 'runOnceForAllItems'; language = 'javaScript'; jsCode = $passIfCopyCode } -X 5880 -Y 360),
    (New-Node -Id '45' -Name 'Upload File' -Type 'n8n-nodes-base.httpRequest' -TypeVersion 4.2 -Parameters @{
        method = 'PUT'
      url = (New-N8nExpression '''https://graph.microsoft.com/v1.0/drives/'' + $json.destinationRoot.driveId + ''/items/'' + $json.destinationFolder.id + '':/'' + encodeURIComponent($json.targetFileName) + '':/content''')
        sendHeaders = $true
        specifyHeaders = 'keypair'
        headerParameters = @{ parameters = @(
        @{ name = 'Authorization'; value = (New-N8nExpression '''Bearer '' + $(''Get Microsoft Token'').first().json.access_token') },
        @{ name = 'Content-Type'; value = (New-N8nExpression '$json.sourceContentType') }
        ) }
        sendBody = $true
        contentType = 'binaryData'
        inputDataFieldName = 'data'
        options = @{ response = @{ response = @{ responseFormat = 'json' } } }
    } -X 6120 -Y 360),
    (New-Node -Id '46' -Name 'Finalize Copied' -Type 'n8n-nodes-base.code' -TypeVersion 2 -Parameters @{ mode = 'runOnceForAllItems'; language = 'javaScript'; jsCode = $finalizeCopiedCode } -X 6360 -Y 360)
)

$connections = [ordered]@{
  'Manual Trigger' = @{ main = , @((New-Connection 'Config')) }
  'Config' = @{ main = , @((New-Connection 'Get Microsoft Token')) }
  'Get Microsoft Token' = @{ main = , @((New-Connection 'Build Resolve Queue')) }
  'Build Resolve Queue' = @{ main = , @((New-Connection 'Parse SharePoint URL')) }
  'Parse SharePoint URL' = @{ main = , @((New-Connection 'Resolve Site Candidates')) }
  'Resolve Site Candidates' = @{ main = , @((New-Connection 'Select Resolved Targets')) }
  'Select Resolved Targets' = @{ main = , @((New-Connection 'List Drives')) }
  'List Drives' = @{ main = , @((New-Connection 'Pick Drive')) }
  'Pick Drive' = @{ main = , @((New-Connection 'Resolve Target Items')) }
  'Resolve Target Items' = @{ main = , @((New-Connection 'Build Runtime')) }
  'Build Runtime' = @{ main = , @((New-Connection 'Prepare Extraction')) }
  'Prepare Extraction' = @{ main = , @(
      (New-Connection 'Pass Early Exit'),
      (New-Connection 'Pass Extractable')
  ) }
  'Pass Early Exit' = @{ main = , @((New-Connection 'Finalize Early Exit')) }
  'Pass Extractable' = @{ main = , @((New-Connection 'Download Original Source Content')) }
  'Download Original Source Content' = @{ main = , @((New-Connection 'Attach Original Source Context')) }
  'Attach Original Source Context' = @{ main = , @(
      (New-Connection 'Pass Direct Extraction Files'),
      (New-Connection 'Pass Docx Files')
  ) }
  'Pass Direct Extraction Files' = @{ main = , @((New-Connection 'Extract Direct Native Text')) }
  'Extract Direct Native Text' = @{ main = , @((New-Connection 'Normalize Extraction')) }
  'Pass Docx Files' = @{ main = , @((New-Connection 'Download Source As Pdf')) }
  'Download Source As Pdf' = @{ main = , @((New-Connection 'Attach Docx Pdf Context')) }
  'Attach Docx Pdf Context' = @{ main = , @((New-Connection 'Extract Docx Pdf Text')) }
  'Extract Docx Pdf Text' = @{ main = , @((New-Connection 'Normalize Extraction')) }
  'Normalize Extraction' = @{ main = , @(
      (New-Connection 'Pass Post Extraction Exit'),
      (New-Connection 'Pass Post Extraction Continue')
  ) }
  'Pass Post Extraction Exit' = @{ main = , @((New-Connection 'Finalize Post Extraction Exit')) }
  'Pass Post Extraction Continue' = @{ main = , @((New-Connection 'Download Workbook Binary')) }
  'Download Workbook Binary' = @{ main = , @((New-Connection 'Attach Workbook Context')) }
  'Attach Workbook Context' = @{ main = , @((New-Connection 'Extract Workbook Rows')) }
  'Extract Workbook Rows' = @{ main = , @((New-Connection 'Normalize Workbook')) }
  'Normalize Workbook' = @{ main = , @((New-Connection 'Match Company')) }
  'Match Company' = @{ main = , @(
      (New-Connection 'Pass Match Exit'),
      (New-Connection 'Pass Match Continue')
  ) }
  'Pass Match Exit' = @{ main = , @((New-Connection 'Finalize Match Exit')) }
  'Pass Match Continue' = @{ main = , @((New-Connection 'List Destination Children')) }
  'List Destination Children' = @{ main = , @((New-Connection 'Resolve Destination Folder')) }
  'Resolve Destination Folder' = @{ main = , @(
      (New-Connection 'Pass Folder Exit'),
      (New-Connection 'Pass Folder Continue')
  ) }
  'Pass Folder Exit' = @{ main = , @((New-Connection 'Finalize Folder Exit')) }
  'Pass Folder Continue' = @{ main = , @((New-Connection 'Build Output Plan')) }
  'Build Output Plan' = @{ main = , @(
      (New-Connection 'Pass Preview'),
      (New-Connection 'Pass Copy')
  ) }
  'Pass Preview' = @{ main = , @((New-Connection 'Finalize Preview')) }
  'Pass Copy' = @{ main = , @((New-Connection 'Upload File')) }
  'Upload File' = @{ main = , @((New-Connection 'Finalize Copied')) }
}

$workflow = [ordered]@{
    name        = 'SharePoint Route Document (Generic Core Nodes)'
    nodes       = $nodes
    pinData     = @{}
    connections = $connections
    active      = $false
    settings    = @{ executionOrder = 'v1' }
    tags        = @()
}

$json = $workflow | ConvertTo-Json -Depth 100
Set-Content -Path $outputPath -Value $json -Encoding UTF8

Write-Host "Wrote $outputPath"