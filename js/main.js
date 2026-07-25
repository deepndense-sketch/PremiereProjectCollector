const csInterface = new CSInterface();
const fs = require('fs');
const os = require('os');
const path = require('path');
const https = require('https');
const crypto = require('crypto');
const childProcess = require('child_process');
const { spawn } = childProcess;

let destination = null;
let compareLocation = null;
let isCopying = false;
let hostScriptReady = false;
let latestPlan = null;
let sourceTree = null;
let compareFiles = [];
let compareScanErrors = [];
let listVisible = false;
let selectionTouched = false;
let localVersion = 'unknown';
let remoteVersion = null;
let selectedSequenceFilters = [];
let trackPresets = [];
let trackPresetsLoaded = false;
let trackPresetIdCounter = 0;
let sequenceOnlyMode = false;
let createReducedProject = false;
let copyProjectFile = true;
let linkProjectAfterCollection = false;
let includedProjectFolders = [];
let ignoredProjectFolders = [];
let trackRangeAnchor = null;
let folderRangeAnchor = null;
let trackClickTimer = null;
let folderClickTimer = null;
let trackConflictPromptResolver = null;
let trackConflictPromptState = null;

const PLUGIN_STORAGE_PREFIX = 'projectcollector.';
const SEQUENCE_FILTERS_STORAGE_KEY = 'projectcollector.sequenceFilters';
const TRACK_PRESETS_STORAGE_KEY = 'projectcollector.trackPresets';
const DESTINATION_STORAGE_KEY = 'projectcollector.destination';
const IGNORE_SECTION_VISIBLE_STORAGE_KEY = 'projectcollector.ignoreSectionVisible';
const SEQUENCE_ONLY_MODE_STORAGE_KEY = 'projectcollector.sequenceOnlyMode';
const CREATE_REDUCED_PROJECT_STORAGE_KEY = 'projectcollector.createReducedProject';
const COPY_PROJECT_FILE_STORAGE_KEY = 'projectcollector.copyProjectFile';
const LINK_PROJECT_AFTER_COLLECTION_STORAGE_KEY = 'projectcollector.linkProjectAfterCollection';

function setText(id, value) {
    document.getElementById(id).textContent = value;
}

function setIgnoreSectionVisibility(visible) {
    const content = document.getElementById('ignoreSectionContent');
    const button = document.getElementById('toggleIgnoreButton');
    content.style.display = visible ? 'block' : 'none';
    button.textContent = visible ? 'Hide' : 'Show';

    try {
        localStorage.setItem(IGNORE_SECTION_VISIBLE_STORAGE_KEY, visible ? '1' : '0');
    } catch (error) {}
}

function toggleIgnoreSection() {
    const content = document.getElementById('ignoreSectionContent');
    setIgnoreSectionVisibility(content.style.display === 'none');
}

function setProjectFolderSectionVisibility(visible) {
    const content = document.getElementById('projectFolderSectionContent');
    const button = document.getElementById('toggleProjectFoldersButton');

    if (!content || !button) {
        return;
    }

    content.style.display = visible ? 'block' : 'none';
    button.textContent = visible ? 'Hide' : 'Show';
}

function toggleProjectFolderSection() {
    const content = document.getElementById('projectFolderSectionContent');
    setProjectFolderSectionVisibility(content && content.style.display === 'none');
}

function setSourceSectionVisibility(visible) {
    const content = document.getElementById('sourceSectionContent');
    const button = document.getElementById('toggleSourceSectionButton');

    if (!content || !button) {
        return;
    }

    content.style.display = visible ? 'block' : 'none';
    button.textContent = visible ? 'Hide' : 'Show';
}

function toggleSourceSection() {
    const content = document.getElementById('sourceSectionContent');
    setSourceSectionVisibility(content && content.style.display === 'none');
}

function syncSequenceModeUI() {
    const sequenceOnlyCheckbox = document.getElementById('sequenceOnlyMode');
    const reducedProjectCheckbox = document.getElementById('createReducedProject');

    if (sequenceOnlyCheckbox) {
        sequenceOnlyCheckbox.checked = sequenceOnlyMode;
    }

    if (reducedProjectCheckbox) {
        reducedProjectCheckbox.checked = createReducedProject;
        reducedProjectCheckbox.disabled = !sequenceOnlyMode;
    }
}

function syncProjectOptionUI() {
    const copyProjectCheckbox = document.getElementById('copyProjectFile');
    const linkProjectCheckbox = document.getElementById('linkProjectAfterCollection');

    if (copyProjectCheckbox) {
        copyProjectCheckbox.checked = copyProjectFile;
    }

    if (linkProjectCheckbox) {
        linkProjectCheckbox.checked = linkProjectAfterCollection && copyProjectFile;
        linkProjectCheckbox.disabled = !copyProjectFile;
    }
}

function toggleSequenceOnlyMode() {
    sequenceOnlyMode = !!document.getElementById('sequenceOnlyMode').checked;
    if (!sequenceOnlyMode) {
        createReducedProject = false;
    }

    try {
        localStorage.setItem(SEQUENCE_ONLY_MODE_STORAGE_KEY, sequenceOnlyMode ? '1' : '0');
        localStorage.setItem(CREATE_REDUCED_PROJECT_STORAGE_KEY, createReducedProject ? '1' : '0');
    } catch (error) {}

    syncSequenceModeUI();
    updateSelectionSummary();
}

function toggleCreateReducedProject() {
    createReducedProject = sequenceOnlyMode && !!document.getElementById('createReducedProject').checked;

    try {
        localStorage.setItem(CREATE_REDUCED_PROJECT_STORAGE_KEY, createReducedProject ? '1' : '0');
    } catch (error) {}

    syncSequenceModeUI();
}

function toggleCopyProjectFile() {
    copyProjectFile = !!document.getElementById('copyProjectFile').checked;
    if (!copyProjectFile) {
        linkProjectAfterCollection = false;
    }

    try {
        localStorage.setItem(COPY_PROJECT_FILE_STORAGE_KEY, copyProjectFile ? '1' : '0');
        localStorage.setItem(LINK_PROJECT_AFTER_COLLECTION_STORAGE_KEY, linkProjectAfterCollection ? '1' : '0');
    } catch (error) {}

    syncProjectOptionUI();
}

function toggleLinkProjectAfterCollection() {
    linkProjectAfterCollection = copyProjectFile && !!document.getElementById('linkProjectAfterCollection').checked;

    try {
        localStorage.setItem(LINK_PROJECT_AFTER_COLLECTION_STORAGE_KEY, linkProjectAfterCollection ? '1' : '0');
    } catch (error) {}

    syncProjectOptionUI();
}

function getExtensionRootPath() {
    try {
        return csInterface.getSystemPath(SystemPath.EXTENSION);
    } catch (error) {
        return __dirname;
    }
}

function getVersionFilePath() {
    return path.join(getExtensionRootPath(), 'version.json');
}

function getUpdateScriptPath() {
    return path.join(getExtensionRootPath(), 'update_from_github.ps1');
}

function getTempUpdaterScriptPath() {
    return path.join(os.tmpdir(), 'PremiereProjectCollector_update_launch.ps1');
}

function getTempUpdaterZipPath() {
    return path.join(os.tmpdir(), 'PremiereProjectCollector_update_package.zip');
}

function getTempUpdaterResultPath() {
    return path.join(os.tmpdir(), 'PremiereProjectCollector_update_result.json');
}

function getTempUpdaterLogPath() {
    return path.join(os.tmpdir(), 'PremiereProjectCollector_update_log.txt');
}

function getUserCepExtensionPath() {
    return path.join(process.env.APPDATA || '', 'Adobe', 'CEP', 'extensions', 'PremiereProjectCollector');
}

function getInstalledExtensionPath() {
    return getExtensionRootPath() || getUserCepExtensionPath();
}

function fileExists(filePath) {
    try {
        return !!filePath && fs.existsSync(filePath);
    } catch (error) {
        return false;
    }
}

function readVersionInfo() {
    try {
        const raw = fs.readFileSync(getVersionFilePath(), 'utf8');
        const parsed = JSON.parse(raw);
        localVersion = parsed.version || 'unknown';
    } catch (error) {
        localVersion = 'unknown';
    }

    return localVersion;
}

function compareVersions(a, b) {
    const aParts = String(a || '0').split('.').map((part) => parseInt(part, 10) || 0);
    const bParts = String(b || '0').split('.').map((part) => parseInt(part, 10) || 0);
    const length = Math.max(aParts.length, bParts.length);

    for (let i = 0; i < length; i += 1) {
        const left = aParts[i] || 0;
        const right = bParts[i] || 0;
        if (left > right) {
            return 1;
        }
        if (left < right) {
            return -1;
        }
    }

    return 0;
}

function setUpdateButton(label, isUpdateAvailable) {
    const button = document.getElementById('updateButton');
    button.textContent = label;
    button.disabled = isCopying || !isUpdateAvailable;
    button.classList.toggle('button-update-ready', isUpdateAvailable);
}

async function checkForUpdates() {
    const remoteUrl = `https://raw.githubusercontent.com/deepndense-sketch/PremiereProjectCollector/main/version.json?ts=${Date.now()}`;
    setUpdateButton(`Version ${localVersion}`, false);

    try {
        const remote = await new Promise((resolve, reject) => {
            const request = https.get(remoteUrl, {
                headers: {
                    'User-Agent': 'PremiereProjectCollector-Updater',
                    'Cache-Control': 'no-cache',
                    Pragma: 'no-cache'
                }
            }, (response) => {
                if (response.statusCode !== 200) {
                    reject(new Error(`HTTP ${response.statusCode}`));
                    response.resume();
                    return;
                }

                let raw = '';
                response.setEncoding('utf8');
                response.on('data', (chunk) => {
                    raw += chunk;
                });
                response.on('end', () => {
                    try {
                        resolve(JSON.parse(raw));
                    } catch (error) {
                        reject(error);
                    }
                });
            });

            request.on('error', reject);
        });

        remoteVersion = remote.version || 'unknown';
        if (compareVersions(remoteVersion, localVersion) > 0) {
            setUpdateButton(`Update to ${remoteVersion}`, true);
        } else {
            setUpdateButton(`Version ${localVersion}`, false);
        }
    } catch (error) {
        remoteVersion = null;
        setUpdateButton(`Version ${localVersion}`, false);
    }
}

function delay(ms) {
    return new Promise((resolve) => {
        setTimeout(resolve, ms);
    });
}

function downloadFile(url, destinationPath) {
    return new Promise((resolve, reject) => {
        const file = fs.createWriteStream(destinationPath);
        const request = https.get(url, (response) => {
            if (response.statusCode >= 300 && response.statusCode < 400 && response.headers.location) {
                file.close(() => {
                    fs.unlink(destinationPath, () => {
                        downloadFile(response.headers.location, destinationPath).then(resolve).catch(reject);
                    });
                });
                return;
            }

            if (response.statusCode !== 200) {
                file.close(() => {
                    fs.unlink(destinationPath, () => {});
                    reject(new Error(`HTTP ${response.statusCode}`));
                });
                response.resume();
                return;
            }

            response.pipe(file);
            file.on('finish', () => {
                file.close(resolve);
            });
        });

        request.on('error', (error) => {
            file.close(() => {
                fs.unlink(destinationPath, () => {});
                reject(error);
            });
        });

        file.on('error', (error) => {
            file.close(() => {
                fs.unlink(destinationPath, () => {});
                reject(error);
            });
        });
    });
}

async function monitorUpdaterCompletion() {
    const maxAttempts = 10;
    const resultPath = getTempUpdaterResultPath();

    for (let attempt = 0; attempt < maxAttempts; attempt += 1) {
        await delay(3000);

        if (fileExists(resultPath)) {
            try {
                const raw = fs.readFileSync(resultPath, 'utf8');
                const parsed = JSON.parse(raw);
                if (parsed.ok) {
                    readVersionInfo();
                    await checkForUpdates();
                    setText('summaryText', `Update complete. Installed version: ${localVersion}. Restart Premiere Pro if the panel is already open.`);
                    return;
                }

                setText('summaryText', `Updater failed. ${parsed.message || 'Unknown error.'} Log: ${parsed.logPath || getTempUpdaterLogPath()}`);
                return;
            } catch (error) {
                setText('summaryText', `Updater finished, but the result file could not be read. ${error.message}`);
                return;
            }
        }

        readVersionInfo();
        await checkForUpdates();

        if (remoteVersion && compareVersions(remoteVersion, localVersion) <= 0) {
            setText('summaryText', `Update complete. Installed version: ${localVersion}. Restart Premiere Pro if the panel is already open.`);
            return;
        }
    }

    setText('summaryText', `Updater finished launching, but this panel still sees version ${localVersion}. Reopen the panel or restart Premiere Pro and check again.`);
}

function runGithubUpdate() {
    if (isCopying) {
        return;
    }

    const updateScriptPath = getUpdateScriptPath();
    if (!fileExists(updateScriptPath)) {
        setText('summaryText', 'Update script was not found.');
        return;
    }

    if (remoteVersion && compareVersions(remoteVersion, localVersion) <= 0) {
        setText('summaryText', `Version ${localVersion} is already installed.`);
        return;
    }

    const tempUpdaterScriptPath = getTempUpdaterScriptPath();
    const tempUpdaterZipPath = getTempUpdaterZipPath();
    const tempUpdaterResultPath = getTempUpdaterResultPath();
    const tempUpdaterLogPath = getTempUpdaterLogPath();
    const remoteZipUrl = 'https://github.com/deepndense-sketch/PremiereProjectCollector/archive/refs/heads/main.zip';

    setText('summaryText', 'Downloading update package from GitHub...');

    try {
        fs.copyFileSync(updateScriptPath, tempUpdaterScriptPath);
        if (fileExists(tempUpdaterZipPath)) {
            fs.unlinkSync(tempUpdaterZipPath);
        }
        if (fileExists(tempUpdaterResultPath)) {
            fs.unlinkSync(tempUpdaterResultPath);
        }
        if (fileExists(tempUpdaterLogPath)) {
            fs.unlinkSync(tempUpdaterLogPath);
        }
    } catch (error) {
        setText('summaryText', `Could not prepare updater. ${error.message}`);
        return;
    }

    downloadFile(remoteZipUrl, tempUpdaterZipPath)
        .then(() => {
            const updateDestination = getInstalledExtensionPath();
            const updaterArguments = [
                '-NoExit',
                '-NoProfile',
                '-ExecutionPolicy',
                'Bypass',
                '-File',
                quoteForWindowsArgument(tempUpdaterScriptPath),
                '-ZipPath',
                quoteForWindowsArgument(tempUpdaterZipPath),
                '-Destination',
                quoteForWindowsArgument(updateDestination),
                '-ResultPath',
                quoteForWindowsArgument(tempUpdaterResultPath),
                '-LogPath',
                quoteForWindowsArgument(tempUpdaterLogPath)
            ].join(' ');
            const escapedArgumentList = escapeForPowerShellSingleQuotedString(updaterArguments);
            const command = `Start-Process -FilePath 'powershell.exe' -Verb RunAs -ArgumentList '${escapedArgumentList}'`;

            childProcess.execFile(
                'powershell.exe',
                ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', command],
                (error) => {
                    if (error) {
                        setText('summaryText', `Could not launch updater. ${error.message}`);
                        return;
                    }

                    setText('summaryText', `Updater launched for ${updateDestination}. Accept the Windows prompt if it appears.`);
                    monitorUpdaterCompletion();
                }
            );
        })
        .catch((error) => {
            setText('summaryText', `Could not prepare updater. ${error.message}`);
        });
}

function escapeForEvalScript(value) {
    return String(value)
        .replace(/\\/g, "\\\\")
        .replace(/"/g, '\\"')
        .replace(/\r/g, "\\r")
        .replace(/\n/g, "\\n");
}

function quoteForWindowsArgument(value) {
    return `"${String(value).replace(/"/g, '\\"')}"`;
}

function escapeForPowerShellSingleQuotedString(value) {
    return String(value).replace(/'/g, "''");
}

function buildSequenceFiltersPayload() {
    return selectedSequenceFilters.map((filter) => ({
        sequenceID: filter.sequenceID || '',
        sequenceName: filter.sequenceName || '',
        ignoredVideoTracks: filter.ignoredVideoTracks || [],
        ignoredAudioTracks: filter.ignoredAudioTracks || []
    }));
}

function sanitizeFileName(name, fallback) {
    const cleaned = String(name || fallback || 'Untitled')
        .replace(/[\\/:*?"<>|]/g, '_')
        .replace(/^\s+/, '')
        .replace(/\s+$/, '');
    return cleaned || fallback || 'Untitled';
}

function getCollectedProjectFileName(projectPath) {
    const extension = path.extname(projectPath || '') || '.prproj';
    const baseName = path.basename(projectPath || 'Premiere_Project', extension);

    return `${sanitizeFileName(`${baseName} BACKUP`, 'Premiere_Project BACKUP')}${extension}`;
}

async function copyProjectFileIntoCollectedRoot(rootPath, projectPath, destinationName) {
    if (!rootPath || !projectPath) {
        return { success: false, message: 'Project path was not available.' };
    }

    const destinationPath = path.join(rootPath, sanitizeFileName(destinationName || path.basename(projectPath), path.basename(projectPath)));
    const result = await copyFileWithRobocopy(projectPath, destinationPath);
    result.destinationPath = destinationPath;
    return result;
}

function callHost(script) {
    return new Promise((resolve) => {
        csInterface.evalScript(script, (result) => {
            resolve(result);
        });
    });
}

function safeJsonParse(raw) {
    try {
        return JSON.parse(raw);
    } catch (error) {
        return null;
    }
}

function ensureDirectorySync(dirPath) {
    fs.mkdirSync(dirPath, { recursive: true });
}

function formatBytes(size) {
    const value = Number(size);

    if (!Number.isFinite(value) || value < 0) {
        return 'unknown size';
    }

    if (value < 1024) {
        return `${value} B`;
    }

    const units = ['KB', 'MB', 'GB', 'TB'];
    let unitIndex = -1;
    let current = value;

    do {
        current /= 1024;
        unitIndex += 1;
    } while (current >= 1024 && unitIndex < units.length - 1);

    return `${current.toFixed(current >= 10 ? 1 : 2)} ${units[unitIndex]}`;
}

function normalizePathForTree(filePath) {
    return String(filePath || '').replace(/\//g, '\\');
}

function tryResolveWindowsPath(filePath) {
    const normalized = normalizePathForTree(filePath);

    try {
        if (fs.realpathSync && typeof fs.realpathSync.native === 'function') {
            return normalizePathForTree(fs.realpathSync.native(normalized));
        }
    } catch (error) {}

    try {
        return normalizePathForTree(fs.realpathSync(normalized));
    } catch (error2) {}

    return normalized;
}

function getCompareFileName(filePath) {
    return path.basename(String(filePath || '')).toLowerCase();
}

function buildCompareSizeKey(fileName, size) {
    return `${fileName}|${size}`;
}

function getSourceCompareSignature(filePath) {
    const normalized = normalizePathForTree(filePath);
    const fileName = getCompareFileName(normalized);
    let size = null;

    try {
        const stat = fs.statSync(normalized);
        if (stat && stat.isFile()) {
            size = stat.size;
        }
    } catch (error) {}

    return {
        fileName,
        size
    };
}

function buildMediaSignatureKey(filePath) {
    const signature = getSourceCompareSignature(filePath);

    if (!signature.fileName || signature.size === null) {
        return '';
    }

    return buildCompareSizeKey(signature.fileName, signature.size);
}

function scanCompareFiles(rootPath) {
    const root = normalizePathForTree(rootPath);
    const files = [];
    const errors = [];
    const visitedDirs = new Set();

    if (!root) {
        return {
            files,
            errors: ['No skip location selected.'],
            blocked: true
        };
    }

    try {
        const rootStat = fs.statSync(root);
        if (!rootStat.isDirectory()) {
            return {
                files,
                errors: [`Compare location is not a folder: ${root}`],
                blocked: true
            };
        }
    } catch (error) {
        return {
            files,
            errors: [`Could not open compare location: ${root}`],
            blocked: true
        };
    }

    const walk = (folderPath) => {
        let folderKey = '';

        try {
            folderKey = tryResolveWindowsPath(folderPath).toLowerCase();
        } catch (error) {
            folderKey = normalizePathForTree(folderPath).toLowerCase();
        }

        if (visitedDirs.has(folderKey)) {
            return;
        }
        visitedDirs.add(folderKey);

        let names = [];
        try {
            names = fs.readdirSync(folderPath);
        } catch (error) {
            errors.push(`Could not read folder: ${folderPath}`);
            return;
        }

        names.forEach((name) => {
            const fullPath = path.join(folderPath, name);
            let stat = null;

            try {
                stat = fs.lstatSync(fullPath);
            } catch (error) {
                errors.push(`Could not read file info: ${fullPath}`);
                return;
            }

            if (stat.isSymbolicLink()) {
                return;
            }

            if (stat.isDirectory()) {
                walk(fullPath);
                return;
            }

            if (stat.isFile()) {
                files.push({
                    name: path.basename(fullPath),
                    path: fullPath,
                    size: stat.size
                });
            }
        });
    };

    walk(root);
    files.sort((left, right) => left.name.localeCompare(right.name) || left.path.localeCompare(right.path));

    return {
        files,
        errors,
        blocked: false
    };
}

function buildCompareLookup(files) {
    const byNameAndSize = new Map();

    (files || []).forEach((file) => {
        const fileName = getCompareFileName(file.name || file.path);

        if (!fileName) {
            return;
        }

        if (Number.isFinite(Number(file.size))) {
            const key = buildCompareSizeKey(fileName, Number(file.size));
            if (!byNameAndSize.has(key)) {
                byNameAndSize.set(key, []);
            }
            byNameAndSize.get(key).push(file);
        }
    });

    return {
        byNameAndSize,
        hashByPath: new Map()
    };
}

function getFileSha256(filePath, hashCache) {
    const cache = hashCache || new Map();
    const cacheKey = normalizeMediaKey(filePath);
    if (cache.has(cacheKey)) {
        return cache.get(cacheKey);
    }

    const hashPromise = new Promise((resolve, reject) => {
        const hash = crypto.createHash('sha256');
        const stream = fs.createReadStream(filePath);

        stream.on('error', reject);
        stream.on('data', (chunk) => hash.update(chunk));
        stream.on('end', () => resolve(hash.digest('hex')));
    });
    cache.set(cacheKey, hashPromise);
    hashPromise.catch(() => {
        if (cache.get(cacheKey) === hashPromise) {
            cache.delete(cacheKey);
        }
    });
    return hashPromise;
}

async function findCompareMatchForTask(task, lookup) {
    if (!task || !lookup) {
        return null;
    }

    const signature = getSourceCompareSignature(task.source);

    if (!signature.fileName) {
        return null;
    }

    if (signature.size === null) {
        return null;
    }

    const candidates = lookup.byNameAndSize.get(buildCompareSizeKey(signature.fileName, signature.size)) || [];
    if (!candidates.length) {
        return null;
    }

    let sourceHash = '';
    try {
        sourceHash = await getFileSha256(task.source, lookup.hashByPath);
    } catch (error) {
        return null;
    }

    for (let index = 0; index < candidates.length; index += 1) {
        const candidate = candidates[index];
        try {
            const candidateHash = await getFileSha256(candidate.path, lookup.hashByPath);
            if (candidateHash === sourceHash) {
                return Object.assign({}, candidate, {
                    sha256: candidateHash
                });
            }
        } catch (error) {}
    }

    return null;
}

async function inspectCompareLocation() {
    if (!compareLocation) {
        return {
            files: [],
            errors: [],
            blocked: false,
            lookup: buildCompareLookup([])
        };
    }

    await new Promise((resolve) => setTimeout(resolve, 0));

    const scan = scanCompareFiles(compareLocation);
    compareFiles = scan.files || [];
    compareScanErrors = scan.errors || [];

    return Object.assign({}, scan, {
        lookup: buildCompareLookup(scan.files)
    });
}

function splitSourcePath(filePath) {
    const normalized = normalizePathForTree(filePath);
    const resolved = tryResolveWindowsPath(normalized);
    const parts = resolved.split('\\').filter(Boolean);

    if (resolved.startsWith('\\\\') && parts.length >= 2) {
        return {
            drive: `\\\\${parts[0]}\\${parts[1]}`,
            segments: parts.slice(2),
            displayPath: normalized
        };
    }

    if (/^[A-Za-z]:\\/.test(resolved)) {
        return {
            drive: resolved.slice(0, 2).toUpperCase(),
            segments: parts.slice(1),
            displayPath: normalized
        };
    }

    return {
        drive: 'Unknown',
        segments: parts,
        displayPath: normalized
    };
}

function createTreeNode(key, name, type, fullPath) {
    return {
        key,
        name,
        type,
        fullPath,
        selected: true,
        explicit: false,
        expanded: false,
        children: []
    };
}

function buildSourceTree(tasks) {
    const driveMap = new Map();

    tasks.forEach((task, taskIndex) => {
        const parsed = splitSourcePath(task.source);
        let driveNode = driveMap.get(parsed.drive);

        if (!driveNode) {
            driveNode = createTreeNode(`drive:${parsed.drive}`, parsed.drive, 'drive', parsed.drive);
            driveNode.childMap = new Map();
            driveMap.set(parsed.drive, driveNode);
        }

        let parentNode = driveNode;
        let parentKey = driveNode.key;
        let fullPath = parsed.drive;

        parsed.segments.forEach((segment, segmentIndex) => {
            const isLast = segmentIndex === parsed.segments.length - 1;
            const type = isLast ? 'file' : 'folder';
            const key = `${parentKey}>${segment.toLowerCase()}`;

            if (!parentNode.childMap) {
                parentNode.childMap = new Map();
            }

            let childNode = parentNode.childMap.get(key);
            fullPath = parentNode.fullPath.startsWith('\\\\')
                ? `${parentNode.fullPath}\\${segment}`
                : `${fullPath}\\${segment}`;

            if (!childNode) {
                childNode = createTreeNode(key, segment, type, type === 'file' ? parsed.displayPath : fullPath);
                if (type !== 'file') {
                    childNode.childMap = new Map();
                } else {
                    childNode.taskIndexes = [];
                }
                parentNode.childMap.set(key, childNode);
                parentNode.children.push(childNode);
            }

            parentNode = childNode;
        });

        if (parentNode.type === 'file') {
            parentNode.taskIndexes.push(taskIndex);
        }
    });

    function finalize(node) {
        node.children.sort((left, right) => {
            if (left.type !== right.type) {
                return left.type === 'file' ? 1 : -1;
            }
            return left.name.localeCompare(right.name);
        });

        node.children.forEach(finalize);
        delete node.childMap;
    }

    const roots = Array.from(driveMap.values()).sort((left, right) => left.name.localeCompare(right.name));
    roots.forEach(finalize);
    return roots;
}

function visitTree(nodes, handler) {
    nodes.forEach((node) => {
        handler(node);
        if (node.children.length) {
            visitTree(node.children, handler);
        }
    });
}

function applySelectionToNode(node, selected, explicit) {
    node.selected = selected;
    node.explicit = explicit;
    node.children.forEach((child) => {
        applySelectionToNode(child, selected, explicit);
    });
}

function syncNodeFromChildren(node) {
    if (!node.children.length) {
        return node.selected;
    }

    const childStates = node.children.map(syncNodeFromChildren);
    const allSelected = childStates.every(Boolean);
    const allDeselected = childStates.every((value) => !value);

    if (allSelected) {
        node.selected = true;
    } else if (allDeselected) {
        node.selected = false;
    }

    node.explicit = false;
    return node.selected;
}

function updateSelectionSummary() {
    if (!latestPlan || !sourceTree) {
        setText('selectionSummary', 'Loading project files from Premiere...');
        return;
    }

    const total = latestPlan.tasks.length;
    const included = getSelectedTasks().length;
    const ignoredTrackSummary = selectedSequenceFilters
        .map((filter) => {
            const parts = [];

            if (filter.ignoredVideoTracks.length) {
                parts.push(`video ${filter.ignoredVideoTracks.map((trackNumber) => `V${trackNumber}`).join(', ')}`);
            }

            if (filter.ignoredAudioTracks.length) {
                parts.push(`audio ${filter.ignoredAudioTracks.map((trackNumber) => `A${trackNumber}`).join(', ')}`);
            }

            return parts.length ? `${filter.sequenceName}: ${parts.join(' | ')}` : '';
        })
        .filter(Boolean);

    const modeSummary = [];

    if (sequenceOnlyMode) {
        modeSummary.push(`selected sequences only: ${selectedSequenceFilters.map((filter) => filter.sequenceName).join(', ')}`);
    }

    if (ignoredTrackSummary.length) {
        modeSummary.push(ignoredTrackSummary.join(' || '));
    }

    if (includedProjectFolders.length) {
        modeSummary.push(`force-copy Premiere folders: ${includedProjectFolders.join(', ')}`);
    }

    if (ignoredProjectFolders.length) {
        modeSummary.push(`ignored Premiere folders: ${ignoredProjectFolders.join(', ')}`);
    }

    const suffix = modeSummary.length ? ` (${modeSummary.join(' || ')})` : '';
    const treeScopeNote = (sequenceOnlyMode || ignoredTrackSummary.length || includedProjectFolders.length || ignoredProjectFolders.length)
        ? ' Source File List still shows the full project tree; final copy is adjusted later by sequence mode, track choices, and Premiere folder rules.'
        : '';

    if (!selectionTouched) {
        setText('selectionSummary', `All ${total} files will be included by default. Once you change the list, only the checked items will be copied.${suffix}${treeScopeNote}`);
        return;
    }

    if (included === 0) {
        setText('selectionSummary', `No files are selected. Copy will process zero files until you check items again.${suffix}${treeScopeNote}`);
        return;
    }

    setText('selectionSummary', `${included} of ${total} files are selected for copy.${suffix}${treeScopeNote}`);
}

function getSelectedTaskIndexSet() {
    const selectedIndexes = new Set();

    if (!sourceTree) {
        return selectedIndexes;
    }

    visitTree(sourceTree, (node) => {
        if (node.type === 'file' && node.selected && Array.isArray(node.taskIndexes)) {
            node.taskIndexes.forEach((taskIndex) => selectedIndexes.add(taskIndex));
        }
    });

    return selectedIndexes;
}

function getSelectedTasks() {
    if (!latestPlan || !sourceTree) {
        return [];
    }

    if (!selectionTouched) {
        return latestPlan.tasks.slice();
    }

    const selectedIndexes = getSelectedTaskIndexSet();
    return latestPlan.tasks.filter((task, index) => selectedIndexes.has(index));
}

function normalizeProjectFolderPath(folderPath) {
    return String(folderPath || '')
        .replace(/\\/g, '/')
        .replace(/^\/+/, '')
        .replace(/\/+$/, '');
}

function isRootProjectFolderPath(folderPath) {
    const normalized = normalizeProjectFolderPath(folderPath);
    return !!normalized && normalized.indexOf('/') === -1;
}

function isTaskInsideProjectFolder(task, folderPath) {
    const taskFolder = normalizeProjectFolderPath(task && task.binPath);
    const ignoredFolder = normalizeProjectFolderPath(folderPath);

    if (!ignoredFolder) {
        return false;
    }

    return taskFolder === ignoredFolder || taskFolder.indexOf(`${ignoredFolder}/`) === 0;
}

function isTaskInsideProjectFolderList(task, folderPaths) {
    return (folderPaths || []).some((folderPath) => isTaskInsideProjectFolder(task, folderPath));
}

function isTaskInsideIgnoredProjectFolder(task) {
    return isTaskInsideProjectFolderList(task, ignoredProjectFolders);
}

function isTaskInsideIncludedProjectFolder(task) {
    return isTaskInsideProjectFolderList(task, includedProjectFolders);
}

function splitRelativeDestination(relativePath) {
    const normalized = String(relativePath || '').replace(/\\/g, '/');
    const slashIndex = normalized.lastIndexOf('/');

    if (slashIndex === -1) {
        return {
            folder: '',
            fileName: normalized
        };
    }

    return {
        folder: normalized.slice(0, slashIndex),
        fileName: normalized.slice(slashIndex + 1)
    };
}

function makeUniqueRelativeDestination(relativePath, usedDestinations) {
    const parsed = splitRelativeDestination(relativePath);
    const extension = path.extname(parsed.fileName);
    const baseName = parsed.fileName.slice(0, parsed.fileName.length - extension.length) || 'Media';
    let candidate = `${parsed.folder ? `${parsed.folder}/` : ''}${parsed.fileName}`;
    let key = candidate.toLowerCase();
    let suffix = 123;

    while (usedDestinations.has(key)) {
        const nextFileName = `${baseName}_${suffix}${extension}`;
        candidate = `${parsed.folder ? `${parsed.folder}/` : ''}${nextFileName}`;
        key = candidate.toLowerCase();
        suffix += 1;
    }

    usedDestinations.add(key);
    return candidate;
}

function getRootProjectFolderNames() {
    const folders = latestPlan && Array.isArray(latestPlan.folders) ? latestPlan.folders : [];
    return new Set(folders
        .map(normalizeProjectFolderPath)
        .filter(isRootProjectFolderPath)
        .map((folderPath) => folderPath.toLowerCase()));
}

function chooseLooseMediaFolderName(rootPath) {
    const existingProjectFolders = getRootProjectFolderNames();
    const baseName = 'CollectedMedias';

    function isAvailable(name) {
        if (existingProjectFolders.has(name.toLowerCase())) {
            return false;
        }

        try {
            return !rootPath || !fs.existsSync(path.join(rootPath, name));
        } catch (error) {
            return true;
        }
    }

    if (isAvailable(baseName)) {
        return baseName;
    }

    for (let suffix = 123; suffix < 1000; suffix += 1) {
        const candidate = `${baseName}_${suffix}`;
        if (isAvailable(candidate)) {
            return candidate;
        }
    }

    return `${baseName}_${Date.now()}`;
}

function getRootLooseFileName(task) {
    const destinationFile = splitRelativeDestination(task.destination || '').fileName;
    if (destinationFile) {
        return destinationFile;
    }

    return sanitizeFileName(path.basename(task.source || task.name || 'Media'), 'Media');
}

function buildCollectedTasks(tasks, rootPath) {
    const looseMediaFolderName = chooseLooseMediaFolderName(rootPath);
    const usedDestinations = new Set();

    return tasks.map((task) => {
        const nextTask = Object.assign({}, task);
        const taskFolder = normalizeProjectFolderPath(task.binPath);

        if (!taskFolder) {
            nextTask.destination = `${looseMediaFolderName}/${getRootLooseFileName(task)}`;
            nextTask.binPath = looseMediaFolderName;
        }

        nextTask.destination = makeUniqueRelativeDestination(nextTask.destination, usedDestinations);
        nextTask.relativePath = nextTask.destination;
        return nextTask;
    });
}

function normalizeMediaKey(filePath) {
    return tryResolveWindowsPath(filePath).toLowerCase();
}

function sanitizeTrackUsageEntries(entries, prefix) {
    return Array.isArray(entries)
        ? entries.map((entry) => ({
            trackNumber: parseInt(entry.trackNumber, 10) || 0,
            label: entry.label || `${prefix}${entry.trackNumber}`,
            clipCount: parseInt(entry.clipCount, 10) || 0,
            mediaPaths: Array.isArray(entry.mediaPaths) ? entry.mediaPaths.slice() : [],
            hasItems: !!entry.hasItems || (parseInt(entry.clipCount, 10) || 0) > 0 || (Array.isArray(entry.mediaPaths) && entry.mediaPaths.length > 0)
        })).filter((entry) => entry.trackNumber > 0)
        : [];
}

function createSequenceFilter(sequenceID, sequenceName, videoTrackUsage, audioTrackUsage, locked) {
    return {
        sequenceID: sequenceID || '',
        sequenceName: sequenceName || 'Unknown Sequence',
        videoTrackUsage: sanitizeTrackUsageEntries(videoTrackUsage, 'V'),
        audioTrackUsage: sanitizeTrackUsageEntries(audioTrackUsage, 'A'),
        ignoredVideoTracks: [],
        ignoredAudioTracks: [],
        selectedPresetId: '',
        locked: !!locked
    };
}

function sanitizeSequenceFilter(rawFilter, locked) {
    const filter = createSequenceFilter(rawFilter.sequenceID, rawFilter.sequenceName, rawFilter.videoTrackUsage, rawFilter.audioTrackUsage, locked);

    filter.ignoredVideoTracks = Array.isArray(rawFilter.ignoredVideoTracks)
        ? rawFilter.ignoredVideoTracks.map((value) => parseInt(value, 10) || 0).filter((value) => value > 0)
        : [];
    filter.ignoredAudioTracks = Array.isArray(rawFilter.ignoredAudioTracks)
        ? rawFilter.ignoredAudioTracks.map((value) => parseInt(value, 10) || 0).filter((value) => value > 0)
        : [];
    filter.selectedPresetId = typeof rawFilter.selectedPresetId === 'string'
        ? rawFilter.selectedPresetId
        : '';

    filter.ignoredVideoTracks = Array.from(new Set(filter.ignoredVideoTracks)).sort((a, b) => a - b);
    filter.ignoredAudioTracks = Array.from(new Set(filter.ignoredAudioTracks)).sort((a, b) => a - b);
    if (filter.selectedPresetId && !trackPresets.some((preset) => preset.id === filter.selectedPresetId)) {
        filter.selectedPresetId = '';
    }
    filter.locked = !!locked;
    return filter;
}

function getDefaultSequenceFilter() {
    if (!latestPlan || !latestPlan.activeSequenceName) {
        return null;
    }

    return createSequenceFilter(
        latestPlan.activeSequenceID || '',
        latestPlan.activeSequenceName,
        latestPlan.videoTrackUsage || [],
        latestPlan.audioTrackUsage || [],
        false
    );
}

function sanitizePresetTrackNumbers(values) {
    return Array.from(new Set(
        (Array.isArray(values) ? values : [])
            .map((value) => parseInt(value, 10) || 0)
            .filter((value) => value > 0)
    )).sort((left, right) => left - right);
}

function sanitizeTrackPreset(rawPreset, index) {
    if (!rawPreset || !String(rawPreset.name || '').trim()) {
        return null;
    }

    const name = String(rawPreset.name).trim();
    const fallbackId = `preset-${index + 1}-${name.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '') || 'track'}`;
    return {
        id: String(rawPreset.id || fallbackId),
        name,
        ignoredVideoTracks: sanitizePresetTrackNumbers(rawPreset.ignoredVideoTracks),
        ignoredAudioTracks: sanitizePresetTrackNumbers(rawPreset.ignoredAudioTracks)
    };
}

function loadTrackPresets() {
    let parsedPresets = [];

    try {
        const parsed = JSON.parse(localStorage.getItem(TRACK_PRESETS_STORAGE_KEY) || '[]');
        parsedPresets = Array.isArray(parsed) ? parsed : [];
    } catch (error) {
        parsedPresets = [];
    }

    const seenIds = new Set();
    const seenNames = new Set();
    trackPresets = parsedPresets.reduce((result, rawPreset, index) => {
        const preset = sanitizeTrackPreset(rawPreset, index);
        const normalizedName = preset ? preset.name.toLowerCase() : '';
        if (!preset || seenIds.has(preset.id) || seenNames.has(normalizedName)) {
            return result;
        }
        seenIds.add(preset.id);
        seenNames.add(normalizedName);
        result.push(preset);
        return result;
    }, []);
    trackPresetsLoaded = true;
}

function saveTrackPresets() {
    try {
        localStorage.setItem(TRACK_PRESETS_STORAGE_KEY, JSON.stringify(trackPresets));
    } catch (error) {}
}

function clearStoredPluginSessionData() {
    const keysToRemove = [];

    try {
        for (let index = 0; index < localStorage.length; index += 1) {
            const key = localStorage.key(index);
            if (
                key
                && key.indexOf(PLUGIN_STORAGE_PREFIX) === 0
                && key !== TRACK_PRESETS_STORAGE_KEY
                && key !== DESTINATION_STORAGE_KEY
            ) {
                keysToRemove.push(key);
            }
        }
        keysToRemove.forEach((key) => localStorage.removeItem(key));
    } catch (error) {}

    return keysToRemove;
}

function deletePluginRelatedData() {
    const shouldDelete = typeof window.confirm !== 'function'
        || window.confirm('Delete saved project, sequence, track-selection, and panel-option data? Track presets and the saved destination will be kept.');
    if (!shouldDelete) {
        return false;
    }

    const removedKeys = clearStoredPluginSessionData();
    const status = document.getElementById('dataCleanupStatus');
    if (status) {
        status.textContent = removedKeys.length
            ? `Deleted ${removedKeys.length} saved data item${removedKeys.length === 1 ? '' : 's'}.`
            : 'No saved project data found.';
    }
    return true;
}

function getTrackPresetById(presetId) {
    return trackPresets.find((preset) => preset.id === presetId) || null;
}

function getNextTrackPresetName(presets) {
    const names = new Set(
        (Array.isArray(presets) ? presets : trackPresets)
            .map((preset) => String(preset && preset.name || '').trim().toLowerCase())
            .filter(Boolean)
    );
    let number = 1;
    while (names.has(`project copy preset ${number}`)) {
        number += 1;
    }
    return `Project Copy Preset ${number}`;
}

function trackNumberListsMatch(left, right) {
    const normalizedLeft = sanitizePresetTrackNumbers(left);
    const normalizedRight = sanitizePresetTrackNumbers(right);
    return normalizedLeft.length === normalizedRight.length
        && normalizedLeft.every((value, index) => value === normalizedRight[index]);
}

function findTrackPresetByFilterSettings(filter) {
    if (!filter) {
        return null;
    }

    return trackPresets.find((preset) => (
        trackNumberListsMatch(preset.ignoredVideoTracks, filter.ignoredVideoTracks)
        && trackNumberListsMatch(preset.ignoredAudioTracks, filter.ignoredAudioTracks)
    )) || null;
}

function createTrackPresetId() {
    let presetId = '';
    do {
        trackPresetIdCounter += 1;
        presetId = `track-preset-${Date.now()}-${trackPresetIdCounter}`;
    } while (getTrackPresetById(presetId));
    return presetId;
}

function upsertTrackPresetFromFilter(filter, requestedName) {
    if (!filter) {
        return null;
    }

    const name = String(requestedName || '').trim();
    if (!name) {
        return null;
    }

    const existing = trackPresets.find((preset) => preset.name.toLowerCase() === name.toLowerCase());
    const preset = existing || {
        id: createTrackPresetId(),
        name,
        ignoredVideoTracks: [],
        ignoredAudioTracks: []
    };

    preset.name = name;
    preset.ignoredVideoTracks = sanitizePresetTrackNumbers(filter.ignoredVideoTracks);
    preset.ignoredAudioTracks = sanitizePresetTrackNumbers(filter.ignoredAudioTracks);
    if (!existing) {
        trackPresets.push(preset);
    }
    return preset;
}

function getSequenceTrackNumbers(filter, kind) {
    const usage = kind === 'video' ? filter.videoTrackUsage : filter.audioTrackUsage;
    return new Set(
        (Array.isArray(usage) ? usage : [])
            .map((entry) => parseInt(entry && entry.trackNumber, 10) || 0)
            .filter((trackNumber) => trackNumber > 0)
    );
}

function applyTrackPresetToFilter(filter, preset) {
    if (!filter || !preset) {
        return false;
    }

    const videoTracks = getSequenceTrackNumbers(filter, 'video');
    const audioTracks = getSequenceTrackNumbers(filter, 'audio');
    filter.ignoredVideoTracks = preset.ignoredVideoTracks.filter((trackNumber) => videoTracks.has(trackNumber));
    filter.ignoredAudioTracks = preset.ignoredAudioTracks.filter((trackNumber) => audioTracks.has(trackNumber));
    filter.selectedPresetId = preset.id;
    return true;
}

function applyTrackPresetToSequence(sequenceKey, presetId) {
    const filter = getSequenceFilterByKey(sequenceKey);
    if (!filter) {
        return false;
    }

    if (!presetId) {
        filter.selectedPresetId = '';
        saveAndRenderTrackFilters();
        return true;
    }

    const preset = getTrackPresetById(presetId);
    if (!applyTrackPresetToFilter(filter, preset)) {
        return false;
    }

    if (trackRangeAnchor && trackRangeAnchor.sequenceKey === sequenceKey) {
        trackRangeAnchor = null;
    }
    saveAndRenderTrackFilters();
    return true;
}

function saveTrackPresetForSequence(sequenceKey) {
    const filter = getSequenceFilterByKey(sequenceKey);
    if (!filter) {
        return;
    }

    const duplicatePreset = findTrackPresetByFilterSettings(filter);
    if (duplicatePreset) {
        applyTrackPresetToFilter(filter, duplicatePreset);
        saveSequenceFilters();
        renderSequenceFilters();
        updateSelectionSummary();
        alert(`This track setting already exists as "${duplicatePreset.name}".`);
        return;
    }

    const defaultName = getNextTrackPresetName(trackPresets);
    const requestedName = window.prompt('Preset name', defaultName);
    if (requestedName === null) {
        return;
    }

    const preset = upsertTrackPresetFromFilter(filter, String(requestedName).trim() || defaultName);
    if (!preset) {
        return;
    }

    selectedSequenceFilters.forEach((sequenceFilter) => {
        if (sequenceFilter === filter || sequenceFilter.selectedPresetId === preset.id) {
            applyTrackPresetToFilter(sequenceFilter, preset);
        }
    });
    trackRangeAnchor = null;
    saveTrackPresets();
    saveSequenceFilters();
    renderSequenceFilters();
    updateSelectionSummary();
}

function deleteTrackPreset(presetId) {
    const preset = getTrackPresetById(presetId);
    if (!preset) {
        return false;
    }

    trackPresets = trackPresets.filter((candidate) => candidate.id !== presetId);
    selectedSequenceFilters.forEach((filter) => {
        if (filter.selectedPresetId === presetId) {
            filter.selectedPresetId = '';
        }
    });
    saveTrackPresets();
    saveSequenceFilters();
    return true;
}

function deleteSelectedTrackPresetForSequence(sequenceKey) {
    const filter = getSequenceFilterByKey(sequenceKey);
    const preset = filter ? getTrackPresetById(filter.selectedPresetId) : null;
    if (!preset) {
        return;
    }

    if (typeof window.confirm === 'function' && !window.confirm(`Delete preset "${preset.name}"?`)) {
        return;
    }

    if (deleteTrackPreset(preset.id)) {
        renderSequenceFilters();
        updateSelectionSummary();
    }
}

function loadSequenceFilters() {
    if (!trackPresetsLoaded) {
        loadTrackPresets();
    }

    let savedFilters = [];
    let hasStoredFilters = false;

    const storageKey = getSequenceFiltersStorageKey();

    try {
        const storedFilters = localStorage.getItem(storageKey);
        hasStoredFilters = storedFilters !== null;
        const parsed = JSON.parse(storedFilters || '[]');
        savedFilters = Array.isArray(parsed) ? parsed : [];
    } catch (error) {
        savedFilters = [];
    }

    const availableSequences = latestPlan && Array.isArray(latestPlan.availableSequences)
        ? latestPlan.availableSequences
        : null;
    const hadSavedSequenceEntries = savedFilters.some((filter) => filter && filter.sequenceName);

    selectedSequenceFilters = savedFilters
        .filter((filter) => filter && filter.sequenceName)
        .filter((filter) => !availableSequences || isSequenceFilterAvailable(filter, availableSequences))
        .map((filter) => sanitizeSequenceFilter(filter, false));

    const defaultFilter = getDefaultSequenceFilter();
    if (!defaultFilter) {
        selectedSequenceFilters.forEach((filter) => {
            const preset = getTrackPresetById(filter.selectedPresetId);
            if (preset) {
                applyTrackPresetToFilter(filter, preset);
            }
        });
        saveSequenceFilters();
        renderSequenceFilters();
        return;
    }

    if (!selectedSequenceFilters.length) {
        if (!hasStoredFilters || hadSavedSequenceEntries) {
            selectedSequenceFilters = [defaultFilter];
            saveSequenceFilters();
        }
        renderSequenceFilters();
        return;
    }

    const defaultIndex = selectedSequenceFilters.findIndex((filter) => filter.sequenceID === defaultFilter.sequenceID || filter.sequenceName === defaultFilter.sequenceName);
    if (defaultIndex >= 0) {
        const existing = selectedSequenceFilters[defaultIndex];
        existing.videoTrackUsage = defaultFilter.videoTrackUsage;
        existing.audioTrackUsage = defaultFilter.audioTrackUsage;
        existing.sequenceID = defaultFilter.sequenceID || existing.sequenceID;
    }

    selectedSequenceFilters = selectedSequenceFilters.map((filter) => sanitizeSequenceFilter(filter, false));
    selectedSequenceFilters.forEach((filter) => {
        const preset = getTrackPresetById(filter.selectedPresetId);
        if (preset) {
            applyTrackPresetToFilter(filter, preset);
        }
    });
    saveSequenceFilters();
    renderSequenceFilters();
}

function saveSequenceFilters() {
    try {
        localStorage.setItem(getSequenceFiltersStorageKey(), JSON.stringify(selectedSequenceFilters));
    } catch (error) {}
}

function getSequenceFiltersStorageKey(plan) {
    const projectPlan = plan || latestPlan || {};
    const projectIdentity = projectPlan.projectPath
        ? normalizeMediaKey(projectPlan.projectPath)
        : `unsaved:${String(projectPlan.projectName || 'premiere-project').toLowerCase()}`;
    return `${SEQUENCE_FILTERS_STORAGE_KEY}:${projectIdentity}`;
}

function isSequenceFilterAvailable(filter, availableSequences) {
    const sequences = Array.isArray(availableSequences) ? availableSequences : [];
    if (!filter || !sequences.length) {
        return false;
    }

    if (filter.sequenceID) {
        return sequences.some((sequence) => sequence && sequence.sequenceID === filter.sequenceID);
    }

    const sequenceName = String(filter.sequenceName || '').toLowerCase();
    return sequences.some((sequence) => (
        sequence
        && sequenceName
        && String(sequence.sequenceName || '').toLowerCase() === sequenceName
    ));
}

function getSequenceFilterByKey(sequenceKey) {
    return selectedSequenceFilters.find((filter) => (filter.sequenceID || filter.sequenceName) === sequenceKey) || null;
}

function getSequenceFilterKey(filter) {
    return filter.sequenceID || filter.sequenceName;
}

function getVisibleTrackEntries(filter, kind) {
    const entries = kind === 'video' ? filter.videoTrackUsage : filter.audioTrackUsage;
    return entries
        .filter((entry) => entry.hasItems || entry.clipCount > 0 || (entry.mediaPaths || []).length > 0)
        .sort((left, right) => left.trackNumber - right.trackNumber);
}

function getIgnoredTrackList(filter, kind) {
    return kind === 'video' ? filter.ignoredVideoTracks : filter.ignoredAudioTracks;
}

function setIgnoredTrackList(filter, kind, values) {
    const uniqueValues = Array.from(new Set(values.map((value) => parseInt(value, 10) || 0).filter((value) => value > 0))).sort((a, b) => a - b);

    if (kind === 'video') {
        filter.ignoredVideoTracks = uniqueValues;
    } else {
        filter.ignoredAudioTracks = uniqueValues;
    }
}

function isTrackIgnored(filter, kind, trackNumber) {
    return getIgnoredTrackList(filter, kind).indexOf(trackNumber) !== -1;
}

function saveAndRenderTrackFilters() {
    saveSequenceFilters();
    renderSequenceFilters();
}

function clearTrackChoices(filters, sequenceKey) {
    const targetKey = sequenceKey || '';
    let resetCount = 0;

    (Array.isArray(filters) ? filters : []).forEach((filter) => {
        if (!targetKey || getSequenceFilterKey(filter) === targetKey) {
            filter.ignoredVideoTracks = [];
            filter.ignoredAudioTracks = [];
            filter.selectedPresetId = '';
            resetCount += 1;
        }
    });

    return resetCount;
}

function findSequenceTrackUsage(filter, usageEntries) {
    const entries = Array.isArray(usageEntries) ? usageEntries : [];
    if (filter.sequenceID) {
        const idMatch = entries.find((entry) => entry && entry.sequenceID === filter.sequenceID);
        if (idMatch) {
            return idMatch;
        }
    }

    const normalizedName = String(filter.sequenceName || '').toLowerCase();
    return entries.find((entry) => (
        entry
        && normalizedName
        && String(entry.sequenceName || '').toLowerCase() === normalizedName
    )) || null;
}

function applySequenceTrackUsagePlan(filters, plan, resetChoices) {
    const usageEntries = plan && Array.isArray(plan.sequences) ? plan.sequences : [];
    const result = {
        refreshedCount: 0,
        missingSequences: [],
        resetCount: 0
    };

    (Array.isArray(filters) ? filters : []).forEach((filter) => {
        if (resetChoices) {
            filter.ignoredVideoTracks = [];
            filter.ignoredAudioTracks = [];
            filter.selectedPresetId = '';
            result.resetCount += 1;
        }

        const incoming = findSequenceTrackUsage(filter, usageEntries);
        if (!incoming) {
            result.missingSequences.push(filter.sequenceName || filter.sequenceID || 'Unknown Sequence');
            return;
        }

        mergeSequenceTrackUsage(filter, createSequenceFilter(
            incoming.sequenceID || filter.sequenceID,
            incoming.sequenceName || filter.sequenceName,
            incoming.videoTrackUsage || [],
            incoming.audioTrackUsage || [],
            filter.locked
        ));
        result.refreshedCount += 1;
    });

    return result;
}

function toggleIgnoredTrack(sequenceKey, kind, trackNumber) {
    const filter = getSequenceFilterByKey(sequenceKey);
    if (!filter || !trackNumber) {
        return;
    }

    const ignored = getIgnoredTrackList(filter, kind).slice();
    const existingIndex = ignored.indexOf(trackNumber);
    trackRangeAnchor = null;

    if (existingIndex === -1) {
        ignored.push(trackNumber);
    } else {
        ignored.splice(existingIndex, 1);
    }

    filter.selectedPresetId = '';
    setIgnoredTrackList(filter, kind, ignored);
    saveAndRenderTrackFilters();
}

function ignoreTrackRange(anchor, filter, kind, trackNumber) {
    if (!anchor || anchor.sequenceKey !== getSequenceFilterKey(filter) || anchor.kind !== kind) {
        trackRangeAnchor = {
            sequenceKey: getSequenceFilterKey(filter),
            kind,
            trackNumber
        };
        renderSequenceFilters();
        return;
    }

    const ignored = getIgnoredTrackList(filter, kind).slice();
    const start = Math.min(anchor.trackNumber, trackNumber);
    const end = Math.max(anchor.trackNumber, trackNumber);

    for (let current = start; current <= end; current += 1) {
        if (ignored.indexOf(current) === -1) {
            ignored.push(current);
        }
    }

    trackRangeAnchor = null;
    filter.selectedPresetId = '';
    setIgnoredTrackList(filter, kind, ignored);
    saveAndRenderTrackFilters();
}

function handleTrackButtonClick(filter, kind, trackNumber) {
    if (trackClickTimer) {
        clearTimeout(trackClickTimer);
    }

    trackClickTimer = setTimeout(() => {
        toggleIgnoredTrack(getSequenceFilterKey(filter), kind, trackNumber);
        trackClickTimer = null;
    }, 220);
}

function handleTrackButtonDoubleClick(filter, kind, trackNumber) {
    if (trackClickTimer) {
        clearTimeout(trackClickTimer);
        trackClickTimer = null;
    }

    ignoreTrackRange(trackRangeAnchor, filter, kind, trackNumber);
}

function renderTrackButtonGroup(filter, kind) {
    const entries = getVisibleTrackEntries(filter, kind);
    const list = document.createElement('div');
    list.className = 'choice-grid track-choice-grid';

    if (!entries.length) {
        const empty = document.createElement('div');
        empty.className = 'small-note';
        empty.textContent = kind === 'video' ? 'No video tracks with clips.' : 'No audio tracks with clips.';
        list.appendChild(empty);
        return list;
    }

    entries.forEach((entry) => {
        const item = document.createElement('div');
        const button = document.createElement('button');
        const sequenceKey = getSequenceFilterKey(filter);
        const isAnchor = trackRangeAnchor
            && trackRangeAnchor.sequenceKey === sequenceKey
            && trackRangeAnchor.kind === kind
            && trackRangeAnchor.trackNumber === entry.trackNumber;
        const ignored = isTrackIgnored(filter, kind, entry.trackNumber);

        item.className = 'choice-item track-choice-item';
        button.type = 'button';
        button.className = `choice-button track-choice-button${ignored ? ' is-ignored' : ''}${isAnchor ? ' is-range-anchor' : ''}`;
        button.title = `${kind === 'video' ? 'Video' : 'Audio'} track ${entry.trackNumber}: ${entry.clipCount} ${entry.clipCount === 1 ? 'clip' : 'clips'}`;
        button.onclick = (event) => {
            if (event.detail === 1) {
                handleTrackButtonClick(filter, kind, entry.trackNumber);
            }
        };
        button.ondblclick = () => handleTrackButtonDoubleClick(filter, kind, entry.trackNumber);

        const title = document.createElement('span');
        title.className = 'choice-title';
        title.textContent = `${kind === 'video' ? 'V' : 'A'}${entry.trackNumber}`;
        button.appendChild(title);

        item.appendChild(button);

        list.appendChild(item);
    });

    return list;
}

function renderSequenceGroup(filter, kind) {
    const group = document.createElement('div');
    group.className = 'sequence-group';

    const title = document.createElement('div');
    title.className = 'label';
    title.textContent = kind === 'video' ? 'Video Tracks' : 'Audio Tracks';
    group.appendChild(title);

    group.appendChild(renderTrackButtonGroup(filter, kind));

    return group;
}

function renderSequencePresetSection(filter) {
    const sequenceKey = getSequenceFilterKey(filter);
    const section = document.createElement('div');
    section.className = 'sequence-preset-section';

    const select = document.createElement('select');
    select.className = 'sequence-preset-select';
    select.setAttribute('aria-label', `Track preset for ${filter.sequenceName}`);

    const customOption = document.createElement('option');
    customOption.value = '';
    customOption.textContent = 'Custom tracks';
    select.appendChild(customOption);

    trackPresets.forEach((preset) => {
        const option = document.createElement('option');
        option.value = preset.id;
        option.textContent = preset.name;
        select.appendChild(option);
    });

    select.value = getTrackPresetById(filter.selectedPresetId) ? filter.selectedPresetId : '';
    select.onchange = () => applyTrackPresetToSequence(sequenceKey, select.value);
    section.appendChild(select);

    const saveButton = document.createElement('button');
    saveButton.type = 'button';
    saveButton.className = 'button-secondary button-small sequence-preset-save';
    saveButton.textContent = 'Save preset';
    saveButton.onclick = () => saveTrackPresetForSequence(sequenceKey);
    section.appendChild(saveButton);

    const deleteButton = document.createElement('button');
    deleteButton.type = 'button';
    deleteButton.className = 'button-danger-soft button-small sequence-preset-delete';
    deleteButton.textContent = 'Delete';
    deleteButton.disabled = !getTrackPresetById(filter.selectedPresetId);
    deleteButton.onclick = () => deleteSelectedTrackPresetForSequence(sequenceKey);
    section.appendChild(deleteButton);

    const resetButton = document.createElement('button');
    resetButton.type = 'button';
    resetButton.className = 'button-secondary button-small sequence-reset-button';
    resetButton.textContent = 'Reset tracks';
    resetButton.onclick = () => resetTrackSelection(sequenceKey);
    section.appendChild(resetButton);

    return section;
}

function renderSequenceFilters() {
    const container = document.getElementById('sequenceFilters');
    const hint = document.getElementById('sequenceFilterHint');
    container.innerHTML = '';

    if (!selectedSequenceFilters.length) {
        const empty = document.createElement('div');
        empty.className = 'small-note';
        empty.textContent = 'No sequences selected.';
        container.appendChild(empty);
        hint.textContent = 'Open a sequence, then add it.';
        updateSelectionSummary();
        return;
    }

    selectedSequenceFilters.forEach((filter, index) => {
        const card = document.createElement('div');
        card.className = 'sequence-card';

        const header = document.createElement('div');
        header.className = 'sequence-header';

        const titleWrap = document.createElement('div');
        titleWrap.className = 'sequence-title-wrap';
        const title = document.createElement('div');
        title.className = 'sequence-title';
        title.textContent = `Seq ${index + 1}: ${filter.sequenceName}`;
        titleWrap.appendChild(title);

        const subtitle = document.createElement('div');
        subtitle.className = 'sequence-subtitle';
        subtitle.textContent = 'Selected sequence';
        titleWrap.appendChild(subtitle);
        header.appendChild(titleWrap);

        const headerActions = document.createElement('div');
        headerActions.className = 'sequence-header-actions';

        const removeButton = document.createElement('button');
        removeButton.type = 'button';
        removeButton.className = 'button-danger-soft button-small sequence-remove-button';
        removeButton.textContent = '×';
        removeButton.title = `Remove ${filter.sequenceName}`;
        removeButton.setAttribute('aria-label', `Remove ${filter.sequenceName}`);
        removeButton.onclick = () => removeSequenceFilter(filter.sequenceID || filter.sequenceName);
        headerActions.appendChild(removeButton);
        header.appendChild(headerActions);

        card.appendChild(header);
        card.appendChild(renderSequencePresetSection(filter));

        const groups = document.createElement('div');
        groups.className = 'sequence-groups';
        groups.appendChild(renderSequenceGroup(filter, 'video'));
        groups.appendChild(renderSequenceGroup(filter, 'audio'));
        card.appendChild(groups);

        container.appendChild(card);
    });

    hint.textContent = 'Refresh reloads and resets every selected sequence.';
    updateSelectionSummary();
}

async function readCurrentActiveSequenceFilter(locked) {
    const raw = await callHost('getActiveSequenceTrackUsage()');
    const data = safeJsonParse(raw);

    if (!data || data.error || !data.sequenceName) {
        alert(data && data.error ? data.error : 'No active sequence is available in Premiere.');
        return null;
    }

    return createSequenceFilter(data.sequenceID || '', data.sequenceName, data.videoTrackUsage || [], data.audioTrackUsage || [], locked);
}

function mergeSequenceTrackUsage(existing, incoming) {
    existing.videoTrackUsage = incoming.videoTrackUsage;
    existing.audioTrackUsage = incoming.audioTrackUsage;
    existing.sequenceID = incoming.sequenceID || existing.sequenceID;
    existing.sequenceName = incoming.sequenceName || existing.sequenceName;
}

async function addCurrentActiveSequence() {
    if (isCopying) {
        return;
    }

    if (!(await ensureHostScriptLoaded())) {
        return;
    }

    const incoming = await readCurrentActiveSequenceFilter(false);
    if (!incoming) {
        return;
    }

    const existingIndex = selectedSequenceFilters.findIndex((filter) => (filter.sequenceID && filter.sequenceID === incoming.sequenceID) || filter.sequenceName === incoming.sequenceName);

    if (existingIndex >= 0) {
        const existing = selectedSequenceFilters[existingIndex];
        mergeSequenceTrackUsage(existing, incoming);
    } else {
        selectedSequenceFilters.push(incoming);
    }

    selectedSequenceFilters = selectedSequenceFilters.map((filter) => sanitizeSequenceFilter(filter, false));
    saveSequenceFilters();
    renderSequenceFilters();
}

async function refreshAllSelectedSequenceTracks(resetChoices) {
    if (!selectedSequenceFilters.length) {
        return {
            refreshedCount: 0,
            missingSequences: [],
            resetCount: 0
        };
    }

    const filtersPayload = selectedSequenceFilters.map((filter) => ({
        sequenceID: filter.sequenceID || '',
        sequenceName: filter.sequenceName || ''
    }));
    const raw = await callHost(`getSequenceTrackUsagePlan("${escapeForEvalScript(JSON.stringify(filtersPayload))}")`);
    const plan = safeJsonParse(raw);

    if (!plan || plan.error || !Array.isArray(plan.sequences)) {
        throw new Error(plan && plan.error
            ? `Could not refresh selected sequence tracks: ${plan.error}`
            : `Could not refresh selected sequence tracks. Raw response: ${raw}`);
    }

    const result = applySequenceTrackUsagePlan(selectedSequenceFilters, plan, !!resetChoices);
    selectedSequenceFilters = selectedSequenceFilters.map((filter) => sanitizeSequenceFilter(filter, false));
    trackRangeAnchor = null;
    saveSequenceFilters();
    renderSequenceFilters();
    return result;
}

function resetTrackSelection(sequenceKey) {
    const resetCount = clearTrackChoices(selectedSequenceFilters, sequenceKey);
    if (!resetCount) {
        return;
    }

    if (trackRangeAnchor && (!sequenceKey || trackRangeAnchor.sequenceKey === sequenceKey)) {
        trackRangeAnchor = null;
    }
    saveAndRenderTrackFilters();
}

function removeSequenceFilter(sequenceKey) {
    selectedSequenceFilters = selectedSequenceFilters.filter((filter) => (filter.sequenceID || filter.sequenceName) !== sequenceKey);
    selectedSequenceFilters = selectedSequenceFilters.map((filter) => sanitizeSequenceFilter(filter, false));
    if (trackRangeAnchor && trackRangeAnchor.sequenceKey === sequenceKey) {
        trackRangeAnchor = null;
    }
    saveSequenceFilters();
    renderSequenceFilters();
}

function buildIgnoredMediaPathsFromSelection() {
    const ignoredPaths = [];
    const seen = new Set();

    function addIgnoredPath(mediaPath) {
        if (!mediaPath) {
            return;
        }

        const key = normalizeMediaKey(mediaPath);
        if (seen.has(key)) {
            return;
        }

        seen.add(key);
        ignoredPaths.push(mediaPath);
    }

    selectedSequenceFilters.forEach((filter) => {
        (filter.videoTrackUsage || []).forEach((entry) => {
            if (filter.ignoredVideoTracks.indexOf(entry.trackNumber) !== -1) {
                (entry.mediaPaths || []).forEach((mediaPath) => {
                    addIgnoredPath(mediaPath);
                });
            }
        });

        (filter.audioTrackUsage || []).forEach((entry) => {
            if (filter.ignoredAudioTracks.indexOf(entry.trackNumber) !== -1) {
                (entry.mediaPaths || []).forEach((mediaPath) => {
                    addIgnoredPath(mediaPath);
                });
            }
        });
    });

    return ignoredPaths;
}

function buildIgnoredMediaSet() {
    return new Set(buildIgnoredMediaPathsFromSelection().map((mediaPath) => normalizeMediaKey(mediaPath)));
}

function getProjectFolderEntries() {
    if (!latestPlan || !Array.isArray(latestPlan.folders)) {
        return [];
    }

    return latestPlan.folders
        .map(normalizeProjectFolderPath)
        .filter(isRootProjectFolderPath)
        .map((folderPath) => {
            const count = (latestPlan.tasks || []).filter((task) => isTaskInsideProjectFolder(task, folderPath)).length;
            return {
                folderPath,
                count
            };
        });
}

function isProjectFolderIgnored(folderPath) {
    const normalized = normalizeProjectFolderPath(folderPath);
    return ignoredProjectFolders.some((value) => normalizeProjectFolderPath(value) === normalized);
}

function isProjectFolderIncluded(folderPath) {
    const normalized = normalizeProjectFolderPath(folderPath);
    return includedProjectFolders.some((value) => normalizeProjectFolderPath(value) === normalized);
}

function setProjectFolderIgnored(folderPath, ignored) {
    const normalized = normalizeProjectFolderPath(folderPath);
    ignoredProjectFolders = ignoredProjectFolders.filter((value) => normalizeProjectFolderPath(value) !== normalized);
    includedProjectFolders = includedProjectFolders.filter((value) => normalizeProjectFolderPath(value) !== normalized);

    if (ignored && normalized) {
        ignoredProjectFolders.push(normalized);
    }
}

function setProjectFolderIncluded(folderPath, included) {
    const normalized = normalizeProjectFolderPath(folderPath);
    includedProjectFolders = includedProjectFolders.filter((value) => normalizeProjectFolderPath(value) !== normalized);
    ignoredProjectFolders = ignoredProjectFolders.filter((value) => normalizeProjectFolderPath(value) !== normalized);

    if (included && normalized) {
        includedProjectFolders.push(normalized);
    }
}

function cycleProjectFolderState(folderPath) {
    const included = isProjectFolderIncluded(folderPath);
    const ignored = isProjectFolderIgnored(folderPath);

    folderRangeAnchor = null;

    if (!included && !ignored) {
        setProjectFolderIncluded(folderPath, true);
    } else if (included) {
        setProjectFolderIgnored(folderPath, true);
    } else {
        setProjectFolderIgnored(folderPath, false);
    }

    renderProjectFolderFilters();
    updateSelectionSummary();
}

function ignoreProjectFolderRange(anchor, folderPath) {
    const entries = getProjectFolderEntries();
    const target = normalizeProjectFolderPath(folderPath);

    if (!anchor) {
        folderRangeAnchor = target;
        renderProjectFolderFilters();
        return;
    }

    const startIndex = entries.findIndex((entry) => entry.folderPath === anchor);
    const endIndex = entries.findIndex((entry) => entry.folderPath === target);

    if (startIndex === -1 || endIndex === -1) {
        folderRangeAnchor = target;
        renderProjectFolderFilters();
        return;
    }

    const start = Math.min(startIndex, endIndex);
    const end = Math.max(startIndex, endIndex);

    for (let index = start; index <= end; index += 1) {
        setProjectFolderIgnored(entries[index].folderPath, true);
    }

    folderRangeAnchor = null;
    renderProjectFolderFilters();
    updateSelectionSummary();
}

function handleProjectFolderClick(folderPath) {
    if (folderClickTimer) {
        clearTimeout(folderClickTimer);
    }

    folderClickTimer = setTimeout(() => {
        cycleProjectFolderState(folderPath);
        folderClickTimer = null;
    }, 220);
}

function handleProjectFolderDoubleClick(folderPath) {
    if (folderClickTimer) {
        clearTimeout(folderClickTimer);
        folderClickTimer = null;
    }

    ignoreProjectFolderRange(folderRangeAnchor, folderPath);
}

function resetProjectFolderSelection() {
    includedProjectFolders = [];
    ignoredProjectFolders = [];
    folderRangeAnchor = null;
    renderProjectFolderFilters();
    updateSelectionSummary();
}

function renderProjectFolderFilters() {
    const container = document.getElementById('projectFolderFilters');
    const hint = document.getElementById('projectFolderHint');

    if (!container) {
        return;
    }

    container.innerHTML = '';
    const entries = getProjectFolderEntries();

    if (!entries.length) {
        const empty = document.createElement('div');
        empty.className = 'small-note';
        empty.textContent = 'No Premiere project folders found.';
        container.appendChild(empty);
        if (hint) {
            hint.textContent = 'Only root folders are shown.';
        }
        return;
    }

    entries.forEach((entry) => {
        const item = document.createElement('div');
        const button = document.createElement('button');
        const included = isProjectFolderIncluded(entry.folderPath);
        const ignored = isProjectFolderIgnored(entry.folderPath);
        const isAnchor = folderRangeAnchor === entry.folderPath;
        const displayName = entry.folderPath.split('/').pop();

        item.className = 'choice-item folder-choice-item';
        button.type = 'button';
        button.className = `choice-button folder-button${included ? ' is-included' : ''}${ignored ? ' is-ignored' : ''}${isAnchor ? ' is-range-anchor' : ''}`;
        button.title = `${displayName}: ${entry.count} ${entry.count === 1 ? 'file' : 'files'}`;
        button.onclick = (event) => {
            if (event.detail === 1) {
                handleProjectFolderClick(entry.folderPath);
            }
        };
        button.ondblclick = () => handleProjectFolderDoubleClick(entry.folderPath);

        const icon = document.createElement('span');
        icon.className = 'folder-icon';
        icon.setAttribute('aria-hidden', 'true');
        button.appendChild(icon);

        const title = document.createElement('span');
        title.className = 'choice-title';
        title.textContent = displayName;
        button.appendChild(title);

        const count = document.createElement('span');
        count.className = 'folder-count';
        count.textContent = String(entry.count);
        button.appendChild(count);

        item.appendChild(button);

        container.appendChild(item);
    });

    if (hint) {
        const includedCount = includedProjectFolders.length;
        const skippedCount = ignoredProjectFolders.length;
        if (includedCount || skippedCount) {
            hint.textContent = `${includedCount} included · ${skippedCount} ignored`;
        } else {
            hint.textContent = 'Gray folders have no special rule.';
        }
    }
}

function buildTaskSelectionMap() {
    const selections = {};

    if (!latestPlan || !sourceTree) {
        return selections;
    }

    visitTree(sourceTree, (node) => {
        if (node.type === 'file' && Array.isArray(node.taskIndexes)) {
            node.taskIndexes.forEach((taskIndex) => {
                const task = latestPlan.tasks[taskIndex];
                if (task) {
                    selections[`${task.source} -> ${task.destination}`] = node.selected;
                }
            });
        }
    });

    return selections;
}

function applyTaskSelectionMap(selections) {
    if (!latestPlan || !sourceTree) {
        return;
    }

    visitTree(sourceTree, (node) => {
        if (node.type === 'file' && Array.isArray(node.taskIndexes)) {
            node.taskIndexes.forEach((taskIndex) => {
                const task = latestPlan.tasks[taskIndex];
                if (task) {
                    const key = `${task.source} -> ${task.destination}`;
                    if (Object.prototype.hasOwnProperty.call(selections, key)) {
                        node.selected = selections[key];
                    }
                }
            });
        }
    });

    sourceTree.forEach(syncNodeFromChildren);
}

function renderSourceTree() {
    const container = document.getElementById('sourceTree');
    container.innerHTML = '';

    if (!sourceTree || !sourceTree.length) {
        const row = document.createElement('li');
        row.className = 'list-empty';
        row.textContent = 'No copyable media found in the current Premiere project.';
        container.appendChild(row);
        updateSelectionSummary();
        return;
    }

    const renderNodes = (nodes, parentElement) => {
        nodes.forEach((node) => {
            const item = document.createElement('li');
            item.className = 'tree-node';

            const row = document.createElement('div');
            row.className = 'tree-row';

            const toggle = document.createElement('button');
            toggle.type = 'button';
            toggle.className = `tree-toggle${node.children.length ? '' : ' is-hidden'}`;
            toggle.textContent = node.children.length ? (node.expanded ? '-' : '+') : '';
            toggle.onclick = () => {
                node.expanded = !node.expanded;
                renderSourceTree();
            };
            row.appendChild(toggle);

            const checkbox = document.createElement('input');
            checkbox.type = 'checkbox';
            checkbox.checked = node.selected;
            checkbox.onchange = () => {
                selectionTouched = true;
                applySelectionToNode(node, checkbox.checked, true);
                sourceTree.forEach(syncNodeFromChildren);
                renderSourceTree();
            };
            row.appendChild(checkbox);

            const label = document.createElement('div');
            label.className = 'tree-label';

            const name = document.createElement('span');
            name.className = 'tree-name';
            name.textContent = node.name;
            label.appendChild(name);

            const meta = document.createElement('span');
            meta.className = 'tree-meta';
            meta.textContent = node.fullPath;
            label.appendChild(meta);

            row.appendChild(label);
            item.appendChild(row);

            if (node.children.length && node.expanded) {
                const children = document.createElement('ul');
                children.className = 'tree-children';
                renderNodes(node.children, children);
                item.appendChild(children);
            }

            parentElement.appendChild(item);
        });
    };

    renderNodes(sourceTree, container);
    updateSelectionSummary();
}

async function loadProjectPlan() {
    const previousSelections = selectionTouched ? buildTaskSelectionMap() : null;
    const previousIncludedFolders = includedProjectFolders.slice();
    const previousIgnoredFolders = ignoredProjectFolders.slice();

    if (!(await ensureHostScriptLoaded())) {
        return false;
    }

    const planRaw = await callHost('getProjectCopyPlan("")');
    const plan = safeJsonParse(planRaw);

    if (!plan || plan.error) {
        setText('selectionSummary', plan && plan.error ? plan.error : `Could not read the project structure from Premiere. Raw response: ${planRaw}`);
        sourceTree = [];
        renderSourceTree();
        return false;
    }

    latestPlan = plan;
    sourceTree = buildSourceTree(plan.tasks || []);
    const availableFolders = new Set((plan.folders || []).map(normalizeProjectFolderPath).filter(isRootProjectFolderPath));
    includedProjectFolders = previousIncludedFolders
        .map(normalizeProjectFolderPath)
        .filter((folderPath) => availableFolders.has(folderPath));
    ignoredProjectFolders = previousIgnoredFolders
        .map(normalizeProjectFolderPath)
        .filter((folderPath) => availableFolders.has(folderPath) && includedProjectFolders.indexOf(folderPath) === -1);
    if (previousSelections) {
        applyTaskSelectionMap(previousSelections);
    }
    loadSequenceFilters();
    renderSourceTree();
    renderProjectFolderFilters();
    return true;
}

async function refreshSourceList() {
    setText('selectionSummary', 'Refreshing project files from Premiere...');
    await loadProjectPlan();
}

async function refreshProject() {
    if (isCopying) {
        return;
    }

    const refreshButton = document.getElementById('refreshProjectButton');
    if (refreshButton) {
        refreshButton.disabled = true;
    }

    try {
        setText('summaryText', 'Refreshing project from Premiere...');
        setText('selectionSummary', 'Refreshing project files from Premiere...');
        const refreshed = await loadProjectPlan();
        if (refreshed) {
            const trackResult = await refreshAllSelectedSequenceTracks(true);
            const missingCount = trackResult.missingSequences.length;
            setText(
                'summaryText',
                missingCount
                    ? `Project refreshed. Reset and updated ${trackResult.refreshedCount} selected sequence${trackResult.refreshedCount === 1 ? '' : 's'}; ${missingCount} sequence${missingCount === 1 ? ' was' : 's were'} not found.`
                    : `Project refreshed. Reset and updated all ${trackResult.refreshedCount} selected sequence${trackResult.refreshedCount === 1 ? '' : 's'}, Premiere folders, and source files.`
            );
        }
    } catch (error) {
        setText('summaryText', error && error.message ? error.message : String(error));
    } finally {
        if (refreshButton) {
            refreshButton.disabled = false;
        }
    }
}

async function copyFileWithRobocopy(source, destinationPath) {
    return new Promise((resolve) => {
        let stagingDir = '';
        let stagedPath = '';
        let settled = false;
        const finish = (result) => {
            if (settled) {
                return;
            }
            settled = true;

            try {
                if (stagedPath && fs.existsSync(stagedPath)) {
                    fs.unlinkSync(stagedPath);
                }
            } catch (cleanupFileError) {}
            try {
                if (stagingDir && fs.existsSync(stagingDir)) {
                    fs.rmdirSync(stagingDir);
                }
            } catch (cleanupDirectoryError) {}

            resolve(result);
        };

        try {
            const sourceDir = path.dirname(source);
            const fileName = path.basename(source);
            const destinationDir = path.dirname(destinationPath);
            ensureDirectorySync(destinationDir);
            stagingDir = path.join(
                destinationDir,
                `.projectcollector-copy-${process.pid}-${Date.now()}-${Math.random().toString(16).slice(2)}`
            );
            ensureDirectorySync(stagingDir);
            stagedPath = path.join(stagingDir, fileName);

            const args = [
                sourceDir,
                stagingDir,
                fileName,
                '/R:1',
                '/W:1',
                '/NJH',
                '/NJS',
                '/NFL',
                '/NDL',
                '/NC',
                '/NS',
                '/NP'
            ];

            const child = spawn('robocopy', args, { windowsHide: true });
            let stderr = '';

            child.stderr.on('data', (chunk) => {
                stderr += chunk.toString();
            });

            child.on('error', (error) => {
                finish({ success: false, message: error.message });
            });

            child.on('close', (code) => {
                if (code !== null && code < 8) {
                    try {
                        if (!fs.existsSync(stagedPath)) {
                            throw new Error(`robocopy reported success but did not create ${stagedPath}`);
                        }

                        const sourceSize = fs.statSync(source).size;
                        const stagedSize = fs.statSync(stagedPath).size;
                        if (sourceSize !== stagedSize) {
                            throw new Error(`Copied file size mismatch: source ${sourceSize} bytes, copied ${stagedSize} bytes`);
                        }

                        if (fs.existsSync(destinationPath)) {
                            fs.unlinkSync(destinationPath);
                        }
                        fs.renameSync(stagedPath, destinationPath);

                        if (!fs.existsSync(destinationPath) || fs.statSync(destinationPath).size !== sourceSize) {
                            throw new Error(`Final copied file could not be verified at ${destinationPath}`);
                        }

                        finish({ success: true, message: '', destinationPath });
                    } catch (finalizeError) {
                        finish({ success: false, message: finalizeError.message });
                    }
                    return;
                }

                finish({
                    success: false,
                    message: stderr || `robocopy failed with exit code ${code}`
                });
            });
        } catch (error) {
            finish({ success: false, message: error.message });
        }
    });
}

async function ensureHostScriptLoaded() {
    if (hostScriptReady) {
        return true;
    }

    const extensionPath = csInterface.getSystemPath(SystemPath.EXTENSION).replace(/\\/g, '/');
    const scriptPath = `${extensionPath}/jsx/collector.jsx`;
    const escapedPath = escapeForEvalScript(scriptPath);
    const result = await callHost(`$.evalFile("${escapedPath}")`);

    if (result === 'EvalScript error.' || result === 'false') {
        setText('summaryText', `Could not load Premiere host script: ${result}`);
        return false;
    }

    hostScriptReady = true;
    return true;
}

function setBusyState(busy) {
    isCopying = busy;
    document.getElementById('chooseButton').disabled = busy;
    const compareButton = document.getElementById('compareButton');
    if (compareButton) {
        compareButton.disabled = busy;
    }
    document.getElementById('collectButton').disabled = busy;
    const refreshButton = document.getElementById('refreshProjectButton');
    if (refreshButton) {
        refreshButton.disabled = busy;
    }
    document.getElementById('updateButton').disabled = busy || !(remoteVersion && compareVersions(remoteVersion, localVersion) > 0);
}

function hideCompletionPrompt() {
    const prompt = document.getElementById('completionPrompt');
    if (prompt) {
        prompt.classList.remove('is-visible');
    }
}

function showCompletionPrompt(success, message) {
    const prompt = document.getElementById('completionPrompt');
    const title = document.getElementById('completionTitle');
    const body = document.getElementById('completionMessage');

    if (!prompt || !title || !body) {
        return;
    }

    title.textContent = success ? 'Done without error' : 'Done with errors';
    title.classList.toggle('has-error', !success);
    body.textContent = message;
    prompt.classList.add('is-visible');
}

function hideTrackConflictPrompt() {
    const prompt = document.getElementById('trackConflictPrompt');
    if (prompt) {
        prompt.classList.remove('is-visible');
    }
}

function isTrackConflictDecision(decision) {
    return decision === 'copy' || decision === 'skip';
}

function countTrackConflictDecisions(decisions) {
    return (decisions || []).reduce((count, decision) => (
        isTrackConflictDecision(decision) ? count + 1 : count
    ), 0);
}

function updateTrackConflictPromptState() {
    if (!trackConflictPromptState) {
        return;
    }

    const conflicts = trackConflictPromptState.conflicts;
    const decisions = trackConflictPromptState.decisions;
    const rows = trackConflictPromptState.rows;
    const resolvedCount = countTrackConflictDecisions(decisions);
    const remainingCount = conflicts.length - resolvedCount;
    const continueButton = document.getElementById('trackConflictContinueButton');
    const status = document.getElementById('trackConflictStatus');

    rows.forEach((rowInfo, index) => {
        const decision = decisions[index] || '';
        rowInfo.copyButton.classList.toggle('is-selected', decision === 'copy');
        rowInfo.skipButton.classList.toggle('is-selected', decision === 'skip');
    });

    if (continueButton) {
        continueButton.disabled = remainingCount > 0;
    }
    if (status) {
        status.textContent = remainingCount > 0
            ? `${resolvedCount} of ${conflicts.length} resolved. Choose an action for ${remainingCount} more.`
            : `All ${conflicts.length} conflict${conflicts.length === 1 ? '' : 's'} resolved. The backup is ready to continue.`;
    }
}

function setTrackConflictChoice(conflictIndex, decision) {
    const index = parseInt(conflictIndex, 10);
    if (
        !trackConflictPromptState
        || !isTrackConflictDecision(decision)
        || !Number.isInteger(index)
        || index < 0
        || index >= trackConflictPromptState.conflicts.length
    ) {
        return;
    }

    trackConflictPromptState.decisions[index] = decision;
    updateTrackConflictPromptState();
}

function setAllTrackConflictChoices(decision) {
    if (!trackConflictPromptState || !isTrackConflictDecision(decision)) {
        return;
    }

    trackConflictPromptState.conflicts.forEach((conflict, index) => {
        setTrackConflictChoice(index, decision);
    });
}

function handleTrackConflictListClick(event) {
    if (!trackConflictPromptState) {
        return;
    }

    const list = document.getElementById('trackConflictList');
    let target = event && event.target;

    while (target && target !== list) {
        const decision = target.getAttribute && target.getAttribute('data-track-conflict-choice');
        const conflictIndex = target.getAttribute && target.getAttribute('data-track-conflict-index');
        if (isTrackConflictDecision(decision) && conflictIndex !== null) {
            setTrackConflictChoice(conflictIndex, decision);
            return;
        }
        target = target.parentNode;
    }
}

function finishTrackConflictPrompt(shouldContinue) {
    if (!trackConflictPromptResolver || !trackConflictPromptState) {
        return;
    }

    const conflicts = trackConflictPromptState.conflicts;
    const decisions = trackConflictPromptState.decisions;
    if (shouldContinue && countTrackConflictDecisions(decisions) !== conflicts.length) {
        return;
    }

    const resolve = trackConflictPromptResolver;
    const result = shouldContinue
        ? conflicts.map((conflict, index) => ({
            mediaKey: conflict.mediaKey,
            decision: decisions[index]
        }))
        : null;
    const list = document.getElementById('trackConflictList');
    if (list) {
        list.onclick = null;
    }
    trackConflictPromptResolver = null;
    trackConflictPromptState = null;
    hideTrackConflictPrompt();
    resolve(result);
}

function showTrackConflictPrompt(conflicts) {
    const prompt = document.getElementById('trackConflictPrompt');
    const list = document.getElementById('trackConflictList');

    if (!prompt || !list || !Array.isArray(conflicts) || !conflicts.length) {
        return Promise.resolve([]);
    }

    list.innerHTML = '';
    trackConflictPromptState = {
        conflicts,
        decisions: new Array(conflicts.length),
        rows: []
    };

    conflicts.forEach((conflict, index) => {
        const row = document.createElement('div');
        row.className = 'track-conflict-row';

        const details = document.createElement('div');
        details.className = 'track-conflict-details';

        const name = document.createElement('div');
        name.className = 'track-conflict-name';
        name.textContent = conflict.name;

        const source = document.createElement('div');
        source.className = 'track-conflict-path';
        source.textContent = conflict.source;

        details.appendChild(name);
        details.appendChild(source);

        const actions = document.createElement('div');
        actions.className = 'track-conflict-row-actions';

        const copyButton = document.createElement('button');
        copyButton.type = 'button';
        copyButton.className = 'track-conflict-choice track-conflict-copy';
        copyButton.setAttribute('data-track-conflict-choice', 'copy');
        copyButton.setAttribute('data-track-conflict-index', String(index));
        copyButton.textContent = 'Copy';

        const skipButton = document.createElement('button');
        skipButton.type = 'button';
        skipButton.className = 'track-conflict-choice track-conflict-skip';
        skipButton.setAttribute('data-track-conflict-choice', 'skip');
        skipButton.setAttribute('data-track-conflict-index', String(index));
        skipButton.textContent = 'Do not copy';

        actions.appendChild(copyButton);
        actions.appendChild(skipButton);
        row.appendChild(details);
        row.appendChild(actions);
        list.appendChild(row);
        trackConflictPromptState.rows.push({
            row,
            copyButton,
            skipButton
        });
    });

    list.onclick = handleTrackConflictListClick;
    prompt.classList.add('is-visible');
    updateTrackConflictPromptState();

    return new Promise((resolve) => {
        trackConflictPromptResolver = resolve;
    });
}

function resetResults() {
    hideCompletionPrompt();
    setText('currentFile', 'Waiting to start');
    setText('progressText', '0 / 0 files copied');
    setText('summaryText', 'Select a destination folder to build the project package.');
    document.getElementById('progressFill').style.width = '0%';
    document.getElementById('errorList').innerHTML = '';
    document.getElementById('missingList').innerHTML = '';
}

function toggleSourceList() {
    listVisible = !listVisible;
    document.getElementById('sourceListBox').style.display = listVisible ? 'block' : 'none';
    setText('showListButton', listVisible ? 'Hide List' : 'Show List');

    if (listVisible) {
        refreshSourceList();
    }
}

function setAllSelections(selected) {
    if (!sourceTree) {
        return;
    }

    selectionTouched = true;
    sourceTree.forEach((node) => applySelectionToNode(node, selected, true));
    sourceTree.forEach(syncNodeFromChildren);
    renderSourceTree();
}

function renderList(listId, items, formatter, emptyMessage) {
    const list = document.getElementById(listId);
    list.innerHTML = '';

    if (!items.length) {
        const row = document.createElement('li');
        row.className = 'list-empty';
        row.textContent = emptyMessage || 'None';
        list.appendChild(row);
        return;
    }

    items.forEach((item) => {
        const row = document.createElement('li');
        row.className = 'list-item';
        row.textContent = formatter(item);
        list.appendChild(row);
    });
}

async function chooseFolder() {
    if (isCopying) {
        return;
    }

    const result = window.cep.fs.showOpenDialogEx(false, true, 'Select Destination Folder');

    if (result.data.length > 0) {
        destination = result.data[0];
        setText('path', destination);
        try {
            localStorage.setItem(DESTINATION_STORAGE_KEY, destination);
        } catch (error) {}
        setText('summaryText', 'Destination ready. Click BACKUP PROJECT to begin.');
    }
}

async function chooseCompareFolder() {
    if (isCopying) {
        return;
    }

    const result = window.cep.fs.showOpenDialogEx(false, true, 'Select Skip Folder');

    if (result.data.length > 0) {
        compareLocation = result.data[0];
        compareFiles = [];
        compareScanErrors = [];
        setText('comparePath', compareLocation);
    }
}

function hasIgnoredTracks(filtersPayload) {
    return filtersPayload.some((filter) => (
        (Array.isArray(filter.ignoredVideoTracks) && filter.ignoredVideoTracks.length)
        || (Array.isArray(filter.ignoredAudioTracks) && filter.ignoredAudioTracks.length)
    ));
}

function copyMediaKeySet(mediaSet) {
    const result = new Set();
    if (mediaSet && typeof mediaSet.forEach === 'function') {
        mediaSet.forEach((mediaKey) => {
            if (mediaKey) {
                result.add(mediaKey);
            }
        });
    }
    return result;
}

function buildTrackRuleContext(tasks, includedMediaSet, ignoredMediaSet) {
    const conflicts = [];
    const seenMedia = new Set();
    const taskMediaSet = new Set();
    const includedSet = copyMediaKeySet(includedMediaSet);
    const ignoredSet = copyMediaKeySet(ignoredMediaSet);
    const conflictMediaSet = new Set();
    const ignoredOnlyMediaSet = new Set();

    (tasks || []).forEach((task) => {
        const mediaKey = normalizeMediaKey(task.source);
        if (!mediaKey) {
            return;
        }
        taskMediaSet.add(mediaKey);

        if (seenMedia.has(mediaKey) || !includedSet.has(mediaKey) || !ignoredSet.has(mediaKey)) {
            return;
        }

        seenMedia.add(mediaKey);
        conflictMediaSet.add(mediaKey);
        conflicts.push({
            name: task.name || path.basename(task.source || '') || 'Unnamed media',
            source: task.source || '',
            mediaKey
        });
    });

    ignoredSet.forEach((mediaKey) => {
        if (taskMediaSet.has(mediaKey) && !conflictMediaSet.has(mediaKey)) {
            ignoredOnlyMediaSet.add(mediaKey);
        }
    });

    return {
        includedMediaSet: includedSet,
        ignoredMediaSet: ignoredSet,
        conflictMediaSet,
        ignoredOnlyMediaSet,
        conflicts
    };
}

function buildTrackConflictDecisionInfo(trackRuleContext, resolvedConflicts) {
    const info = {
        valid: false,
        error: '',
        copyMediaSet: new Set(),
        skipMediaSet: new Set()
    };
    const expectedMediaSet = trackRuleContext && trackRuleContext.conflictMediaSet
        ? trackRuleContext.conflictMediaSet
        : new Set();
    const decidedMediaSet = new Set();

    if (!Array.isArray(resolvedConflicts)) {
        info.error = 'Track conflict choices were not returned in a valid format.';
        return info;
    }

    (resolvedConflicts || []).forEach((conflict) => {
        if (info.error) {
            return;
        }

        if (
            !conflict
            || !conflict.mediaKey
            || !expectedMediaSet.has(conflict.mediaKey)
            || !isTrackConflictDecision(conflict.decision)
        ) {
            info.error = 'A track conflict choice referred to media outside the conflict list.';
            return;
        }
        if (decidedMediaSet.has(conflict.mediaKey)) {
            info.error = 'A track conflict file received more than one decision.';
            return;
        }

        decidedMediaSet.add(conflict.mediaKey);
        if (conflict.decision === 'copy') {
            info.copyMediaSet.add(conflict.mediaKey);
        } else {
            info.skipMediaSet.add(conflict.mediaKey);
        }
    });

    if (!info.error && decidedMediaSet.size !== expectedMediaSet.size) {
        info.error = 'Every listed track conflict must receive exactly one choice.';
    }

    info.valid = !info.error;
    return info;
}

function getTrackConflictDecision(task, trackRuleContext, conflictDecisionInfo) {
    if (!trackRuleContext || !conflictDecisionInfo || !conflictDecisionInfo.valid) {
        return '';
    }

    const mediaKey = normalizeMediaKey(task.source);
    if (!trackRuleContext.conflictMediaSet.has(mediaKey)) {
        return '';
    }
    if (conflictDecisionInfo.copyMediaSet.has(mediaKey)) {
        return 'copy';
    }
    if (conflictDecisionInfo.skipMediaSet.has(mediaKey)) {
        return 'skip';
    }

    return '';
}

async function buildIgnoredMediaSetForCopy(filtersPayload) {
    const ignoredPaths = buildIgnoredMediaPathsFromSelection();
    const addIgnoredPaths = (mediaPaths) => {
        (mediaPaths || []).forEach((mediaPath) => {
            if (mediaPath) {
                ignoredPaths.push(mediaPath);
            }
        });
    };
    const buildIgnoredInfo = (warning) => {
        const ignoredMediaSet = new Set();

        ignoredPaths.forEach((mediaPath) => {
            ignoredMediaSet.add(normalizeMediaKey(mediaPath));
        });

        return {
            ignoredMediaSet,
            warning: warning || ''
        };
    };

    if (!hasIgnoredTracks(filtersPayload)) {
        return buildIgnoredInfo('');
    }

    const ignoredRaw = await callHost(`getIgnoredTrackMediaPlan("${escapeForEvalScript(JSON.stringify(filtersPayload))}")`);
    const ignoredPlan = safeJsonParse(ignoredRaw);

    if (!ignoredPlan || ignoredPlan.error) {
        return buildIgnoredInfo(
            ignoredPlan && ignoredPlan.error
                ? `Ignored track nested-media check failed: ${ignoredPlan.error}`
                : `Ignored track nested-media check failed. Raw response: ${ignoredRaw}`
        );
    }

    addIgnoredPaths(ignoredPlan.mediaPaths || []);

    const missingWarning = Array.isArray(ignoredPlan.missingSequences) && ignoredPlan.missingSequences.length
        ? `Ignored track check could not match sequence${ignoredPlan.missingSequences.length === 1 ? '' : 's'}: ${ignoredPlan.missingSequences.join(', ')}`
        : '';

    return buildIgnoredInfo(missingWarning);
}

function buildLinkProjectTasks(allTasks, copiedTasks, compareMatches) {
    const copiedDestinationBySource = new Map();
    const skipDestinationBySource = new Map();
    const relinkTasksBySource = new Map();

    (copiedTasks || []).forEach((task) => {
        if (!task || !task.source || !task.destinationPath) {
            return;
        }

        copiedDestinationBySource.set(normalizeMediaKey(task.source), {
            source: task.source,
            destination: task.destinationPath
        });
    });

    (compareMatches || []).forEach((entry) => {
        const task = entry && entry.task;
        const match = entry && entry.match;
        if (!task || !task.source || !match || !match.path) {
            return;
        }

        skipDestinationBySource.set(normalizeMediaKey(task.source), {
            source: task.source,
            destination: match.path
        });
    });

    (allTasks || []).forEach((task) => {
        if (!task || !task.source) {
            return;
        }

        const sourceKey = normalizeMediaKey(task.source);
        const copiedTask = copiedDestinationBySource.get(sourceKey);
        const skipTask = skipDestinationBySource.get(sourceKey);
        const collectedDestination = copiedTask ? copiedTask.destination : '';
        const skipDestination = skipTask ? skipTask.destination : '';
        const relinkTask = {
            source: task.source,
            destination: collectedDestination || skipDestination || task.source,
            targetKind: collectedDestination ? 'collected' : (skipDestination ? 'skip-location' : 'original')
        };
        const existingTask = relinkTasksBySource.get(sourceKey);

        if (!existingTask || (
            existingTask.targetKind !== 'collected'
            && (relinkTask.targetKind === 'collected' || relinkTask.targetKind === 'skip-location')
        )) {
            relinkTasksBySource.set(sourceKey, relinkTask);
        }
    });

    copiedDestinationBySource.forEach((copiedTask, sourceKey) => {
        if (!relinkTasksBySource.has(sourceKey)) {
            relinkTasksBySource.set(sourceKey, {
                source: copiedTask.source,
                destination: copiedTask.destination,
                targetKind: 'collected'
            });
        }
    });

    return Array.from(relinkTasksBySource.values());
}

function createCopyRuleContext(options) {
    const source = options || {};
    return {
        treeSelectedTaskSet: source.treeSelectedTaskSet || new Set(),
        sequenceScopedMediaSet: source.sequenceScopedMediaSet || null,
        trackRuleContext: source.trackRuleContext || buildTrackRuleContext([], new Set(), new Set()),
        conflictDecisionInfo: source.conflictDecisionInfo || null,
        includedProjectFolders: Array.isArray(source.includedProjectFolders)
            ? source.includedProjectFolders.slice()
            : includedProjectFolders.slice(),
        ignoredProjectFolders: Array.isArray(source.ignoredProjectFolders)
            ? source.ignoredProjectFolders.slice()
            : ignoredProjectFolders.slice()
    };
}

function getCopySkipReason(task, ruleContext) {
    const rules = ruleContext || createCopyRuleContext();
    const mediaKey = normalizeMediaKey(task.source);

    if (isTaskInsideProjectFolderList(task, rules.ignoredProjectFolders)) {
        return `skipped because Premiere folder "${task.binPath}" is ignored`;
    }

    if (rules.sequenceScopedMediaSet && !rules.sequenceScopedMediaSet.has(mediaKey)) {
        return 'skipped because it is not used by the chosen sequences';
    }

    const trackConflictDecision = getTrackConflictDecision(
        task,
        rules.trackRuleContext,
        rules.conflictDecisionInfo
    );
    if (trackConflictDecision === 'skip') {
        return 'skipped by your track conflict decision';
    }
    if (trackConflictDecision === 'copy') {
        return '';
    }

    if (rules.trackRuleContext.conflictMediaSet.has(mediaKey)) {
        return 'skipped because its track conflict was not resolved';
    }

    if (rules.trackRuleContext.ignoredOnlyMediaSet.has(mediaKey)) {
        return 'skipped by ignored track selection';
    }

    if (isTaskInsideProjectFolderList(task, rules.includedProjectFolders)) {
        return '';
    }

    if (!rules.treeSelectedTaskSet.has(task)) {
        return 'skipped by Source File List selection';
    }

    return '';
}

async function buildCopyReadyContext() {
    setText('summaryText', 'Reading Premiere project structure...');

    const hostLoaded = await ensureHostScriptLoaded();
    if (!hostLoaded) {
        return { ok: false };
    }

    if (!latestPlan || !sourceTree) {
        const planLoaded = await loadProjectPlan();
        if (!planLoaded || !latestPlan) {
            return { ok: false };
        }
    }

    const treeSelectedTasks = getSelectedTasks();
    const treeSelectedTaskSet = new Set(treeSelectedTasks);
    const filtersPayload = buildSequenceFiltersPayload().filter((filter) => filter.sequenceID || filter.sequenceName);
    const ignoredMediaInfo = await buildIgnoredMediaSetForCopy(filtersPayload);
    const ignoredMediaSet = ignoredMediaInfo.ignoredMediaSet;
    const copyWarnings = ignoredMediaInfo.warning ? [ignoredMediaInfo.warning] : [];
    let sequenceScopedMediaSet = null;
    let sequenceScopeInfo = null;
    let scopedPlan = null;
    const ignoredTracksSelected = hasIgnoredTracks(filtersPayload);

    if (sequenceOnlyMode && !filtersPayload.length) {
        alert('Add at least one sequence before using selected-sequence collection mode.');
        return { ok: false };
    }

    if (sequenceOnlyMode || ignoredTracksSelected) {
        if (!filtersPayload.length) {
            setText('summaryText', 'Could not inspect track selections because no sequence is selected.');
            return { ok: false };
        }

        const scopedRaw = await callHost(`getSequenceScopedMediaPlan("${escapeForEvalScript(JSON.stringify(filtersPayload))}")`);
        scopedPlan = safeJsonParse(scopedRaw);
        if (!scopedPlan || scopedPlan.error) {
            setText('summaryText', scopedPlan && scopedPlan.error ? scopedPlan.error : `Could not build the selected-sequence media plan. Raw response: ${scopedRaw}`);
            return { ok: false };
        }

        if (sequenceOnlyMode) {
            sequenceScopedMediaSet = new Set((scopedPlan.mediaPaths || []).map((mediaPath) => normalizeMediaKey(mediaPath)));
            sequenceScopeInfo = scopedPlan;
        }
    }

    const includedMediaSet = new Set(((scopedPlan && scopedPlan.mediaPaths) || []).map((mediaPath) => normalizeMediaKey(mediaPath)));
    const trackRuleContext = buildTrackRuleContext(
        latestPlan.tasks || [],
        includedMediaSet,
        ignoredMediaSet
    );
    const copyRuleContext = createCopyRuleContext({
        treeSelectedTaskSet,
        sequenceScopedMediaSet,
        trackRuleContext,
        includedProjectFolders,
        ignoredProjectFolders
    });

    return {
        ok: true,
        copyRuleContext,
        copyWarnings,
        sequenceScopeInfo,
        trackConflicts: trackRuleContext.conflicts
    };
}

async function collect() {
    try {
        await runCollection();
    } catch (error) {
        const message = error && error.message ? error.message : String(error || 'Unknown backup error');
        trackConflictPromptResolver = null;
        trackConflictPromptState = null;
        hideTrackConflictPrompt();
        setText('currentFile', 'Backup stopped by an unexpected error');
        setText('summaryText', `Backup could not continue. ${message}`);
        renderList(
            'errorList',
            [{ source: 'Project Collector', destination: destination || '', message }],
            (item) => `${item.source} -> ${item.destination} | ${item.message}`,
            'No errors.'
        );
        showCompletionPrompt(false, `Backup could not continue. ${message}`);
        setBusyState(false);
        console.error('Project Collector backup failed', error);
    }
}

async function runCollection() {
    if (isCopying) {
        return;
    }

    if (!destination) {
        alert('Select destination first');
        return;
    }

    resetResults();
    setBusyState(true);
    setText('summaryText', 'Loading Premiere host script...');

    const context = await buildCopyReadyContext();
    if (!context.ok) {
        setBusyState(false);
        return;
    }

    let copyRuleContext = context.copyRuleContext;
    const copyWarnings = context.copyWarnings;
    const sequenceScopeInfo = context.sequenceScopeInfo;
    const trackConflicts = context.trackConflicts;

    if (trackConflicts.length) {
        setText('currentFile', 'Waiting for track conflict decisions');
        setText('summaryText', `${trackConflicts.length} media conflict${trackConflicts.length === 1 ? '' : 's'} found between included and ignored tracks.`);
        const resolvedConflicts = await showTrackConflictPrompt(trackConflicts);
        if (!resolvedConflicts) {
            setBusyState(false);
            setText('currentFile', 'Backup cancelled');
            setText('summaryText', 'Backup cancelled before any files were copied.');
            return;
        }
        const conflictDecisionInfo = buildTrackConflictDecisionInfo(
            copyRuleContext.trackRuleContext,
            resolvedConflicts
        );
        if (!conflictDecisionInfo.valid) {
            throw new Error(conflictDecisionInfo.error || 'Track conflict choices could not be validated.');
        }
        copyRuleContext = createCopyRuleContext(Object.assign({}, copyRuleContext, {
            conflictDecisionInfo
        }));
    }

    const selectedTasksBeforeCompare = (latestPlan.tasks || []).filter(
        (task) => !getCopySkipReason(task, copyRuleContext)
    );

    let compareInfo = {
        files: [],
        errors: [],
        blocked: false,
        lookup: buildCompareLookup([])
    };
    let compareMatches = [];

    if (compareLocation) {
        setText('currentFile', 'Inspecting compare location');
        compareInfo = await inspectCompareLocation();

        if (compareInfo.blocked) {
            setBusyState(false);
            setText('summaryText', compareInfo.errors[0] || 'Compare location could not be inspected.');
            return;
        }

        if (compareInfo.errors.length) {
            copyWarnings.push(`Compare location scan had ${compareInfo.errors.length} warning${compareInfo.errors.length === 1 ? '' : 's'}. Some folders or files could not be read.`);
        }
    }

    const selectedTasks = [];
    for (let compareIndex = 0; compareIndex < selectedTasksBeforeCompare.length; compareIndex += 1) {
        const task = selectedTasksBeforeCompare[compareIndex];
        if (compareLocation) {
            setText('currentFile', `Verifying skip location ${compareIndex + 1} / ${selectedTasksBeforeCompare.length}: ${task.name}`);
        }
        const match = compareLocation ? await findCompareMatchForTask(task, compareInfo.lookup) : null;

        if (match) {
            compareMatches.push({
                task,
                match
            });
        } else {
            selectedTasks.push(task);
        }
    }

    if (compareLocation) {
        copyWarnings.push(`${compareInfo.files.length} compare file${compareInfo.files.length === 1 ? '' : 's'} checked. ${compareMatches.length} already exist and were skipped.`);
    }

    const rootPath = path.join(destination, latestPlan.projectName);
    const plan = {
        rootPath,
        tasks: buildCollectedTasks(selectedTasks, rootPath),
        missingMedia: Array.isArray(latestPlan.missingMedia) ? latestPlan.missingMedia : []
    };
    const total = plan.tasks.length;

    const willCreateReducedProject = sequenceOnlyMode && createReducedProject && sequenceScopeInfo && Array.isArray(sequenceScopeInfo.includedSequenceIDs) && sequenceScopeInfo.includedSequenceIDs.length;
    if (total > 0 || copyProjectFile || willCreateReducedProject) {
        try {
            ensureDirectorySync(plan.rootPath);
        } catch (error) {
            setBusyState(false);
            setText('summaryText', `Could not prepare destination folder. ${error.message}`);
            return;
        }
    }

    const failures = [];
    const copiedMediaTasks = [];
    const missingMedia = Array.isArray(plan.missingMedia) ? plan.missingMedia : [];
    const skippedItems = [];

    (latestPlan.tasks || []).forEach((task) => {
        const skipReason = getCopySkipReason(task, copyRuleContext);

        if (skipReason) {
            skippedItems.push(`${task.source} -> ${skipReason}`);
        }
    });

    compareMatches.forEach((item) => {
        skippedItems.push(`${item.task.source} -> skipped because it already exists in compare location: ${item.match.path}`);
    });

    if (sequenceOnlyMode && sequenceScopeInfo) {
        setText('summaryText', `Windows robocopy mode active. Copying only media used by ${sequenceScopeInfo.includedSequences.length} chosen/nested sequences into ${plan.rootPath}`);
    } else {
        setText('summaryText', `Windows robocopy mode active. Copying into ${plan.rootPath}`);
    }
    setText('progressText', `0 / ${total} files copied`);
    document.getElementById('progressFill').style.width = total === 0 ? '100%' : '0%';

    for (let index = 0; index < total; index += 1) {
        const task = plan.tasks[index];
        const destinationPath = path.join(plan.rootPath, task.destination);
        setText('currentFile', `${task.name} -> ${task.destination}`);

        const copyResult = await copyFileWithRobocopy(task.source, destinationPath);

        if (!copyResult.success) {
            failures.push({
                source: task.source,
                destination: destinationPath,
                message: copyResult.message || 'Unknown copy error'
            });
        } else {
            copiedMediaTasks.push(Object.assign({}, task, {
                destinationPath
            }));
        }

        const completed = index + 1;
        const percent = total === 0 ? 100 : Math.round((completed / total) * 100);
        document.getElementById('progressFill').style.width = `${percent}%`;
        setText('progressText', `${completed} / ${total} files processed`);

        if (completed % 20 === 0) {
            await new Promise((resolve) => setTimeout(resolve, 0));
        }
    }

    let copiedProjectMessage = '';
    let linkedProjectMessage = '';
    let copiedProjectPath = '';
    if (copyProjectFile) {
        setText('currentFile', 'Saving and copying Premiere project file');
        const projectSaveRaw = await callHost('saveCurrentProjectAndGetPath()');
        const projectSaveInfo = safeJsonParse(projectSaveRaw);

        if (projectSaveInfo && !projectSaveInfo.error && projectSaveInfo.projectPath) {
            const projectCopyResult = await copyProjectFileIntoCollectedRoot(plan.rootPath, projectSaveInfo.projectPath, getCollectedProjectFileName(projectSaveInfo.projectPath));
            if (projectCopyResult.success) {
                copiedProjectPath = projectCopyResult.destinationPath;
                copiedProjectMessage = ` Project file copied as ${path.basename(copiedProjectPath)}.`;
            } else {
                copiedProjectMessage = ` Project file copy failed: ${projectCopyResult.message || 'Unknown error'}.`;
                failures.push({
                    source: projectSaveInfo.projectPath,
                    destination: plan.rootPath,
                    message: projectCopyResult.message || 'Project file copy failed'
                });
            }
        } else {
            copiedProjectMessage = ` Project file copy skipped.${projectSaveInfo && projectSaveInfo.error ? ` ${projectSaveInfo.error}` : ''}`;
            failures.push({
                source: latestPlan && latestPlan.projectPath ? latestPlan.projectPath : 'Premiere project file',
                destination: plan.rootPath,
                message: projectSaveInfo && projectSaveInfo.error ? projectSaveInfo.error : 'Project file path was not available'
            });
        }

        if (linkProjectAfterCollection) {
            if (copiedProjectPath) {
                setText('currentFile', 'Linking copied project to collected media');
                const projectRelinkTasks = buildLinkProjectTasks(latestPlan.tasks || [], copiedMediaTasks, compareMatches);
                const linkRaw = await callHost(
                    `linkProjectCopyToCollectedMedia("${escapeForEvalScript(copiedProjectPath)}","${escapeForEvalScript(JSON.stringify(projectRelinkTasks))}")`
                );
                const linkResult = safeJsonParse(linkRaw);

                if (linkResult && linkResult.success === true && !linkResult.error) {
                    linkedProjectMessage = ` BACKUP project offlined ${linkResult.offlineCount || 0} files, linked ${linkResult.linkedCollectedCount || 0} to collected media, ${linkResult.linkedSkipLocationCount || 0} to verified skip-location files, restored ${linkResult.linkedOriginalCount || 0} to original paths, and remains open.`;
                    if (Array.isArray(linkResult.failed) && linkResult.failed.length) {
                        linkResult.failed.forEach((message) => {
                            failures.push({
                                source: 'BACKUP project relink',
                                destination: copiedProjectPath,
                                message
                            });
                        });
                        linkedProjectMessage = ` BACKUP project relink reported ${linkResult.failed.length} error${linkResult.failed.length === 1 ? '' : 's'}.`;
                    }
                } else {
                    const linkFailureMessage = linkResult && linkResult.error
                        ? linkResult.error
                        : (linkResult && Array.isArray(linkResult.failed) && linkResult.failed.length
                            ? linkResult.failed.join('; ')
                            : `Could not link copied project. Raw response: ${linkRaw}`);
                    linkedProjectMessage = ` BACKUP project relink failed and was not saved. ${linkFailureMessage}`;
                    failures.push({
                        source: 'BACKUP project relink',
                        destination: copiedProjectPath,
                        message: linkFailureMessage
                    });
                }
            } else {
                linkedProjectMessage = ' Project link skipped because the project file was not copied.';
            }
        }
    } else {
        copiedProjectMessage = ' Project file copy disabled.';
    }

    setText('currentFile', total ? 'Copy finished' : 'No copyable media found');
    let reducedProjectMessage = '';

    if (willCreateReducedProject) {
        const reducedRaw = await callHost(
            `createReducedProjectFromSequenceSelection("${escapeForEvalScript(plan.rootPath)}","${escapeForEvalScript(JSON.stringify(sequenceScopeInfo.includedSequenceIDs))}")`
        );
        const reducedProjectResult = safeJsonParse(reducedRaw);
        if (reducedProjectResult && !reducedProjectResult.error && reducedProjectResult.reducedProjectPath) {
            reducedProjectMessage = ` Reduced project created: ${reducedProjectResult.reducedProjectPath}`;
        } else {
            reducedProjectMessage = ` Reduced project could not be created.${reducedProjectResult && reducedProjectResult.error ? ` ${reducedProjectResult.error}` : ''}`;
            failures.push({
                source: 'Reduced Premiere project',
                destination: plan.rootPath,
                message: reducedProjectResult && reducedProjectResult.error ? reducedProjectResult.error : 'Reduced project could not be created'
            });
        }
    }

    const skippedAndWarnings = missingMedia.concat(skippedItems, copyWarnings);
    const finalSummary = `Completed. ${copiedMediaTasks.length} of ${total} media files copied, ${failures.length} error${failures.length === 1 ? '' : 's'}, ${skippedAndWarnings.length} skipped.${copiedProjectMessage}${linkedProjectMessage}${reducedProjectMessage}`;

    setText(
        'summaryText',
        finalSummary
    );

    renderList('errorList', failures, (item) => `${item.source} -> ${item.destination} | ${item.message}`, 'No error. All files copied successfully.');
    renderList('missingList', skippedAndWarnings, (item) => item, 'No skipped items.');
    showCompletionPrompt(
        failures.length === 0,
        failures.length === 0
            ? `All selected files copied successfully.${copiedProjectMessage}${linkedProjectMessage}${reducedProjectMessage}`
            : `Copy finished with ${failures.length} error${failures.length === 1 ? '' : 's'}. Check Files Not Copied for details.`
    );
    setBusyState(false);
}

document.addEventListener('DOMContentLoaded', async () => {
    readVersionInfo();
    resetResults();
    try {
        const savedDestination = localStorage.getItem(DESTINATION_STORAGE_KEY) || '';
        if (savedDestination) {
            destination = savedDestination;
            setText('path', destination);
            setText('summaryText', 'Saved destination loaded. Click BACKUP PROJECT to begin.');
        }
    } catch (error) {}

    document.getElementById('sourceListBox').style.display = 'none';
    setText('showListButton', 'Show List');
    setSourceSectionVisibility(false);
    try {
        const savedSequenceOnlyMode = localStorage.getItem(SEQUENCE_ONLY_MODE_STORAGE_KEY);
        sequenceOnlyMode = savedSequenceOnlyMode === null ? true : savedSequenceOnlyMode === '1';
        createReducedProject = localStorage.getItem(CREATE_REDUCED_PROJECT_STORAGE_KEY) === '1';
    } catch (error) {
        sequenceOnlyMode = true;
        createReducedProject = false;
    }
    try {
        const savedIgnoreVisible = localStorage.getItem(IGNORE_SECTION_VISIBLE_STORAGE_KEY);
        const ignoreVisible = savedIgnoreVisible === null ? true : savedIgnoreVisible === '1';
        setIgnoreSectionVisibility(ignoreVisible);
    } catch (error) {
        setIgnoreSectionVisibility(true);
    }
    setProjectFolderSectionVisibility(false);
    try {
        const savedCopyProjectFile = localStorage.getItem(COPY_PROJECT_FILE_STORAGE_KEY);
        copyProjectFile = savedCopyProjectFile === null ? true : savedCopyProjectFile === '1';
        linkProjectAfterCollection = localStorage.getItem(LINK_PROJECT_AFTER_COLLECTION_STORAGE_KEY) === '1';
    } catch (error) {
        copyProjectFile = true;
        linkProjectAfterCollection = false;
    }
    setUpdateButton(`Version ${localVersion}`, false);
    checkForUpdates();
    syncSequenceModeUI();
    syncProjectOptionUI();
    await loadProjectPlan();
});

window.addEventListener('focus', () => {
    refreshSourceList();
});
