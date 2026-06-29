const csInterface = new CSInterface();
const fs = require('fs');
const os = require('os');
const path = require('path');
const https = require('https');
const childProcess = require('child_process');
const { spawn } = childProcess;

let destination = null;
let isCopying = false;
let hostScriptReady = false;
let latestPlan = null;
let sourceTree = null;
let listVisible = false;
let selectionTouched = false;
let localVersion = 'unknown';
let remoteVersion = null;
let selectedSequenceFilters = [];
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

const SEQUENCE_FILTERS_STORAGE_KEY = 'projectcollector.sequenceFilters';
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
            const escapedScriptPath = tempUpdaterScriptPath.replace(/'/g, "''");
            const escapedZipPath = tempUpdaterZipPath.replace(/'/g, "''");
            const destination = getUserCepExtensionPath().replace(/'/g, "''");
            const escapedResultPath = tempUpdaterResultPath.replace(/'/g, "''");
            const escapedLogPath = tempUpdaterLogPath.replace(/'/g, "''");
            const command = `Start-Process PowerShell -Verb RunAs -ArgumentList '-NoExit','-NoProfile','-ExecutionPolicy','Bypass','-File','${escapedScriptPath}','-ZipPath','${escapedZipPath}','-Destination','${destination}','-ResultPath','${escapedResultPath}','-LogPath','${escapedLogPath}'`;

            childProcess.execFile(
                'powershell.exe',
                ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', command],
                (error) => {
                    if (error) {
                        setText('summaryText', `Could not launch updater. ${error.message}`);
                        return;
                    }

                    setText('summaryText', `Updater launched for ${getUserCepExtensionPath()}. Accept the Windows prompt if it appears.`);
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
    const sequenceNames = sequenceOnlyMode
        ? selectedSequenceFilters
            .filter((filter) => filter && filter.sequenceName)
            .map((filter) => filter.sequenceName)
        : [];
    const baseName = sequenceNames.length === 1
        ? sequenceNames[0]
        : (latestPlan && latestPlan.projectName ? latestPlan.projectName : path.basename(projectPath || 'Premiere_Project', extension));

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

function isTaskInsideIgnoredProjectFolder(task) {
    return ignoredProjectFolders.some((folderPath) => isTaskInsideProjectFolder(task, folderPath));
}

function isTaskInsideIncludedProjectFolder(task) {
    return includedProjectFolders.some((folderPath) => isTaskInsideProjectFolder(task, folderPath));
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

    filter.ignoredVideoTracks = Array.from(new Set(filter.ignoredVideoTracks)).sort((a, b) => a - b);
    filter.ignoredAudioTracks = Array.from(new Set(filter.ignoredAudioTracks)).sort((a, b) => a - b);
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
        true
    );
}

function loadSequenceFilters() {
    let savedFilters = [];

    try {
        const parsed = JSON.parse(localStorage.getItem(SEQUENCE_FILTERS_STORAGE_KEY) || '[]');
        savedFilters = Array.isArray(parsed) ? parsed : [];
    } catch (error) {
        savedFilters = [];
    }

    selectedSequenceFilters = savedFilters
        .filter((filter) => filter && filter.sequenceName)
        .map((filter, index) => sanitizeSequenceFilter(filter, index === 0));

    const defaultFilter = getDefaultSequenceFilter();
    if (!defaultFilter) {
        renderSequenceFilters();
        return;
    }

    if (!selectedSequenceFilters.length) {
        selectedSequenceFilters = [defaultFilter];
        saveSequenceFilters();
        renderSequenceFilters();
        return;
    }

    const defaultIndex = selectedSequenceFilters.findIndex((filter) => filter.sequenceID === defaultFilter.sequenceID || filter.sequenceName === defaultFilter.sequenceName);
    if (defaultIndex >= 0) {
        const existing = selectedSequenceFilters[defaultIndex];
        existing.videoTrackUsage = defaultFilter.videoTrackUsage;
        existing.audioTrackUsage = defaultFilter.audioTrackUsage;
        existing.locked = true;
        if (defaultIndex !== 0) {
            selectedSequenceFilters.splice(defaultIndex, 1);
            selectedSequenceFilters.unshift(existing);
        }
    }

    selectedSequenceFilters = selectedSequenceFilters.map((filter, index) => sanitizeSequenceFilter(filter, index === 0));
    saveSequenceFilters();
    renderSequenceFilters();
}

function saveSequenceFilters() {
    try {
        localStorage.setItem(SEQUENCE_FILTERS_STORAGE_KEY, JSON.stringify(selectedSequenceFilters));
    } catch (error) {}
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
    list.className = 'choice-grid';

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

        item.className = 'choice-item';
        button.type = 'button';
        button.className = `choice-button${ignored ? ' is-ignored' : ''}${isAnchor ? ' is-range-anchor' : ''}`;
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

        const meta = document.createElement('span');
        meta.className = 'choice-meta';
        meta.textContent = `${entry.clipCount} ${entry.clipCount === 1 ? 'clip' : 'clips'}`;
        item.appendChild(button);
        item.appendChild(meta);

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

function renderSequenceFilters() {
    const container = document.getElementById('sequenceFilters');
    const hint = document.getElementById('sequenceFilterHint');
    container.innerHTML = '';

    if (!selectedSequenceFilters.length) {
        const empty = document.createElement('div');
        empty.className = 'small-note';
        empty.textContent = 'No sequences selected yet. Open a sequence in Premiere and add it here.';
        container.appendChild(empty);
        hint.textContent = 'Switch to a sequence in Premiere, then click Add Current Active Sequence.';
        updateSelectionSummary();
        return;
    }

    selectedSequenceFilters.forEach((filter, index) => {
        const card = document.createElement('div');
        card.className = 'sequence-card';

        const header = document.createElement('div');
        header.className = 'sequence-header';

        const titleWrap = document.createElement('div');
        const title = document.createElement('div');
        title.className = 'sequence-title';
        title.textContent = `Seq ${index + 1}: ${filter.sequenceName}`;
        titleWrap.appendChild(title);

        const subtitle = document.createElement('div');
        subtitle.className = 'sequence-subtitle';
        subtitle.textContent = filter.locked
            ? 'Current active sequence loaded by default. Use Refresh Tracks after changing it in Premiere.'
            : 'Added from another active sequence. Remove it anytime if you no longer want its track filters.';
        titleWrap.appendChild(subtitle);
        header.appendChild(titleWrap);

        if (!filter.locked) {
            const removeButton = document.createElement('button');
            removeButton.type = 'button';
            removeButton.className = 'button-danger-soft button-small';
            removeButton.textContent = 'x';
            removeButton.onclick = () => removeSequenceFilter(filter.sequenceID || filter.sequenceName);
            header.appendChild(removeButton);
        }

        card.appendChild(header);

        const groups = document.createElement('div');
        groups.className = 'sequence-groups';
        groups.appendChild(renderSequenceGroup(filter, 'video'));
        groups.appendChild(renderSequenceGroup(filter, 'audio'));
        card.appendChild(groups);

        container.appendChild(card);
    });

    hint.textContent = 'Nested sequence media is collected correctly even when a visible track count only reflects direct clips. Refresh updates track counts and keeps your green/faded choices.';
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
        if (existingIndex === 0) {
            existing.locked = true;
        }
    } else {
        selectedSequenceFilters.push(incoming);
    }

    selectedSequenceFilters = selectedSequenceFilters.map((filter, index) => sanitizeSequenceFilter(filter, index === 0));
    saveSequenceFilters();
    renderSequenceFilters();
}

async function refreshTrackSelection() {
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
    if (existingIndex === -1) {
        alert('This active sequence is not in the track list yet. Click Add Current Active Sequence first.');
        return;
    }

    mergeSequenceTrackUsage(selectedSequenceFilters[existingIndex], incoming);
    selectedSequenceFilters = selectedSequenceFilters.map((filter, index) => sanitizeSequenceFilter(filter, index === 0));
    saveSequenceFilters();
    renderSequenceFilters();
}

function resetTrackSelection() {
    selectedSequenceFilters = [];
    trackRangeAnchor = null;
    try {
        localStorage.removeItem(SEQUENCE_FILTERS_STORAGE_KEY);
    } catch (error) {}
    loadSequenceFilters();
}

function removeSequenceFilter(sequenceKey) {
    selectedSequenceFilters = selectedSequenceFilters.filter((filter, index) => index === 0 || (filter.sequenceID || filter.sequenceName) !== sequenceKey);
    selectedSequenceFilters = selectedSequenceFilters.map((filter, index) => sanitizeSequenceFilter(filter, index === 0));
    saveSequenceFilters();
    renderSequenceFilters();
}

function buildIgnoredMediaSet() {
    const ignoredPaths = new Set();

    selectedSequenceFilters.forEach((filter) => {
        (filter.videoTrackUsage || []).forEach((entry) => {
            if (filter.ignoredVideoTracks.indexOf(entry.trackNumber) !== -1) {
                (entry.mediaPaths || []).forEach((mediaPath) => {
                    ignoredPaths.add(normalizeMediaKey(mediaPath));
                });
            }
        });

        (filter.audioTrackUsage || []).forEach((entry) => {
            if (filter.ignoredAudioTracks.indexOf(entry.trackNumber) !== -1) {
                (entry.mediaPaths || []).forEach((mediaPath) => {
                    ignoredPaths.add(normalizeMediaKey(mediaPath));
                });
            }
        });
    });

    return ignoredPaths;
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
            hint.textContent = 'Only root bins are shown here. Root-level media that is not inside a Premiere folder will be copied into one CollectedMedias folder.';
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

        item.className = 'choice-item';
        button.type = 'button';
        button.className = `choice-button folder-button${included ? ' is-included' : ''}${ignored ? ' is-ignored' : ''}${isAnchor ? ' is-range-anchor' : ''}`;
        button.onclick = (event) => {
            if (event.detail === 1) {
                handleProjectFolderClick(entry.folderPath);
            }
        };
        button.ondblclick = () => handleProjectFolderDoubleClick(entry.folderPath);

        const title = document.createElement('span');
        title.className = 'choice-title';
        title.textContent = displayName;
        button.appendChild(title);

        const meta = document.createElement('span');
        meta.className = 'choice-meta';
        meta.textContent = `${entry.count} ${entry.count === 1 ? 'file' : 'files'}`;
        item.appendChild(button);
        item.appendChild(meta);

        container.appendChild(item);
    });

    if (hint) {
        const includedCount = includedProjectFolders.length;
        const skippedCount = ignoredProjectFolders.length;
        if (includedCount || skippedCount) {
            hint.textContent = `${includedCount} force-copy folder${includedCount === 1 ? '' : 's'}, ${skippedCount} ignored folder${skippedCount === 1 ? '' : 's'}. Ignored bins win first; green bins copy even when track choices would skip them.`;
        } else {
            hint.textContent = 'Neutral root bins add no rule. Root-level media that is not inside a Premiere folder will be copied into one CollectedMedias folder.';
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

async function refreshEverything() {
    if (isCopying) {
        return;
    }

    setText('summaryText', 'Refreshing Premiere project data...');
    await refreshSourceList();
    setText('summaryText', 'Refresh complete.');
}

async function copyFileWithRobocopy(source, destinationPath) {
    return new Promise((resolve) => {
        try {
            const sourceDir = path.dirname(source);
            const fileName = path.basename(source);
            const destinationDir = path.dirname(destinationPath);
            ensureDirectorySync(destinationDir);

            const args = [
                sourceDir,
                destinationDir,
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
                resolve({ success: false, message: error.message });
            });

            child.on('close', (code) => {
                if (code !== null && code < 8) {
                    resolve({ success: true, message: '' });
                    return;
                }

                resolve({
                    success: false,
                    message: stderr || `robocopy failed with exit code ${code}`
                });
            });
        } catch (error) {
            resolve({ success: false, message: error.message });
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
    document.getElementById('collectButton').disabled = busy;
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

function resetResults() {
    hideCompletionPrompt();
    setText('currentFile', 'Waiting to start');
    setText('progressText', '0 / 0 files copied');
    setText('summaryText', 'Select a destination folder to build the project package.');
    document.getElementById('progressFill').style.width = '0%';
    document.getElementById('errorList').innerHTML = '';
    document.getElementById('missingList').innerHTML = '';
}

async function resetEverything() {
    if (isCopying) {
        return;
    }

    selectedSequenceFilters = [];
    includedProjectFolders = [];
    ignoredProjectFolders = [];
    trackRangeAnchor = null;
    folderRangeAnchor = null;
    selectionTouched = false;
    listVisible = false;
    sequenceOnlyMode = true;
    createReducedProject = false;
    copyProjectFile = true;
    linkProjectAfterCollection = false;

    try {
        localStorage.removeItem(SEQUENCE_FILTERS_STORAGE_KEY);
        localStorage.setItem(SEQUENCE_ONLY_MODE_STORAGE_KEY, '1');
        localStorage.setItem(CREATE_REDUCED_PROJECT_STORAGE_KEY, '0');
        localStorage.setItem(COPY_PROJECT_FILE_STORAGE_KEY, '1');
        localStorage.setItem(LINK_PROJECT_AFTER_COLLECTION_STORAGE_KEY, '0');
    } catch (error) {}

    setSourceSectionVisibility(false);
    setProjectFolderSectionVisibility(false);
    const sourceListBox = document.getElementById('sourceListBox');
    if (sourceListBox) {
        sourceListBox.style.display = 'none';
    }
    setText('showListButton', 'Show List');
    syncSequenceModeUI();
    syncProjectOptionUI();
    resetResults();
    setText('summaryText', 'Reset complete. Loading fresh Premiere project data...');
    await loadProjectPlan();
    setText('summaryText', 'Reset complete. Select a destination folder to build the project package.');
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
        setText('summaryText', 'Destination ready. Click Copy Project Media to begin.');
    }
}

function hasIgnoredTracks(filtersPayload) {
    return filtersPayload.some((filter) => (
        (Array.isArray(filter.ignoredVideoTracks) && filter.ignoredVideoTracks.length)
        || (Array.isArray(filter.ignoredAudioTracks) && filter.ignoredAudioTracks.length)
    ));
}

async function buildIgnoredMediaSetForCopy(filtersPayload) {
    const fallbackIgnoredMedia = buildIgnoredMediaSet();

    if (!hasIgnoredTracks(filtersPayload)) {
        return {
            ignoredMediaSet: fallbackIgnoredMedia,
            warning: ''
        };
    }

    const ignoredRaw = await callHost(`getIgnoredTrackMediaPlan("${escapeForEvalScript(JSON.stringify(filtersPayload))}")`);
    const ignoredPlan = safeJsonParse(ignoredRaw);

    if (!ignoredPlan || ignoredPlan.error) {
        return {
            ignoredMediaSet: fallbackIgnoredMedia,
            warning: ignoredPlan && ignoredPlan.error
                ? `Ignored track nested-media check failed: ${ignoredPlan.error}`
                : `Ignored track nested-media check failed. Raw response: ${ignoredRaw}`
        };
    }

    return {
        ignoredMediaSet: new Set((ignoredPlan.mediaPaths || []).map((mediaPath) => normalizeMediaKey(mediaPath))),
        warning: ''
    };
}

function buildLinkProjectTasks(copiedTasks) {
    return copiedTasks.map((task) => ({
        source: task.source,
        destination: task.destinationPath
    }));
}

async function collect() {
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

    const hostLoaded = await ensureHostScriptLoaded();
    if (!hostLoaded) {
        setBusyState(false);
        return;
    }

    setText('summaryText', 'Reading Premiere project structure...');

    if (!latestPlan || !sourceTree) {
        const planLoaded = await loadProjectPlan();
        if (!planLoaded || !latestPlan) {
            setBusyState(false);
            return;
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

    if (sequenceOnlyMode) {
        if (!filtersPayload.length) {
            alert('Add at least one sequence before using selected-sequence collection mode.');
            setBusyState(false);
            return;
        }

        const scopedRaw = await callHost(`getSequenceScopedMediaPlan("${escapeForEvalScript(JSON.stringify(filtersPayload))}")`);
        const scopedPlan = safeJsonParse(scopedRaw);
        if (!scopedPlan || scopedPlan.error) {
            setBusyState(false);
            setText('summaryText', scopedPlan && scopedPlan.error ? scopedPlan.error : `Could not build the selected-sequence media plan. Raw response: ${scopedRaw}`);
            return;
        }

        sequenceScopedMediaSet = new Set((scopedPlan.mediaPaths || []).map((mediaPath) => normalizeMediaKey(mediaPath)));
        sequenceScopeInfo = scopedPlan;
    }

    const selectedTasks = (latestPlan.tasks || []).filter((task) => {
        const mediaKey = normalizeMediaKey(task.source);
        const insideIncludedFolder = isTaskInsideIncludedProjectFolder(task);

        if (isTaskInsideIgnoredProjectFolder(task)) {
            return false;
        }

        if (insideIncludedFolder) {
            return true;
        }

        if (!treeSelectedTaskSet.has(task)) {
            return false;
        }

        if (sequenceScopedMediaSet && !sequenceScopedMediaSet.has(mediaKey)) {
            return false;
        }

        if (ignoredMediaSet.has(mediaKey)) {
            return false;
        }

        return true;
    });

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
        const mediaKey = normalizeMediaKey(task.source);
        const insideIncludedFolder = isTaskInsideIncludedProjectFolder(task);

        if (isTaskInsideIgnoredProjectFolder(task)) {
            skippedItems.push(`${task.source} -> skipped because Premiere folder "${task.binPath}" is ignored`);
            return;
        }

        if (insideIncludedFolder) {
            return;
        }

        if (!treeSelectedTaskSet.has(task)) {
            skippedItems.push(`${task.source} -> skipped by Source File List selection`);
            return;
        }

        if (sequenceScopedMediaSet && !sequenceScopedMediaSet.has(mediaKey)) {
            skippedItems.push(`${task.source} -> skipped because it is not used by the chosen sequences`);
            return;
        }

        if (ignoredMediaSet.has(mediaKey)) {
            skippedItems.push(`${task.source} -> skipped by ignored track selection`);
        }
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
                const linkRaw = await callHost(
                    `linkProjectCopyToCollectedMedia("${escapeForEvalScript(copiedProjectPath)}","${escapeForEvalScript(JSON.stringify(buildLinkProjectTasks(copiedMediaTasks)))}")`
                );
                const linkResult = safeJsonParse(linkRaw);

                if (linkResult && !linkResult.error) {
                    linkedProjectMessage = ` Project linked to ${linkResult.linkedCount || 0} collected media files.`;
                    if (Array.isArray(linkResult.failed) && linkResult.failed.length) {
                        linkResult.failed.forEach((message) => {
                            failures.push({
                                source: copiedProjectPath,
                                destination: copiedProjectPath,
                                message
                            });
                        });
                        linkedProjectMessage = ` Project linked to ${linkResult.linkedCount || 0} collected media files with ${linkResult.failed.length} relink errors.`;
                    }
                } else {
                    linkedProjectMessage = ` Project link failed.${linkResult && linkResult.error ? ` ${linkResult.error}` : ''}`;
                    failures.push({
                        source: copiedProjectPath,
                        destination: copiedProjectPath,
                        message: linkResult && linkResult.error ? linkResult.error : `Could not link copied project. Raw response: ${linkRaw}`
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
            setText('summaryText', 'Saved destination loaded. Click Copy Project Media to begin.');
        }
    } catch (error) {}

    document.getElementById('sourceListBox').style.display = 'none';
    setText('showListButton', 'Show List');
    setSourceSectionVisibility(false);
    try {
        sequenceOnlyMode = true;
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
