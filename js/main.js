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

const SEQUENCE_FILTERS_STORAGE_KEY = 'projectcollector.sequenceFilters';
const DESTINATION_STORAGE_KEY = 'projectcollector.destination';
const IGNORE_SECTION_VISIBLE_STORAGE_KEY = 'projectcollector.ignoreSectionVisible';
const SEQUENCE_ONLY_MODE_STORAGE_KEY = 'projectcollector.sequenceOnlyMode';
const CREATE_REDUCED_PROJECT_STORAGE_KEY = 'projectcollector.createReducedProject';

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

async function copyProjectFileIntoCollectedRoot(rootPath, projectPath) {
    if (!rootPath || !projectPath) {
        return { success: false, message: 'Project path was not available.' };
    }

    const destinationPath = path.join(rootPath, path.basename(projectPath));
    return copyFileWithRobocopy(projectPath, destinationPath);
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
        expanded: type !== 'file',
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

    const suffix = modeSummary.length ? ` (${modeSummary.join(' || ')})` : '';
    const treeScopeNote = (sequenceOnlyMode || ignoredTrackSummary.length)
        ? ' Source File List still shows the full project tree; final copy is reduced later by sequence mode and ignored tracks.'
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

function normalizeMediaKey(filePath) {
    return tryResolveWindowsPath(filePath).toLowerCase();
}

function sanitizeTrackUsageEntries(entries, prefix) {
    return Array.isArray(entries)
        ? entries.map((entry) => ({
            trackNumber: parseInt(entry.trackNumber, 10) || 0,
            label: entry.label || `${prefix}${entry.trackNumber}`,
            clipCount: parseInt(entry.clipCount, 10) || 0,
            mediaPaths: Array.isArray(entry.mediaPaths) ? entry.mediaPaths.slice() : []
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
    } else {
        selectedSequenceFilters.unshift(defaultFilter);
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

function getSequenceFilter(sequenceName) {
    return selectedSequenceFilters.find((filter) => filter.sequenceName === sequenceName) || null;
}

function getSequenceFilterByKey(sequenceKey) {
    return selectedSequenceFilters.find((filter) => (filter.sequenceID || filter.sequenceName) === sequenceKey) || null;
}

function buildTrackOptions(filter, kind) {
    const entries = kind === 'video' ? filter.videoTrackUsage : filter.audioTrackUsage;
    const ignored = kind === 'video' ? filter.ignoredVideoTracks : filter.ignoredAudioTracks;
    return entries.filter((entry) => ignored.indexOf(entry.trackNumber) === -1);
}

function renderTrackChipList(container, kind, filter) {
    const ignored = kind === 'video' ? filter.ignoredVideoTracks : filter.ignoredAudioTracks;

    if (!ignored.length) {
        const empty = document.createElement('div');
        empty.className = 'small-note';
        empty.textContent = kind === 'video' ? 'No ignored video tracks.' : 'No ignored audio tracks.';
        container.appendChild(empty);
        return;
    }

    ignored.forEach((trackNumber) => {
        const chip = document.createElement('div');
        chip.className = 'chip';
        chip.textContent = `${kind === 'video' ? 'Ignore Video Medias on Layer V' : 'Ignore Audio Medias on Layer A'}${trackNumber}`;

        const removeButton = document.createElement('button');
        removeButton.type = 'button';
        removeButton.textContent = 'x';
        removeButton.onclick = () => removeIgnoredTrack(filter.sequenceName, kind, trackNumber);
        chip.appendChild(removeButton);
        container.appendChild(chip);
    });
}

function renderSequenceGroup(filter, kind) {
    const group = document.createElement('div');
    group.className = 'sequence-group';

    const title = document.createElement('div');
    title.className = 'label';
    title.textContent = kind === 'video' ? 'Ignore Video Medias On Layers' : 'Ignore Audio Medias On Layers';
    group.appendChild(title);

    const row = document.createElement('div');
    row.className = 'filter-row';

    const select = document.createElement('select');
    const options = buildTrackOptions(filter, kind);

    if (!options.length) {
        const option = document.createElement('option');
        option.value = '';
        option.textContent = kind === 'video' ? 'No video tracks left' : 'No audio tracks left';
        select.appendChild(option);
        select.disabled = true;
    } else {
        options.forEach((entry) => {
            const option = document.createElement('option');
            option.value = String(entry.trackNumber);
            option.textContent = `${entry.label} (${entry.clipCount} clips)`;
            select.appendChild(option);
        });
    }

    row.appendChild(select);

    const note = document.createElement('div');
    note.className = 'small-note';
    note.textContent = kind === 'video'
        ? 'Choose a video layer to ignore for this sequence.'
        : 'Choose an audio layer to ignore for this sequence.';
    row.appendChild(note);

    const addButton = document.createElement('button');
    addButton.type = 'button';
    addButton.className = 'button-accent button-small';
    addButton.textContent = kind === 'video' ? 'Add V' : 'Add A';
    addButton.disabled = !options.length;
    addButton.onclick = () => {
        const trackNumber = parseInt(select.value, 10) || 0;
        if (trackNumber) {
            addIgnoredTrack(filter.sequenceID || filter.sequenceName, kind, trackNumber);
        }
    };
    row.appendChild(addButton);

    group.appendChild(row);

    const chips = document.createElement('div');
    chips.className = 'chip-list';
    renderTrackChipList(chips, kind, filter);
    group.appendChild(chips);

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
            ? 'Current active sequence by default. This one stays unless you refresh the active sequence.'
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

    hint.textContent = 'Switch to another sequence in Premiere, then click Add Current Active Sequence to include it too. Opening the same sequence again refreshes it instead of duplicating it.';
    updateSelectionSummary();
}

async function addCurrentActiveSequence() {
    if (isCopying) {
        return;
    }

    if (!(await ensureHostScriptLoaded())) {
        return;
    }

    const raw = await callHost('getActiveSequenceTrackUsage()');
    const data = safeJsonParse(raw);

    if (!data || data.error || !data.sequenceName) {
        alert(data && data.error ? data.error : 'No active sequence is available in Premiere.');
        return;
    }

    const incoming = createSequenceFilter(data.sequenceID || '', data.sequenceName, data.videoTrackUsage || [], data.audioTrackUsage || [], false);
    const existingIndex = selectedSequenceFilters.findIndex((filter) => (filter.sequenceID && filter.sequenceID === incoming.sequenceID) || filter.sequenceName === incoming.sequenceName);

    if (existingIndex >= 0) {
        const existing = selectedSequenceFilters[existingIndex];
        existing.videoTrackUsage = incoming.videoTrackUsage;
        existing.audioTrackUsage = incoming.audioTrackUsage;
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

function removeSequenceFilter(sequenceKey) {
    selectedSequenceFilters = selectedSequenceFilters.filter((filter, index) => index === 0 || (filter.sequenceID || filter.sequenceName) !== sequenceKey);
    selectedSequenceFilters = selectedSequenceFilters.map((filter, index) => sanitizeSequenceFilter(filter, index === 0));
    saveSequenceFilters();
    renderSequenceFilters();
}

function addIgnoredTrack(sequenceKey, kind, trackNumber) {
    const filter = getSequenceFilterByKey(sequenceKey);
    if (!filter || !trackNumber) {
        return;
    }

    const key = kind === 'video' ? 'ignoredVideoTracks' : 'ignoredAudioTracks';
    if (filter[key].indexOf(trackNumber) === -1) {
        filter[key].push(trackNumber);
        filter[key].sort((a, b) => a - b);
    }

    saveSequenceFilters();
    renderSequenceFilters();
}

function removeIgnoredTrack(sequenceKey, kind, trackNumber) {
    const filter = getSequenceFilterByKey(sequenceKey);
    if (!filter) {
        return;
    }

    const key = kind === 'video' ? 'ignoredVideoTracks' : 'ignoredAudioTracks';
    filter[key] = filter[key].filter((value) => value !== trackNumber);

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
    if (previousSelections) {
        applyTaskSelectionMap(previousSelections);
    }
    loadSequenceFilters();
    renderSourceTree();
    return true;
}

async function refreshSourceList() {
    setText('selectionSummary', 'Refreshing project files from Premiere...');
    await loadProjectPlan();
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

function resetResults() {
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

function renderList(listId, items, formatter) {
    const list = document.getElementById(listId);
    list.innerHTML = '';

    if (!items.length) {
        const row = document.createElement('li');
        row.className = 'list-empty';
        row.textContent = 'None';
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

    const escapedDestination = escapeForEvalScript(destination);
    const treeSelectedTasks = getSelectedTasks();
    const ignoredMediaSet = buildIgnoredMediaSet();
    let sequenceScopedMediaSet = null;
    let sequenceScopeInfo = null;

    if (sequenceOnlyMode) {
        const filtersPayload = buildSequenceFiltersPayload().filter((filter) => filter.sequenceID || filter.sequenceName);
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

        if (!sequenceScopedMediaSet.size) {
            setBusyState(false);
            alert('No media was found for the selected sequences. Check the selected sequences and ignored tracks, then try again.');
            return;
        }
    }

    const selectedTasks = treeSelectedTasks.filter((task) => {
        const mediaKey = normalizeMediaKey(task.source);

        if (sequenceScopedMediaSet && !sequenceScopedMediaSet.has(mediaKey)) {
            return false;
        }

        if (ignoredMediaSet.has(mediaKey)) {
            return false;
        }

        return true;
    });
    const plan = {
        rootPath: path.join(destination, latestPlan.projectName),
        tasks: selectedTasks,
        missingMedia: Array.isArray(latestPlan.missingMedia) ? latestPlan.missingMedia : []
    };

    const prepRaw = await callHost(`prepareProjectStructure("${escapedDestination}")`);
    const prep = safeJsonParse(prepRaw);
    if (!prep || prep.error) {
        setBusyState(false);
        setText('summaryText', prep && prep.error ? prep.error : `Could not prepare destination folders. Raw response: ${prepRaw}`);
        return;
    }

    const total = plan.tasks.length;
    const failures = [];
    const missingMedia = Array.isArray(plan.missingMedia) ? plan.missingMedia : [];
    const skippedBySelection = (latestPlan.tasks || [])
        .filter((task) => !treeSelectedTasks.includes(task))
        .map((task) => `${task.source} -> skipped by selection`);
    const skippedBySequenceScope = sequenceScopedMediaSet
        ? (latestPlan.tasks || [])
            .filter((task) => !sequenceScopedMediaSet.has(normalizeMediaKey(task.source)))
            .map((task) => `${task.source} -> skipped because it is not used by the chosen sequences`)
        : [];
    const skippedByIgnoredTracks = (latestPlan.tasks || [])
        .filter((task) => ignoredMediaSet.has(normalizeMediaKey(task.source)))
        .map((task) => `${task.source} -> skipped by ignored track filter`);

    if (sequenceOnlyMode && sequenceScopeInfo) {
        setText('summaryText', `Windows robocopy mode active. Copying only media used by ${sequenceScopeInfo.includedSequences.length} chosen/nested sequences into ${plan.rootPath}`);
    } else {
        setText('summaryText', `Windows robocopy mode active. Copying into ${plan.rootPath}`);
    }
    setText('progressText', `0 / ${total} files copied`);

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
    if (latestPlan && latestPlan.projectPath) {
        setText('currentFile', 'Saving and copying Premiere project file');
        const projectSaveRaw = await callHost('saveCurrentProjectAndGetPath()');
        const projectSaveInfo = safeJsonParse(projectSaveRaw);

        if (projectSaveInfo && !projectSaveInfo.error && projectSaveInfo.projectPath) {
            const projectCopyResult = await copyProjectFileIntoCollectedRoot(plan.rootPath, projectSaveInfo.projectPath);
            if (projectCopyResult.success) {
                copiedProjectMessage = ' Project file copied.';
            } else {
                copiedProjectMessage = ` Project file copy failed: ${projectCopyResult.message || 'Unknown error'}.`;
            }
        } else {
            copiedProjectMessage = ` Project file copy skipped.${projectSaveInfo && projectSaveInfo.error ? ` ${projectSaveInfo.error}` : ''}`;
        }
    }

    setText('currentFile', total ? 'Copy finished' : 'No copyable media found');
    let reducedProjectMessage = '';

    if (sequenceOnlyMode && createReducedProject && sequenceScopeInfo && Array.isArray(sequenceScopeInfo.includedSequenceIDs) && sequenceScopeInfo.includedSequenceIDs.length) {
        const reducedRaw = await callHost(
            `createReducedProjectFromSequenceSelection("${escapeForEvalScript(plan.rootPath)}","${escapeForEvalScript(JSON.stringify(sequenceScopeInfo.includedSequenceIDs))}")`
        );
        const reducedProjectResult = safeJsonParse(reducedRaw);
        if (reducedProjectResult && !reducedProjectResult.error && reducedProjectResult.reducedProjectPath) {
            reducedProjectMessage = ` Reduced project created: ${reducedProjectResult.reducedProjectPath}`;
        } else {
            reducedProjectMessage = ` Reduced project could not be created.${reducedProjectResult && reducedProjectResult.error ? ` ${reducedProjectResult.error}` : ''}`;
        }
    }

    setText(
        'summaryText',
        `Completed. ${total - failures.length} copied, ${failures.length} failed, ${missingMedia.length + skippedBySelection.length + skippedBySequenceScope.length + skippedByIgnoredTracks.length} skipped.${copiedProjectMessage}${reducedProjectMessage}`
    );

    renderList('errorList', failures, (item) => `${item.source} -> ${item.destination} | ${item.message}`);
    renderList('missingList', missingMedia.concat(skippedBySelection, skippedBySequenceScope, skippedByIgnoredTracks), (item) => item);
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
    try {
        sequenceOnlyMode = localStorage.getItem(SEQUENCE_ONLY_MODE_STORAGE_KEY) === '1';
        createReducedProject = localStorage.getItem(CREATE_REDUCED_PROJECT_STORAGE_KEY) === '1';
    } catch (error) {
        sequenceOnlyMode = false;
        createReducedProject = false;
    }
    try {
        const ignoreVisible = localStorage.getItem(IGNORE_SECTION_VISIBLE_STORAGE_KEY) === '1';
        setIgnoreSectionVisibility(ignoreVisible);
    } catch (error) {
        setIgnoreSectionVisibility(false);
    }
    setUpdateButton(`Version ${localVersion}`, false);
    checkForUpdates();
    syncSequenceModeUI();
    await loadProjectPlan();
});

window.addEventListener('focus', () => {
    refreshSourceList();
});
