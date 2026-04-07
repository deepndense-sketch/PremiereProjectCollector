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
let listVisible = true;
let selectionTouched = false;
let localVersion = 'unknown';
let remoteVersion = null;
let ignoredVideoTracks = [];
let ignoredAudioTracks = [];

const IGNORED_VIDEO_TRACKS_STORAGE_KEY = 'projectcollector.ignoredVideoTracks';
const IGNORED_AUDIO_TRACKS_STORAGE_KEY = 'projectcollector.ignoredAudioTracks';

function setText(id, value) {
    document.getElementById(id).textContent = value;
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
    const remoteUrl = 'https://raw.githubusercontent.com/deepndense-sketch/PremiereProjectCollector/main/version.json';
    setUpdateButton(`Version ${localVersion}`, false);

    try {
        const remote = await new Promise((resolve, reject) => {
            https.get(remoteUrl, (response) => {
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
            }).on('error', reject);
        });

        remoteVersion = remote.version || 'unknown';
        if (compareVersions(remoteVersion, localVersion) > 0) {
            setUpdateButton(`Update to ${remoteVersion}`, true);
        } else {
            setUpdateButton(`Version ${localVersion}`, false);
        }
    } catch (error) {
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
    const ignoredTrackSummary = [];

    if (ignoredVideoTracks.length) {
        ignoredTrackSummary.push(`ignored video: ${ignoredVideoTracks.map((trackNumber) => `V${trackNumber}`).join(', ')}`);
    }

    if (ignoredAudioTracks.length) {
        ignoredTrackSummary.push(`ignored audio: ${ignoredAudioTracks.map((trackNumber) => `A${trackNumber}`).join(', ')}`);
    }

    const suffix = ignoredTrackSummary.length ? ` (${ignoredTrackSummary.join(' | ')})` : '';

    if (!selectionTouched) {
        setText('selectionSummary', `All ${total} files will be included by default. Once you change the list, only the checked items will be copied.${suffix}`);
        return;
    }

    if (included === 0) {
        setText('selectionSummary', `No files are selected. Copy will process zero files until you check items again.${suffix}`);
        return;
    }

    setText('selectionSummary', `${included} of ${total} files are selected for copy.${suffix}`);
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

function loadIgnoredTrackState() {
    try {
        const savedVideo = JSON.parse(localStorage.getItem(IGNORED_VIDEO_TRACKS_STORAGE_KEY) || '[]');
        ignoredVideoTracks = Array.isArray(savedVideo) ? savedVideo : [];
    } catch (error) {
        ignoredVideoTracks = [];
    }

    try {
        const savedAudio = JSON.parse(localStorage.getItem(IGNORED_AUDIO_TRACKS_STORAGE_KEY) || '[]');
        ignoredAudioTracks = Array.isArray(savedAudio) ? savedAudio : [];
    } catch (error) {
        ignoredAudioTracks = [];
    }
}

function saveIgnoredTrackState() {
    try {
        localStorage.setItem(IGNORED_VIDEO_TRACKS_STORAGE_KEY, JSON.stringify(ignoredVideoTracks));
        localStorage.setItem(IGNORED_AUDIO_TRACKS_STORAGE_KEY, JSON.stringify(ignoredAudioTracks));
    } catch (error) {}
}

function getTrackUsageEntries(kind) {
    if (!latestPlan) {
        return [];
    }

    return kind === 'video'
        ? (latestPlan.videoTrackUsage || [])
        : (latestPlan.audioTrackUsage || []);
}

function renderIgnoredTrackChips(kind) {
    const container = document.getElementById(kind === 'video' ? 'ignoredVideoTracks' : 'ignoredAudioTracks');
    const values = kind === 'video' ? ignoredVideoTracks : ignoredAudioTracks;
    container.innerHTML = '';

    if (!values.length) {
        const empty = document.createElement('div');
        empty.className = 'small-note';
        empty.textContent = kind === 'video' ? 'No ignored video tracks.' : 'No ignored audio tracks.';
        container.appendChild(empty);
        return;
    }

    values.forEach((trackNumber) => {
        const chip = document.createElement('div');
        chip.className = 'chip';
        chip.textContent = `${kind === 'video' ? 'Ignore Video Medias on Layer V' : 'Ignore Audio Medias on Layer A'}${trackNumber}`;

        const removeButton = document.createElement('button');
        removeButton.type = 'button';
        removeButton.textContent = 'x';
        removeButton.onclick = () => removeIgnoredTrack(kind, trackNumber);
        chip.appendChild(removeButton);
        container.appendChild(chip);
    });
}

function renderTrackSelect(kind) {
    const select = document.getElementById(kind === 'video' ? 'ignoreVideoTrackSelect' : 'ignoreAudioTrackSelect');
    const ignored = kind === 'video' ? ignoredVideoTracks : ignoredAudioTracks;
    const entries = getTrackUsageEntries(kind);
    select.innerHTML = '';

    const availableEntries = entries.filter((entry) => ignored.indexOf(entry.trackNumber) === -1);

    if (!availableEntries.length) {
        const option = document.createElement('option');
        option.value = '';
        option.textContent = kind === 'video' ? 'No video tracks left' : 'No audio tracks left';
        select.appendChild(option);
        select.disabled = true;
        return;
    }

    availableEntries.forEach((entry) => {
        const option = document.createElement('option');
        option.value = String(entry.trackNumber);
        option.textContent = `${entry.label} (${entry.clipCount} clips)`;
        select.appendChild(option);
    });

    select.disabled = false;
}

function renderTrackFilters() {
    renderTrackSelect('video');
    renderTrackSelect('audio');
    renderIgnoredTrackChips('video');
    renderIgnoredTrackChips('audio');
}

function addIgnoredTrack(kind) {
    const select = document.getElementById(kind === 'video' ? 'ignoreVideoTrackSelect' : 'ignoreAudioTrackSelect');
    const value = parseInt(select.value, 10) || 0;
    if (!value) {
        return;
    }

    if (kind === 'video') {
        if (ignoredVideoTracks.indexOf(value) === -1) {
            ignoredVideoTracks.push(value);
            ignoredVideoTracks.sort((a, b) => a - b);
        }
    } else {
        if (ignoredAudioTracks.indexOf(value) === -1) {
            ignoredAudioTracks.push(value);
            ignoredAudioTracks.sort((a, b) => a - b);
        }
    }

    saveIgnoredTrackState();
    renderTrackFilters();
    updateSelectionSummary();
}

function removeIgnoredTrack(kind, trackNumber) {
    if (kind === 'video') {
        ignoredVideoTracks = ignoredVideoTracks.filter((value) => value !== trackNumber);
    } else {
        ignoredAudioTracks = ignoredAudioTracks.filter((value) => value !== trackNumber);
    }

    saveIgnoredTrackState();
    renderTrackFilters();
    updateSelectionSummary();
}

function buildIgnoredMediaSet() {
    const ignoredPaths = new Set();

    getTrackUsageEntries('video').forEach((entry) => {
        if (ignoredVideoTracks.indexOf(entry.trackNumber) !== -1) {
            (entry.mediaPaths || []).forEach((mediaPath) => {
                ignoredPaths.add(normalizeMediaKey(mediaPath));
            });
        }
    });

    getTrackUsageEntries('audio').forEach((entry) => {
        if (ignoredAudioTracks.indexOf(entry.trackNumber) !== -1) {
            (entry.mediaPaths || []).forEach((mediaPath) => {
                ignoredPaths.add(normalizeMediaKey(mediaPath));
            });
        }
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
    renderTrackFilters();
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
    const ignoredMediaSet = buildIgnoredMediaSet();
    const treeSelectedTasks = getSelectedTasks();
    const selectedTasks = treeSelectedTasks.filter((task) => !ignoredMediaSet.has(normalizeMediaKey(task.source)));
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
    const skippedByIgnoredTracks = (latestPlan.tasks || [])
        .filter((task) => ignoredMediaSet.has(normalizeMediaKey(task.source)))
        .map((task) => `${task.source} -> skipped by ignored track filter`);

    setText('summaryText', `Windows robocopy mode active. Copying into ${plan.rootPath}`);
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

    setText('currentFile', total ? 'Copy finished' : 'No copyable media found');
    setText(
        'summaryText',
        `Completed. ${total - failures.length} copied, ${failures.length} failed, ${missingMedia.length + skippedBySelection.length + skippedByIgnoredTracks.length} skipped.`
    );

    renderList('errorList', failures, (item) => `${item.source} -> ${item.destination} | ${item.message}`);
    renderList('missingList', missingMedia.concat(skippedBySelection, skippedByIgnoredTracks), (item) => item);
    setBusyState(false);
}

document.addEventListener('DOMContentLoaded', async () => {
    readVersionInfo();
    loadIgnoredTrackState();
    resetResults();
    document.getElementById('sourceListBox').style.display = 'block';
    setText('showListButton', 'Hide List');
    setUpdateButton(`Version ${localVersion}`, false);
    checkForUpdates();
    renderTrackFilters();
    await loadProjectPlan();
});

window.addEventListener('focus', () => {
    refreshSourceList();
});
