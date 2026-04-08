var PROJECT_COLLECTOR_LOADED = true;

function pcJsonEscape(value) {
    if (value === null || value === undefined) {
        return "";
    }

    return String(value)
        .split('\\').join('\\\\')
        .split('"').join('\\"')
        .split('\r').join('\\r')
        .split('\n').join('\\n');
}

function pcJsonError(message) {
    return '{"error":"' + pcJsonEscape(message) + '"}';
}

function pcJoinPath(basePath, childName) {
    if (!basePath || basePath === '') {
        return childName;
    }
    return basePath + '/' + childName;
}

function pcSanitizeName(name) {
    var value = String(name || 'Untitled');
    value = value.replace(/[\\\/:\*\?"<>\|]/g, '_');
    value = value.replace(/^\s+/, '');
    value = value.replace(/\s+$/, '');
    return value;
}

function pcEnsureFolder(path) {
    var folder = new Folder(path);
    if (folder.exists) {
        return true;
    }

    if (folder.parent && !folder.parent.exists) {
        pcEnsureFolder(folder.parent.fsName);
    }

    return folder.create();
}

function pcIsBin(item) {
    try {
        if (item && item.type === ProjectItemType.BIN) {
            return true;
        }
    } catch (e) {}

    try {
        if (item && item.children && item.children.numItems !== undefined) {
            return true;
        }
    } catch (e2) {}

    return false;
}

function pcProjectName() {
    var projectName = '';

    try {
        if (app && app.project && app.project.name) {
            projectName = app.project.name;
        }
    } catch (e) {}

    if (!projectName) {
        projectName = 'Premiere_Project';
    }

    projectName = projectName.replace(/\.[^\.]+$/, '');
    return pcSanitizeName(projectName);
}

function pcPushFolder(folders, folderMap, relativePath) {
    if (!folderMap[relativePath]) {
        folderMap[relativePath] = true;
        folders.push(relativePath);
    }
}

function pcCollect(item, currentRelativePath, folders, folderMap, tasks, taskMap, missingMedia) {
    if (!item || !item.children || item.children.numItems === undefined) {
        return;
    }

    var i;
    for (i = 0; i < item.children.numItems; i++) {
        var child = item.children[i];

        if (pcIsBin(child)) {
            var binName = pcSanitizeName(child.name || ('Bin_' + i));
            var nextRelativePath = pcJoinPath(currentRelativePath, binName);
            pcPushFolder(folders, folderMap, nextRelativePath);
            pcCollect(child, nextRelativePath, folders, folderMap, tasks, taskMap, missingMedia);
            continue;
        }

        if (!child || !child.getMediaPath) {
            continue;
        }

        var mediaPath = '';
        try {
            mediaPath = child.getMediaPath();
        } catch (e3) {
            mediaPath = '';
        }

        if (!mediaPath || mediaPath === '') {
            missingMedia.push((child.name || 'Unknown Item') + ' | No media path available');
            continue;
        }

        var fileName = mediaPath.split(/[\\\/]/).pop();
        var relativeFilePath = pcJoinPath(currentRelativePath, pcSanitizeName(fileName));
        var uniqueKey = mediaPath + ' -> ' + relativeFilePath;

        if (taskMap[uniqueKey]) {
            continue;
        }

        taskMap[uniqueKey] = true;
        tasks.push({
            name: child.name || fileName,
            source: mediaPath,
            destination: relativeFilePath,
            binPath: currentRelativePath,
            relativePath: relativeFilePath
        });
    }
}

function pcTasksJson(tasks) {
    var out = [];
    var i;
    for (i = 0; i < tasks.length; i++) {
        var task = tasks[i];
        out.push(
            '{' +
            '"name":"' + pcJsonEscape(task.name) + '",' +
            '"source":"' + pcJsonEscape(task.source) + '",' +
            '"destination":"' + pcJsonEscape(task.destination) + '",' +
            '"binPath":"' + pcJsonEscape(task.binPath) + '",' +
            '"relativePath":"' + pcJsonEscape(task.relativePath) + '"' +
            '}'
        );
    }
    return '[' + out.join(',') + ']';
}

function pcStringsJson(items) {
    var out = [];
    var i;
    for (i = 0; i < items.length; i++) {
        out.push('"' + pcJsonEscape(items[i]) + '"');
    }
    return '[' + out.join(',') + ']';
}

function pcTrackUsageJson(items) {
    var out = [];
    var i;
    for (i = 0; i < items.length; i++) {
        var entry = items[i];
        out.push(
            '{' +
            '"trackNumber":' + entry.trackNumber + ',' +
            '"label":"' + pcJsonEscape(entry.label) + '",' +
            '"clipCount":' + entry.clipCount + ',' +
            '"mediaPaths":' + pcStringsJson(entry.mediaPaths) +
            '}'
        );
    }
    return '[' + out.join(',') + ']';
}

function pcTrackCollectionUsage(tracks, prefix) {
    var usage = [];
    var i;

    if (!tracks || tracks.numTracks === undefined) {
        return usage;
    }

    for (i = 0; i < tracks.numTracks; i++) {
        var mediaMap = {};
        var mediaPaths = [];
        var clipCount = 0;
        var track = tracks[i];
        var clips = null;
        var j;

        try {
            clips = track.clips;
        } catch (e) {
            clips = null;
        }

        if (clips && clips.numItems !== undefined) {
            for (j = 0; j < clips.numItems; j++) {
                var clip = clips[j];
                var projectItem = null;
                var mediaPath = '';

                try {
                    projectItem = clip.projectItem;
                } catch (e2) {
                    projectItem = null;
                }

                if (!projectItem || !projectItem.getMediaPath) {
                    continue;
                }

                try {
                    mediaPath = projectItem.getMediaPath();
                } catch (e3) {
                    mediaPath = '';
                }

                if (!mediaPath || mediaPath === '') {
                    continue;
                }

                clipCount += 1;

                if (!mediaMap[mediaPath]) {
                    mediaMap[mediaPath] = true;
                    mediaPaths.push(mediaPath);
                }
            }
        }

        usage.push({
            trackNumber: i + 1,
            label: prefix + (i + 1),
            clipCount: clipCount,
            mediaPaths: mediaPaths
        });
    }

    return usage;
}

function pcSequenceInfoJson(items) {
    var out = [];
    var i;
    for (i = 0; i < items.length; i++) {
        var entry = items[i];
        out.push(
            '{' +
            '"sequenceID":"' + pcJsonEscape(entry.sequenceID) + '",' +
            '"sequenceName":"' + pcJsonEscape(entry.sequenceName) + '"' +
            '}'
        );
    }
    return '[' + out.join(',') + ']';
}

function pcBuildSequenceMap() {
    var sequenceMap = {};
    var nodeMap = {};
    var i;

    if (!app || !app.project || !app.project.sequences) {
        return {
            bySequenceID: sequenceMap,
            byNodeId: nodeMap
        };
    }

    for (i = 0; i < app.project.sequences.numSequences; i++) {
        var sequence = app.project.sequences[i];
        if (!sequence) {
            continue;
        }

        try {
            if (sequence.sequenceID) {
                sequenceMap[sequence.sequenceID] = sequence;
            }
        } catch (e) {}

        try {
            if (sequence.projectItem && sequence.projectItem.nodeId) {
                nodeMap[sequence.projectItem.nodeId] = sequence;
            }
        } catch (e2) {}
    }

    return {
        bySequenceID: sequenceMap,
        byNodeId: nodeMap
    };
}

function pcFindSequenceByProjectItem(projectItem, sequenceMaps) {
    if (!projectItem || !sequenceMaps || !sequenceMaps.byNodeId) {
        return null;
    }

    try {
        if (projectItem.nodeId && sequenceMaps.byNodeId[projectItem.nodeId]) {
            return sequenceMaps.byNodeId[projectItem.nodeId];
        }
    } catch (e) {}

    return null;
}

function pcParseJsonArray(raw) {
    if (!raw || raw === '') {
        return [];
    }

    try {
        if (JSON && JSON.parse) {
            return JSON.parse(raw);
        }
    } catch (e) {}

    try {
        return eval('(' + raw + ')');
    } catch (e2) {}

    return [];
}

function pcBuildFilterMap(filters) {
    var map = {};
    var i;

    for (i = 0; i < filters.length; i++) {
        var filter = filters[i];
        if (!filter || !filter.sequenceID) {
            continue;
        }

        map[filter.sequenceID] = {
            ignoredVideoTracks: filter.ignoredVideoTracks || [],
            ignoredAudioTracks: filter.ignoredAudioTracks || []
        };
    }

    return map;
}

function pcArrayContainsInt(items, value) {
    var i;
    for (i = 0; i < items.length; i++) {
        if (parseInt(items[i], 10) === value) {
            return true;
        }
    }
    return false;
}

function pcCollectSequenceMedia(sequence, filterMap, sequenceMaps, visitedMap, mediaMap, mediaPaths, includedSequenceMap, includedSequences) {
    if (!sequence) {
        return;
    }

    var sequenceID = '';
    try {
        sequenceID = sequence.sequenceID || '';
    } catch (e) {}

    if (!sequenceID || visitedMap[sequenceID]) {
        return;
    }

    visitedMap[sequenceID] = true;
    includedSequenceMap[sequenceID] = true;
    includedSequences.push({
        sequenceID: sequenceID,
        sequenceName: sequence.name || 'Unknown Sequence'
    });

    var filter = filterMap[sequenceID] || {
        ignoredVideoTracks: [],
        ignoredAudioTracks: []
    };

    function collectTrackMedia(tracks, isVideo) {
        var i;
        if (!tracks || tracks.numTracks === undefined) {
            return;
        }

        for (i = 0; i < tracks.numTracks; i++) {
            var trackNumber = i + 1;
            var ignoredTracks = isVideo ? filter.ignoredVideoTracks : filter.ignoredAudioTracks;
            var track = tracks[i];
            var clips = null;
            var j;

            if (pcArrayContainsInt(ignoredTracks, trackNumber)) {
                continue;
            }

            try {
                clips = track.clips;
            } catch (e2) {
                clips = null;
            }

            if (!clips || clips.numItems === undefined) {
                continue;
            }

            for (j = 0; j < clips.numItems; j++) {
                var clip = clips[j];
                var projectItem = null;
                var mediaPath = '';
                var nestedSequence = null;
                var isSequenceItem = false;

                try {
                    projectItem = clip.projectItem;
                } catch (e3) {
                    projectItem = null;
                }

                if (!projectItem) {
                    continue;
                }

                try {
                    isSequenceItem = projectItem.isSequence();
                } catch (e4) {
                    isSequenceItem = false;
                }

                if (isSequenceItem) {
                    nestedSequence = pcFindSequenceByProjectItem(projectItem, sequenceMaps);
                    if (nestedSequence) {
                        pcCollectSequenceMedia(nestedSequence, filterMap, sequenceMaps, visitedMap, mediaMap, mediaPaths, includedSequenceMap, includedSequences);
                        continue;
                    }
                }

                if (!projectItem.getMediaPath) {
                    continue;
                }

                try {
                    mediaPath = projectItem.getMediaPath();
                } catch (e5) {
                    mediaPath = '';
                }

                if (!mediaPath || mediaPath === '') {
                    continue;
                }

                if (!mediaMap[mediaPath]) {
                    mediaMap[mediaPath] = true;
                    mediaPaths.push(mediaPath);
                }
            }
        }
    }

    collectTrackMedia(sequence.videoTracks, true);
    collectTrackMedia(sequence.audioTracks, false);
}

function pcBuildSequenceScopedPlan(filters) {
    var sequenceMaps = pcBuildSequenceMap();
    var filterMap = pcBuildFilterMap(filters);
    var visitedMap = {};
    var mediaMap = {};
    var mediaPaths = [];
    var includedSequenceMap = {};
    var includedSequences = [];
    var missingSequences = [];
    var selectedSequenceIDs = [];
    var i;

    for (i = 0; i < filters.length; i++) {
        var filter = filters[i];
        var sequence = null;

        if (!filter || !filter.sequenceID) {
            continue;
        }

        selectedSequenceIDs.push(filter.sequenceID);
        sequence = sequenceMaps.bySequenceID[filter.sequenceID];

        if (!sequence) {
            missingSequences.push(filter.sequenceName || filter.sequenceID);
            continue;
        }

        pcCollectSequenceMedia(sequence, filterMap, sequenceMaps, visitedMap, mediaMap, mediaPaths, includedSequenceMap, includedSequences);
    }

    return {
        mediaPaths: mediaPaths,
        includedSequences: includedSequences,
        missingSequences: missingSequences,
        selectedSequenceIDs: selectedSequenceIDs
    };
}

function pcBuildPlan(destination) {
    if (!app || !app.project || !app.project.rootItem) {
        throw new Error('No Premiere project is currently open.');
    }

    var rootName = pcProjectName();
    var rootPath = pcJoinPath(destination, rootName);
    var folders = [''];
    var folderMap = { '': true };
    var tasks = [];
    var taskMap = {};
    var missingMedia = [];
    var activeSequenceName = '';
    var activeSequenceID = '';
    var videoTrackUsage = [];
    var audioTrackUsage = [];

    pcCollect(app.project.rootItem, '', folders, folderMap, tasks, taskMap, missingMedia);

    try {
        if (app.project.activeSequence) {
            activeSequenceName = app.project.activeSequence.name || '';
            activeSequenceID = app.project.activeSequence.sequenceID || '';
            videoTrackUsage = pcTrackCollectionUsage(app.project.activeSequence.videoTracks, 'V');
            audioTrackUsage = pcTrackCollectionUsage(app.project.activeSequence.audioTracks, 'A');
        }
    } catch (e4) {}

    return {
        projectName: rootName,
        rootPath: rootPath,
        folders: folders,
        tasks: tasks,
        missingMedia: missingMedia,
        activeSequenceName: activeSequenceName,
        activeSequenceID: activeSequenceID,
        videoTrackUsage: videoTrackUsage,
        audioTrackUsage: audioTrackUsage
    };
}

function getProjectCopyPlan(destination) {
    try {
        var plan = pcBuildPlan(destination);
        return '{' +
            '"projectName":"' + pcJsonEscape(plan.projectName) + '",' +
            '"rootPath":"' + pcJsonEscape(plan.rootPath) + '",' +
            '"folders":' + pcStringsJson(plan.folders) + ',' +
            '"tasks":' + pcTasksJson(plan.tasks) + ',' +
            '"missingMedia":' + pcStringsJson(plan.missingMedia) + ',' +
            '"activeSequenceName":"' + pcJsonEscape(plan.activeSequenceName) + '",' +
            '"activeSequenceID":"' + pcJsonEscape(plan.activeSequenceID) + '",' +
            '"videoTrackUsage":' + pcTrackUsageJson(plan.videoTrackUsage) + ',' +
            '"audioTrackUsage":' + pcTrackUsageJson(plan.audioTrackUsage) +
            '}';
    } catch (e) {
        return pcJsonError(e.toString());
    }
}

function getActiveSequenceTrackUsage() {
    try {
        var activeSequenceName = '';
        var activeSequenceID = '';
        var videoTrackUsage = [];
        var audioTrackUsage = [];

        if (app && app.project && app.project.activeSequence) {
            activeSequenceName = app.project.activeSequence.name || '';
            activeSequenceID = app.project.activeSequence.sequenceID || '';
            videoTrackUsage = pcTrackCollectionUsage(app.project.activeSequence.videoTracks, 'V');
            audioTrackUsage = pcTrackCollectionUsage(app.project.activeSequence.audioTracks, 'A');
        }

        return '{' +
            '"sequenceName":"' + pcJsonEscape(activeSequenceName) + '",' +
            '"sequenceID":"' + pcJsonEscape(activeSequenceID) + '",' +
            '"videoTrackUsage":' + pcTrackUsageJson(videoTrackUsage) + ',' +
            '"audioTrackUsage":' + pcTrackUsageJson(audioTrackUsage) +
            '}';
    } catch (e) {
        return pcJsonError(e.toString());
    }
}

function getSequenceScopedMediaPlan(filtersJson) {
    try {
        var filters = pcParseJsonArray(filtersJson);
        var plan = pcBuildSequenceScopedPlan(filters);

        return '{' +
            '"mediaPaths":' + pcStringsJson(plan.mediaPaths) + ',' +
            '"includedSequences":' + pcSequenceInfoJson(plan.includedSequences) + ',' +
            '"includedSequenceIDs":' + pcStringsJson((function () {
                var ids = [];
                var i;
                for (i = 0; i < plan.includedSequences.length; i++) {
                    ids.push(plan.includedSequences[i].sequenceID);
                }
                return ids;
            }())) + ',' +
            '"missingSequences":' + pcStringsJson(plan.missingSequences) + ',' +
            '"selectedSequenceIDs":' + pcStringsJson(plan.selectedSequenceIDs) +
            '}';
    } catch (e) {
        return pcJsonError(e.toString());
    }
}

function createReducedProjectFromSequenceSelection(destinationFolder, sequenceIDsJson) {
    try {
        if (!app || !app.project) {
            throw new Error('No Premiere project is currently open.');
        }

        var originalProjectPath = app.project.path || '';
        var originalProjectName = pcProjectName();
        var sequenceIDs = pcParseJsonArray(sequenceIDsJson);

        if (!originalProjectPath || originalProjectPath === '') {
            throw new Error('Save the original Premiere project first before creating a reduced project copy.');
        }

        if (!sequenceIDs || !sequenceIDs.length) {
            throw new Error('No sequence IDs were provided for the reduced project.');
        }

        pcEnsureFolder(destinationFolder);

        var reducedProjectPath = destinationFolder + '/' + pcSanitizeName(originalProjectName + '_Selected_Sequences.prproj');
        var createResult = app.newProject(reducedProjectPath);
        var importResult = app.project.importSequences(originalProjectPath, sequenceIDs);
        var saveResult = app.project.save();
        var reopenResult = false;

        try {
            reopenResult = app.openDocument(originalProjectPath, true, true, true, true);
        } catch (e2) {
            reopenResult = false;
        }

        return '{' +
            '"success":true,' +
            '"reducedProjectPath":"' + pcJsonEscape(reducedProjectPath) + '",' +
            '"created":"' + pcJsonEscape(createResult) + '",' +
            '"imported":"' + pcJsonEscape(importResult) + '",' +
            '"saved":"' + pcJsonEscape(saveResult) + '",' +
            '"reopenedOriginal":"' + pcJsonEscape(reopenResult) + '"' +
            '}';
    } catch (e) {
        return pcJsonError(e.toString());
    }
}

function prepareProjectStructure(destination) {
    try {
        var plan = pcBuildPlan(destination);
        var created = 0;
        var i;

        for (i = 0; i < plan.folders.length; i++) {
            var relativeFolder = plan.folders[i];
            var fullPath = relativeFolder ? pcJoinPath(plan.rootPath, relativeFolder) : plan.rootPath;
            if (pcEnsureFolder(fullPath)) {
                created += 1;
            }
        }

        return '{"rootPath":"' + pcJsonEscape(plan.rootPath) + '","createdCount":' + created + '}';
    } catch (e2) {
        return pcJsonError(e2.toString());
    }
}

function copyPlannedFile(sourcePath, destinationPath) {
    try {
        var srcFile = new File(sourcePath);
        var dstFile = new File(destinationPath);
        var copied = false;
        var message = '';

        if (!srcFile.exists) {
            message = 'Source file does not exist';
        } else {
            pcEnsureFolder(dstFile.parent.fsName);

            if (dstFile.exists) {
                try {
                    dstFile.remove();
                } catch (e3) {}
            }

            copied = srcFile.copy(destinationPath);
            if (!copied) {
                message = srcFile.error || dstFile.error || 'Copy failed';
            }
        }

        return '{' +
            '"success":' + (copied ? 'true' : 'false') + ',' +
            '"source":"' + pcJsonEscape(sourcePath) + '",' +
            '"destination":"' + pcJsonEscape(destinationPath) + '",' +
            '"message":"' + pcJsonEscape(message) + '"' +
            '}';
    } catch (e4) {
        return pcJsonError(e4.toString());
    }
}
