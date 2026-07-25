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
            '"hasItems":' + (entry.hasItems ? 'true' : 'false') + ',' +
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
        var hasItems = false;
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

                hasItems = true;
                clipCount += 1;

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
            hasItems: hasItems,
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

function pcSequenceTrackUsageJson(items) {
    var out = [];
    var i;
    for (i = 0; i < items.length; i++) {
        var entry = items[i];
        out.push(
            '{' +
            '"sequenceID":"' + pcJsonEscape(entry.sequenceID) + '",' +
            '"sequenceName":"' + pcJsonEscape(entry.sequenceName) + '",' +
            '"videoTrackUsage":' + pcTrackUsageJson(entry.videoTrackUsage) + ',' +
            '"audioTrackUsage":' + pcTrackUsageJson(entry.audioTrackUsage) +
            '}'
        );
    }
    return '[' + out.join(',') + ']';
}

function pcBuildSequenceMap() {
    var sequenceMap = {};
    var nameMap = {};
    var nodeMap = {};
    var i;

    if (!app || !app.project || !app.project.sequences) {
        return {
            bySequenceID: sequenceMap,
            bySequenceName: nameMap,
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
            if (sequence.name) {
                nameMap[String(sequence.name).toLowerCase()] = sequence;
            }
        } catch (eName) {}

        try {
            if (sequence.projectItem && sequence.projectItem.nodeId) {
                nodeMap[sequence.projectItem.nodeId] = sequence;
            }
        } catch (e2) {}
    }

    return {
        bySequenceID: sequenceMap,
        bySequenceName: nameMap,
        byNodeId: nodeMap
    };
}

function pcFindSequenceFromFilter(filter, sequenceMaps) {
    if (!filter || !sequenceMaps) {
        return null;
    }

    if (filter.sequenceID && sequenceMaps.bySequenceID && sequenceMaps.bySequenceID[filter.sequenceID]) {
        return sequenceMaps.bySequenceID[filter.sequenceID];
    }

    if (filter.sequenceName && sequenceMaps.bySequenceName) {
        return sequenceMaps.bySequenceName[String(filter.sequenceName).toLowerCase()] || null;
    }

    return null;
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

function pcBuildFilterMap(filters, sequenceMaps) {
    var map = {};
    var i;

    for (i = 0; i < filters.length; i++) {
        var filter = filters[i];
        var sequence = pcFindSequenceFromFilter(filter, sequenceMaps);
        var sequenceID = '';

        if (!filter || !sequence) {
            continue;
        }

        try {
            sequenceID = sequence.sequenceID || filter.sequenceID || '';
        } catch (e) {
            sequenceID = filter.sequenceID || '';
        }

        if (sequenceID) {
            map[sequenceID] = {
                ignoredVideoTracks: filter.ignoredVideoTracks || [],
                ignoredAudioTracks: filter.ignoredAudioTracks || []
            };
        }
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

function pcAddMediaPath(mediaMap, mediaPaths, mediaPath) {
    if (!mediaPath || mediaPath === '') {
        return;
    }

    if (!mediaMap[mediaPath]) {
        mediaMap[mediaPath] = true;
        mediaPaths.push(mediaPath);
    }
}

function pcCollectProjectItemMedia(projectItem, sequenceMaps, visitedSequenceMap, mediaMap, mediaPaths) {
    if (!projectItem) {
        return;
    }

    var nestedSequence = null;
    var isSequenceItem = false;
    var mediaPath = '';

    try {
        isSequenceItem = projectItem.isSequence();
    } catch (e) {
        isSequenceItem = false;
    }

    if (isSequenceItem) {
        nestedSequence = pcFindSequenceByProjectItem(projectItem, sequenceMaps);
        if (nestedSequence) {
            pcCollectAllSequenceMedia(nestedSequence, sequenceMaps, visitedSequenceMap, mediaMap, mediaPaths);
        }
        return;
    }

    if (!projectItem.getMediaPath) {
        return;
    }

    try {
        mediaPath = projectItem.getMediaPath();
    } catch (e2) {
        mediaPath = '';
    }

    pcAddMediaPath(mediaMap, mediaPaths, mediaPath);
}

function pcCollectAllSequenceMedia(sequence, sequenceMaps, visitedSequenceMap, mediaMap, mediaPaths) {
    if (!sequence) {
        return;
    }

    var sequenceID = '';
    try {
        sequenceID = sequence.sequenceID || '';
    } catch (e) {}

    if (sequenceID && visitedSequenceMap[sequenceID]) {
        return;
    }

    if (sequenceID) {
        visitedSequenceMap[sequenceID] = true;
    }

    function collectTracks(tracks) {
        var i;
        if (!tracks || tracks.numTracks === undefined) {
            return;
        }

        for (i = 0; i < tracks.numTracks; i++) {
            var track = tracks[i];
            var clips = null;
            var j;

            try {
                clips = track.clips;
            } catch (e2) {
                clips = null;
            }

            if (!clips || clips.numItems === undefined) {
                continue;
            }

            for (j = 0; j < clips.numItems; j++) {
                var projectItem = null;

                try {
                    projectItem = clips[j].projectItem;
                } catch (e3) {
                    projectItem = null;
                }

                pcCollectProjectItemMedia(projectItem, sequenceMaps, visitedSequenceMap, mediaMap, mediaPaths);
            }
        }
    }

    collectTracks(sequence.videoTracks);
    collectTracks(sequence.audioTracks);
}

function pcCollectIgnoredTrackMedia(sequence, filter, sequenceMaps, visitedSequenceMap, mediaMap, mediaPaths) {
    if (!sequence || !filter) {
        return;
    }

    function collectSelectedTracks(tracks, ignoredTracks) {
        var i;
        if (!tracks || tracks.numTracks === undefined) {
            return;
        }

        for (i = 0; i < tracks.numTracks; i++) {
            var trackNumber = i + 1;
            var track = tracks[i];
            var clips = null;
            var j;

            if (!pcArrayContainsInt(ignoredTracks, trackNumber)) {
                continue;
            }

            try {
                clips = track.clips;
            } catch (e) {
                clips = null;
            }

            if (!clips || clips.numItems === undefined) {
                continue;
            }

            for (j = 0; j < clips.numItems; j++) {
                var projectItem = null;

                try {
                    projectItem = clips[j].projectItem;
                } catch (e2) {
                    projectItem = null;
                }

                pcCollectProjectItemMedia(projectItem, sequenceMaps, visitedSequenceMap, mediaMap, mediaPaths);
            }
        }
    }

    collectSelectedTracks(sequence.videoTracks, filter.ignoredVideoTracks || []);
    collectSelectedTracks(sequence.audioTracks, filter.ignoredAudioTracks || []);
}

function pcBuildSequenceScopedPlan(filters) {
    var sequenceMaps = pcBuildSequenceMap();
    var filterMap = pcBuildFilterMap(filters, sequenceMaps);
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

        if (!filter) {
            continue;
        }

        sequence = pcFindSequenceFromFilter(filter, sequenceMaps);

        if (!sequence) {
            missingSequences.push(filter.sequenceName || filter.sequenceID);
            continue;
        }

        try {
            selectedSequenceIDs.push(sequence.sequenceID || filter.sequenceID);
        } catch (eSequenceID) {
            selectedSequenceIDs.push(filter.sequenceID);
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

function pcBuildIgnoredTrackMediaPlan(filters) {
    var sequenceMaps = pcBuildSequenceMap();
    var mediaMap = {};
    var mediaPaths = [];
    var visitedSequenceMap = {};
    var missingSequences = [];
    var i;

    for (i = 0; i < filters.length; i++) {
        var filter = filters[i];
        var sequence = null;

        if (!filter) {
            continue;
        }

        sequence = pcFindSequenceFromFilter(filter, sequenceMaps);
        if (!sequence) {
            missingSequences.push(filter.sequenceName || filter.sequenceID);
            continue;
        }

        pcCollectIgnoredTrackMedia(sequence, filter, sequenceMaps, visitedSequenceMap, mediaMap, mediaPaths);
    }

    return {
        mediaPaths: mediaPaths,
        missingSequences: missingSequences
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
    var projectPath = '';
    var activeSequenceName = '';
    var activeSequenceID = '';
    var videoTrackUsage = [];
    var audioTrackUsage = [];
    var availableSequences = [];

    pcCollect(app.project.rootItem, '', folders, folderMap, tasks, taskMap, missingMedia);

    try {
        projectPath = app.project.path || '';
    } catch (eProjectPath) {}

    try {
        if (app.project.sequences) {
            var sequenceIndex;
            for (sequenceIndex = 0; sequenceIndex < app.project.sequences.numSequences; sequenceIndex++) {
                var availableSequence = app.project.sequences[sequenceIndex];
                if (availableSequence) {
                    availableSequences.push({
                        sequenceID: availableSequence.sequenceID || '',
                        sequenceName: availableSequence.name || ''
                    });
                }
            }
        }
    } catch (eSequences) {}

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
        projectPath: projectPath,
        availableSequences: availableSequences,
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
            '"projectPath":"' + pcJsonEscape(plan.projectPath) + '",' +
            '"availableSequences":' + pcSequenceInfoJson(plan.availableSequences) + ',' +
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

function getSequenceTrackUsagePlan(filtersJson) {
    try {
        var filters = pcParseJsonArray(filtersJson);
        var sequenceMaps = pcBuildSequenceMap();
        var sequences = [];
        var missingSequences = [];
        var seenSequences = {};
        var i;

        for (i = 0; i < filters.length; i++) {
            var filter = filters[i];
            var sequence = pcFindSequenceFromFilter(filter, sequenceMaps);
            if (!sequence) {
                missingSequences.push(filter.sequenceName || filter.sequenceID || 'Unknown Sequence');
                continue;
            }

            var sequenceID = '';
            var sequenceName = '';
            try {
                sequenceID = sequence.sequenceID || filter.sequenceID || '';
            } catch (eID) {
                sequenceID = filter.sequenceID || '';
            }
            try {
                sequenceName = sequence.name || filter.sequenceName || '';
            } catch (eName) {
                sequenceName = filter.sequenceName || '';
            }

            var sequenceKey = sequenceID
                ? 'id:' + sequenceID
                : 'name:' + String(sequenceName).toLowerCase();
            if (seenSequences[sequenceKey]) {
                continue;
            }
            seenSequences[sequenceKey] = true;

            sequences.push({
                sequenceID: sequenceID,
                sequenceName: sequenceName,
                videoTrackUsage: pcTrackCollectionUsage(sequence.videoTracks, 'V'),
                audioTrackUsage: pcTrackCollectionUsage(sequence.audioTracks, 'A')
            });
        }

        return '{' +
            '"sequences":' + pcSequenceTrackUsageJson(sequences) + ',' +
            '"missingSequences":' + pcStringsJson(missingSequences) +
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

function getIgnoredTrackMediaPlan(filtersJson) {
    try {
        var filters = pcParseJsonArray(filtersJson);
        var plan = pcBuildIgnoredTrackMediaPlan(filters);

        return '{' +
            '"mediaPaths":' + pcStringsJson(plan.mediaPaths) + ',' +
            '"missingSequences":' + pcStringsJson(plan.missingSequences) +
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

function saveCurrentProjectAndGetPath() {
    try {
        if (!app || !app.project) {
            throw new Error('No Premiere project is currently open.');
        }

        var projectPath = app.project.path || '';
        if (!projectPath || projectPath === '') {
            throw new Error('Save the Premiere project first before copying the project file.');
        }

        app.project.save();

        return '{' +
            '"projectPath":"' + pcJsonEscape(projectPath) + '"' +
            '}';
    } catch (e) {
        return pcJsonError(e.toString());
    }
}

function pcNormalizeRelinkPath(pathValue) {
    return String(pathValue || '').split('\\').join('/').toLowerCase();
}

function pcBuildRelinkMap(tasks) {
    var map = {};
    var i;

    for (i = 0; i < tasks.length; i++) {
        var task = tasks[i];
        if (!task || !task.source || !task.destination) {
            continue;
        }

        map[pcNormalizeRelinkPath(task.source)] = {
            destination: task.destination,
            targetKind: task.targetKind === 'collected'
                ? 'collected'
                : (task.targetKind === 'skip-location' ? 'skip-location' : 'original')
        };
    }

    return map;
}

function pcCollectRelinkProjectItems(item, relinkMap, records, result) {
    if (!item || !item.children || item.children.numItems === undefined) {
        return;
    }

    var i;
    for (i = 0; i < item.children.numItems; i++) {
        var child = item.children[i];

        if (pcIsBin(child)) {
            pcCollectRelinkProjectItems(child, relinkMap, records, result);
            continue;
        }

        if (!child || !child.getMediaPath) {
            continue;
        }

        var currentPath = '';
        try {
            currentPath = child.getMediaPath();
        } catch (e) {
            result.failed.push((child.name || 'Unnamed project item') + ' | could not read original media path: ' + e.toString());
            continue;
        }

        if (!currentPath || currentPath === '') {
            result.skippedNoPathCount += 1;
            continue;
        }

        var relinkTask = relinkMap[pcNormalizeRelinkPath(currentPath)];
        records.push({
            item: child,
            name: child.name || currentPath,
            source: currentPath,
            destination: relinkTask && relinkTask.destination ? relinkTask.destination : currentPath,
            targetKind: relinkTask && relinkTask.targetKind === 'collected'
                ? 'collected'
                : (relinkTask && relinkTask.targetKind === 'skip-location' ? 'skip-location' : 'original')
        });
    }
}

function pcOfflineRelinkRecords(records, result) {
    var i;

    for (i = 0; i < records.length; i++) {
        var record = records[i];

        if (!record.item || !record.item.setOffline) {
            result.failed.push(record.name + ' | Premiere does not allow this file-backed item to be taken offline');
            continue;
        }

        try {
            var alreadyOffline = false;
            if (record.item.isOffline) {
                try {
                    alreadyOffline = record.item.isOffline() === true;
                } catch (initialOfflineCheckError) {}
            }

            if (alreadyOffline) {
                result.offlineCount += 1;
                continue;
            }

            var offlineResult = record.item.setOffline();
            var isOffline = offlineResult === true;

            if (record.item.isOffline) {
                try {
                    isOffline = record.item.isOffline() === true;
                } catch (offlineCheckError) {}
            }

            if (!isOffline) {
                result.failed.push(record.name + ' | Premiere could not take the item offline');
                continue;
            }

            result.offlineCount += 1;
        } catch (e) {
            result.failed.push(record.name + ' | could not take item offline: ' + e.toString());
        }
    }
}

function pcRelinkOfflineRecords(records, result) {
    var i;

    for (i = 0; i < records.length; i++) {
        var record = records[i];

        if (!record.item || !record.item.changeMediaPath) {
            result.failed.push(record.name + ' | Premiere does not allow this item to change media path');
            continue;
        }

        try {
            var changeResult = record.item.changeMediaPath(record.destination, true);
            var linkedPath = '';
            var linkedSuccessfully = changeResult === 0 || changeResult === true;

            try {
                linkedPath = record.item.getMediaPath();
            } catch (pathCheckError) {
                linkedPath = '';
            }

            if (linkedPath && pcNormalizeRelinkPath(linkedPath) === pcNormalizeRelinkPath(record.destination)) {
                linkedSuccessfully = true;
            } else if (linkedPath) {
                linkedSuccessfully = false;
            }

            if (!linkedSuccessfully) {
                result.failed.push(record.name + ' | Premiere could not link to ' + record.destination);
                continue;
            }

            result.linkedCount += 1;
            if (record.targetKind === 'collected') {
                result.linkedCollectedCount += 1;
            } else if (record.targetKind === 'skip-location') {
                result.linkedSkipLocationCount += 1;
            } else {
                result.linkedOriginalCount += 1;
            }
        } catch (e) {
            result.failed.push(record.name + ' | could not link to ' + record.destination + ': ' + e.toString());
        }
    }
}

function pcFindOpenProjectByPath(projectPath) {
    var normalizedTarget = pcNormalizeRelinkPath(projectPath);
    var activeProject = null;

    try {
        activeProject = app && app.project ? app.project : null;
        if (activeProject && pcNormalizeRelinkPath(activeProject.path || '') === normalizedTarget) {
            return activeProject;
        }
    } catch (activeProjectError) {}

    var projectCollections = [];
    try {
        if (app && app.projects) {
            projectCollections.push(app.projects);
        }
    } catch (projectsError) {}
    try {
        if (app && app.production && app.production.projects) {
            projectCollections.push(app.production.projects);
        }
    } catch (productionProjectsError) {}

    var collectionIndex;
    for (collectionIndex = 0; collectionIndex < projectCollections.length; collectionIndex++) {
        var projects = projectCollections[collectionIndex];
        var projectCount = 0;
        try {
            projectCount = projects.numProjects !== undefined ? projects.numProjects : projects.length;
        } catch (projectCountError) {
            projectCount = 0;
        }

        var projectIndex;
        for (projectIndex = 0; projectIndex <= projectCount; projectIndex++) {
            var candidate = null;
            try {
                candidate = projects[projectIndex];
            } catch (candidateError) {
                candidate = null;
            }

            try {
                if (candidate && pcNormalizeRelinkPath(candidate.path || '') === normalizedTarget) {
                    return candidate;
                }
            } catch (candidatePathError) {}
        }
    }

    return null;
}

function pcCloseCopiedProjectAndRestoreOriginal(copiedProject, copiedProjectPath, originalProjectPath, result) {
    if (copiedProject && copiedProject.closeDocument) {
        try {
            var closeResult = copiedProject.closeDocument(0, 0);
            if (closeResult !== 0 && closeResult !== true && closeResult !== undefined) {
                result.failed.push('Premiere could not close the BACKUP project after relinking.');
                result.success = false;
            }
        } catch (closeError) {
            result.failed.push('Could not close the BACKUP project: ' + closeError.toString());
            result.success = false;
        }
    }

    if (originalProjectPath && pcNormalizeRelinkPath(originalProjectPath) !== pcNormalizeRelinkPath(copiedProjectPath)) {
        try {
            var activeProjectPath = app && app.project ? (app.project.path || '') : '';
            if (pcNormalizeRelinkPath(activeProjectPath) !== pcNormalizeRelinkPath(originalProjectPath)) {
                app.openDocument(originalProjectPath, true, true, true, true);
            }
        } catch (reopenError) {
            result.failed.push('Could not return to the original project: ' + reopenError.toString());
            result.success = false;
        }
    }
}

function linkProjectCopyToCollectedMedia(projectCopyPath, tasksJson) {
    var originalProjectPath = '';
    var originalProject = null;
    var copiedProject = null;
    var result = null;

    try {
        if (!app || !app.project) {
            throw new Error('No Premiere project is currently open.');
        }

        originalProject = app.project;
        originalProjectPath = originalProject.path || '';
        if (!projectCopyPath || projectCopyPath === '') {
            throw new Error('Copied project path was not provided.');
        }

        result = {
            success: false,
            offlineCount: 0,
            linkedCount: 0,
            linkedCollectedCount: 0,
            linkedSkipLocationCount: 0,
            linkedOriginalCount: 0,
            skippedNoPathCount: 0,
            failed: []
        };

        var copiedProjectFile = new File(projectCopyPath);
        if (!copiedProjectFile.exists) {
            throw new Error('Copied BACKUP project does not exist at ' + projectCopyPath);
        }

        var tasks = pcParseJsonArray(tasksJson);
        var relinkMap = pcBuildRelinkMap(tasks);
        if (!originalProject.closeDocument) {
            throw new Error('This Premiere version cannot temporarily close the original project for safe BACKUP relinking.');
        }

        var originalSaveResult = originalProject.save();
        if (originalSaveResult !== 0 && originalSaveResult !== true && originalSaveResult !== undefined) {
            throw new Error('Premiere could not save the original project before BACKUP relinking.');
        }

        var originalCloseResult = originalProject.closeDocument(0, 0);
        if (originalCloseResult !== 0 && originalCloseResult !== true && originalCloseResult !== undefined) {
            throw new Error('Premiere could not temporarily close the original project before opening the BACKUP.');
        }

        var openResult = app.openDocument(projectCopyPath, true, true, true, true);
        var records = [];
        copiedProject = pcFindOpenProjectByPath(projectCopyPath);

        if (!copiedProject) {
            throw new Error('Premiere did not open the copied BACKUP project after the original was closed. openDocument returned: ' + openResult);
        }

        pcCollectRelinkProjectItems(copiedProject.rootItem, relinkMap, records, result);

        if (result.failed.length === 0) {
            pcOfflineRelinkRecords(records, result);
        }

        if (result.failed.length === 0 && result.offlineCount !== records.length) {
            result.failed.push('Premiere did not take every file-backed project item offline.');
        }

        if (result.failed.length === 0) {
            pcRelinkOfflineRecords(records, result);
        }

        if (result.failed.length === 0 && result.linkedCount !== records.length) {
            result.failed.push('Premiere did not relink every file-backed project item.');
        }

        if (result.failed.length === 0) {
            var saveResult = copiedProject.save();
            if (saveResult === 0 || saveResult === true || saveResult === undefined) {
                result.success = true;
            } else {
                result.failed.push('Premiere could not save the relinked BACKUP project.');
            }
        }

        if (!result.success) {
            pcCloseCopiedProjectAndRestoreOriginal(copiedProject, projectCopyPath, originalProjectPath, result);
        }

        return '{' +
            '"success":' + (result.success ? 'true' : 'false') + ',' +
            '"opened":"' + pcJsonEscape(openResult) + '",' +
            '"offlineCount":' + result.offlineCount + ',' +
            '"linkedCount":' + result.linkedCount + ',' +
            '"linkedCollectedCount":' + result.linkedCollectedCount + ',' +
            '"linkedSkipLocationCount":' + result.linkedSkipLocationCount + ',' +
            '"linkedOriginalCount":' + result.linkedOriginalCount + ',' +
            '"skippedNoPathCount":' + result.skippedNoPathCount + ',' +
            '"backupLeftOpen":' + (result.success ? 'true' : 'false') + ',' +
            '"failed":' + pcStringsJson(result.failed) +
            '}';
    } catch (e) {
        if (result) {
            result.success = false;
            result.failed.push(e.toString());
            pcCloseCopiedProjectAndRestoreOriginal(copiedProject, projectCopyPath, originalProjectPath, result);
        }

        if (originalProjectPath && pcNormalizeRelinkPath(originalProjectPath) !== pcNormalizeRelinkPath(projectCopyPath)) {
            try {
                var recoveryProjectPath = app && app.project ? (app.project.path || '') : '';
                if (pcNormalizeRelinkPath(recoveryProjectPath) !== pcNormalizeRelinkPath(originalProjectPath)) {
                    app.openDocument(originalProjectPath, true, true, true, true);
                }
            } catch (e2) {}
        }

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
