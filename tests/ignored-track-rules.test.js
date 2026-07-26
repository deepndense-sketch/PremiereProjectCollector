const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

function createTestElement() {
    const attributes = new Map();
    const classes = new Set();
    return {
        children: [],
        classList: {
            add(name) {
                classes.add(name);
            },
            remove(name) {
                classes.delete(name);
            },
            toggle(name, force) {
                if (force === true) {
                    classes.add(name);
                } else if (force === false) {
                    classes.delete(name);
                } else if (classes.has(name)) {
                    classes.delete(name);
                } else {
                    classes.add(name);
                }
            }
        },
        disabled: false,
        innerHTML: '',
        onclick: null,
        parentNode: null,
        style: {},
        textContent: '',
        appendChild(child) {
            child.parentNode = this;
            this.children.push(child);
        },
        getAttribute(name) {
            return attributes.has(name) ? attributes.get(name) : null;
        },
        setAttribute(name, value) {
            attributes.set(name, String(value));
        }
    };
}

function createTestDocument() {
    const elementIds = [
        'chooseButton',
        'collectButton',
        'compareButton',
        'completionMessage',
        'completionPrompt',
        'completionTitle',
        'currentFile',
        'errorList',
        'missingList',
        'progressFill',
        'progressText',
        'refreshProjectButton',
        'sequenceFilterHint',
        'sequenceFilters',
        'selectionSummary',
        'summaryText',
        'trackConflictContinueButton',
        'trackConflictList',
        'trackConflictPrompt',
        'trackConflictStatus',
        'updateButton'
    ];
    const elements = new Map(elementIds.map((id) => [id, createTestElement()]));

    return {
        addEventListener() {},
        createElement() {
            return createTestElement();
        },
        getElementById(id) {
            return elements.get(id) || null;
        },
        querySelector() {
            return null;
        }
    };
}

function createMemoryStorage(initialValues) {
    const values = new Map(Object.entries(initialValues || {}));
    return {
        get length() {
            return values.size;
        },
        getItem(key) {
            return values.has(key) ? values.get(key) : null;
        },
        key(index) {
            return Array.from(values.keys())[index] || null;
        },
        removeItem(key) {
            values.delete(key);
        },
        setItem(key, value) {
            values.set(key, String(value));
        }
    };
}

function loadCollectorLogic(testDocument, options) {
    const settings = options || {};
    const scriptPath = path.join(__dirname, '..', 'js', 'main.js');
    const source = fs.readFileSync(scriptPath, 'utf8');
    const context = vm.createContext({
        Buffer,
        Map,
        Set,
        URL,
        clearInterval,
        clearTimeout,
        console,
        process,
        require,
        setInterval,
        setTimeout,
        CSInterface: function CSInterface() {},
        SystemPath: { EXTENSION: 'extension' },
        document: testDocument || {
            addEventListener() {},
            getElementById() {
                return null;
            },
            querySelector() {
                return null;
            }
        },
        localStorage: settings.localStorage || createMemoryStorage(),
        window: {
            addEventListener() {},
            confirm: settings.confirm || (() => true),
            prompt: settings.prompt || (() => null)
        },
        alert: settings.alert || (() => {})
    });

    vm.runInContext(source, context, { filename: scriptPath });
    return context;
}

test('panel uses the compact destination labels and places sequence options with backup options', () => {
    const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');

    assert.match(html, />Skip Files From</);
    assert.match(html, />Copy Files To</);
    assert.doesNotMatch(html, /1\. Skip Existing Media/);
    assert.doesNotMatch(html, /Files already in this folder are skipped/);
    assert.doesNotMatch(html, /2\. Backup Destination/);
    assert.match(html, /Click to include or ignore/);
    assert.match(html, /Refresh reloads and resets every selected sequence/);
    assert.ok(html.indexOf('id="sequenceOnlyMode"') < html.indexOf('id="copyProjectFile"'));
    assert.ok(html.indexOf('id="createReducedProject"') < html.indexOf('id="copyProjectFile"'));
    assert.ok(html.indexOf('id="linkProjectAfterCollection"') < html.indexOf('id="collectButton"'));
});

test('unresolved included-and-ignored conflict never copies', () => {
    const context = loadCollectorLogic();
    const reason = vm.runInContext(`(() => {
        const task = { name: 'shared.mov', source: 'C:/Media/shared.mov', binPath: '' };
        const mediaKey = normalizeMediaKey(task.source);
        const trackRuleContext = buildTrackRuleContext(
            [task],
            new Set([mediaKey]),
            new Set([mediaKey])
        );
        const copyRuleContext = createCopyRuleContext({
            treeSelectedTaskSet: new Set([task]),
            sequenceScopedMediaSet: new Set([mediaKey]),
            trackRuleContext
        });
        return getCopySkipReason(task, copyRuleContext);
    })()`, context);

    assert.equal(reason, 'skipped because its track conflict was not resolved');
});

test('ignored Premiere project folder still wins', () => {
    const context = loadCollectorLogic();
    const reason = vm.runInContext(`(() => {
        const task = { name: 'shared.mov', source: 'C:/Media/shared.mov', binPath: 'Do Not Back Up' };
        const copyRuleContext = createCopyRuleContext({
            treeSelectedTaskSet: new Set([task]),
            trackRuleContext: buildTrackRuleContext([task], new Set(), new Set()),
            ignoredProjectFolders: ['Do Not Back Up']
        });
        return getCopySkipReason(task, copyRuleContext);
    })()`, context);

    assert.equal(reason, 'skipped because Premiere folder "Do Not Back Up" is ignored');
});

test('ignored-track matching does not widen to a different file with the same name and size', (t) => {
    const tempRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'project-collector-exact-path-'));
    const ignoredPath = path.join(tempRoot, 'ignored', 'shared.mov');
    const includedPath = path.join(tempRoot, 'included', 'shared.mov');
    fs.mkdirSync(path.dirname(ignoredPath), { recursive: true });
    fs.mkdirSync(path.dirname(includedPath), { recursive: true });
    fs.writeFileSync(ignoredPath, 'same-size-content');
    fs.writeFileSync(includedPath, 'same-size-content');
    t.after(() => fs.rmSync(tempRoot, { recursive: true, force: true }));

    const context = loadCollectorLogic();
    const result = vm.runInContext(`(() => {
        const ignoredTask = { name: 'shared.mov', source: ${JSON.stringify(ignoredPath)}, binPath: '' };
        const includedTask = { name: 'shared.mov', source: ${JSON.stringify(includedPath)}, binPath: '' };
        const ignoredMediaSet = new Set([normalizeMediaKey(ignoredTask.source)]);
        const includedMediaSet = new Set([
            normalizeMediaKey(ignoredTask.source),
            normalizeMediaKey(includedTask.source)
        ]);
        const trackRuleContext = buildTrackRuleContext(
            [ignoredTask, includedTask],
            includedMediaSet,
            ignoredMediaSet
        );
        return {
            unrelatedFileIgnored: trackRuleContext.ignoredOnlyMediaSet.has(
                normalizeMediaKey(includedTask.source)
            ),
            conflicts: trackRuleContext.conflicts.map((conflict) => conflict.source)
        };
    })()`, context);

    assert.equal(result.unrelatedFileIgnored, false);
    assert.deepEqual(Array.from(result.conflicts), [ignoredPath]);
});

test('root copy-rule precedence is deterministic across every rule source', () => {
    const context = loadCollectorLogic();
    const result = vm.runInContext(`(() => {
        function evaluate(options) {
            const task = {
                name: options.name || 'media.mov',
                source: options.source,
                binPath: options.binPath || ''
            };
            const mediaKey = normalizeMediaKey(task.source);
            const includedMediaSet = options.includedTrack ? new Set([mediaKey]) : new Set();
            const ignoredMediaSet = options.ignoredTrack ? new Set([mediaKey]) : new Set();
            const trackRuleContext = buildTrackRuleContext(
                [task],
                includedMediaSet,
                ignoredMediaSet
            );
            let conflictDecisionInfo = null;
            if (options.decision) {
                conflictDecisionInfo = buildTrackConflictDecisionInfo(trackRuleContext, [{
                    mediaKey,
                    decision: options.decision
                }]);
            }
            const copyRuleContext = createCopyRuleContext({
                treeSelectedTaskSet: options.sourceSelected ? new Set([task]) : new Set(),
                sequenceScopedMediaSet: options.sequenceScoped
                    ? (options.inSequence ? new Set([mediaKey]) : new Set())
                    : null,
                trackRuleContext,
                conflictDecisionInfo,
                includedProjectFolders: options.includedFolder ? [task.binPath] : [],
                ignoredProjectFolders: options.ignoredFolder ? [task.binPath] : []
            });
            return getCopySkipReason(task, copyRuleContext);
        }

        return {
            ignoredFolderBeforeCopyDecision: evaluate({
                source: 'C:/Media/folder-ignore.mov',
                binPath: 'Ignore',
                sourceSelected: true,
                includedTrack: true,
                ignoredTrack: true,
                decision: 'copy',
                ignoredFolder: true
            }),
            sequenceScopeBeforeCopyDecision: evaluate({
                source: 'C:/Media/out-of-sequence.mov',
                sourceSelected: true,
                sequenceScoped: true,
                inSequence: false,
                includedTrack: true,
                ignoredTrack: true,
                decision: 'copy'
            }),
            conflictCopyBeforeSourceList: evaluate({
                source: 'C:/Media/conflict-copy.mov',
                sourceSelected: false,
                includedTrack: true,
                ignoredTrack: true,
                decision: 'copy'
            }),
            conflictSkip: evaluate({
                source: 'C:/Media/conflict-skip.mov',
                sourceSelected: true,
                includedTrack: true,
                ignoredTrack: true,
                decision: 'skip'
            }),
            unresolvedConflict: evaluate({
                source: 'C:/Media/conflict-unresolved.mov',
                sourceSelected: true,
                includedTrack: true,
                ignoredTrack: true
            }),
            ignoredTrackBeforeIncludedFolder: evaluate({
                source: 'C:/Media/ignored-track.mov',
                binPath: 'Force',
                sourceSelected: true,
                ignoredTrack: true,
                includedFolder: true
            }),
            includedFolderBeforeSourceList: evaluate({
                source: 'C:/Media/included-folder.mov',
                binPath: 'Force',
                sourceSelected: false,
                includedFolder: true
            }),
            sourceListSkip: evaluate({
                source: 'C:/Media/source-list-skip.mov',
                sourceSelected: false
            }),
            normalCopy: evaluate({
                source: 'C:/Media/normal-copy.mov',
                sourceSelected: true
            })
        };
    })()`, context);

    assert.equal(result.ignoredFolderBeforeCopyDecision, 'skipped because Premiere folder "Ignore" is ignored');
    assert.equal(result.sequenceScopeBeforeCopyDecision, 'skipped because it is not used by the chosen sequences');
    assert.equal(result.conflictCopyBeforeSourceList, '');
    assert.equal(result.conflictSkip, 'skipped by your track conflict decision');
    assert.equal(result.unresolvedConflict, 'skipped because its track conflict was not resolved');
    assert.equal(result.ignoredTrackBeforeIncludedFolder, 'skipped by ignored track selection');
    assert.equal(result.includedFolderBeforeSourceList, '');
    assert.equal(result.sourceListSkip, 'skipped by Source File List selection');
    assert.equal(result.normalCopy, '');
});

test('validated conflict decisions reject missing, duplicate, and out-of-scope media', () => {
    const context = loadCollectorLogic();
    const result = vm.runInContext(`(() => {
        const conflictTask = { name: 'conflict.mov', source: 'C:/Media/conflict.mov', binPath: '' };
        const otherTask = { name: 'other.mov', source: 'C:/Media/other.mov', binPath: '' };
        const conflictKey = normalizeMediaKey(conflictTask.source);
        const otherKey = normalizeMediaKey(otherTask.source);
        const trackRuleContext = buildTrackRuleContext(
            [conflictTask, otherTask],
            new Set([conflictKey, otherKey]),
            new Set([conflictKey])
        );
        return {
            valid: buildTrackConflictDecisionInfo(trackRuleContext, [{
                mediaKey: conflictKey,
                decision: 'copy'
            }]).valid,
            missing: buildTrackConflictDecisionInfo(trackRuleContext, []).valid,
            duplicate: buildTrackConflictDecisionInfo(trackRuleContext, [{
                mediaKey: conflictKey,
                decision: 'copy'
            }, {
                mediaKey: conflictKey,
                decision: 'skip'
            }]).valid,
            outside: buildTrackConflictDecisionInfo(trackRuleContext, [{
                mediaKey: otherKey,
                decision: 'skip'
            }]).valid
        };
    })()`, context);

    assert.equal(result.valid, true);
    assert.equal(result.missing, false);
    assert.equal(result.duplicate, false);
    assert.equal(result.outside, false);
});

test('collection automatically skips shared ignored media and copies allowed media', async (t) => {
    const tempRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'project-collector-ignored-'));
    const sourceDir = path.join(tempRoot, 'source');
    const destinationDir = path.join(tempRoot, 'backup');
    const ignoredSourcePath = path.join(sourceDir, 'shared.mov');
    const allowedSourcePath = path.join(sourceDir, 'allowed.mov');
    const collectedRoot = path.join(destinationDir, 'Ignored_Track_Project', 'CollectedMedias');
    fs.mkdirSync(sourceDir, { recursive: true });
    fs.writeFileSync(ignoredSourcePath, 'ignored-shared-media');
    fs.writeFileSync(allowedSourcePath, 'allowed-media');
    t.after(() => fs.rmSync(tempRoot, { recursive: true, force: true }));

    const context = loadCollectorLogic(createTestDocument());
    vm.runInContext(`
        destination = ${JSON.stringify(destinationDir)};
        compareLocation = null;
        copyProjectFile = false;
        linkProjectAfterCollection = false;
        sequenceOnlyMode = false;
        createReducedProject = false;
        latestPlan = {
            projectName: 'Ignored_Track_Project',
            missingMedia: [],
            tasks: [{
                name: 'shared.mov',
                source: ${JSON.stringify(ignoredSourcePath)},
                destination: 'Media/shared.mov',
                binPath: '',
                relativePath: 'Media/shared.mov'
            }, {
                name: 'allowed.mov',
                source: ${JSON.stringify(allowedSourcePath)},
                destination: 'Media/allowed.mov',
                binPath: '',
                relativePath: 'Media/allowed.mov'
            }]
        };
        sourceTree = [];
        buildCopyReadyContext = async function buildCopyReadyContextForTest() {
            const treeSelectedTaskSet = new Set(latestPlan.tasks);
            const ignoredMediaSet = new Set([normalizeMediaKey(latestPlan.tasks[0].source)]);
            const trackRuleContext = buildTrackRuleContext(
                latestPlan.tasks,
                new Set(),
                ignoredMediaSet
            );
            return {
                ok: true,
                copyRuleContext: createCopyRuleContext({
                    treeSelectedTaskSet,
                    trackRuleContext
                }),
                copyWarnings: [],
                sequenceScopeInfo: null,
                trackConflicts: trackRuleContext.conflicts
            };
        };
    `, context);

    await vm.runInContext('collect()', context);

    assert.equal(fs.existsSync(path.join(collectedRoot, 'shared.mov')), false);
    assert.equal(fs.existsSync(path.join(collectedRoot, 'allowed.mov')), true);
    assert.equal(fs.readFileSync(path.join(collectedRoot, 'allowed.mov'), 'utf8'), 'allowed-media');
    assert.equal(vm.runInContext('isCopying', context), false);
});

test('individual row clicks apply mixed Copy and Do not copy choices to the real copy loop', async (t) => {
    const tempRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'project-collector-individual-'));
    const sourceDir = path.join(tempRoot, 'source');
    const destinationDir = path.join(tempRoot, 'backup');
    const copySourcePath = path.join(sourceDir, 'copy-this.mov');
    const skipSourcePath = path.join(sourceDir, 'skip-this.mov');
    const allowedSourcePath = path.join(sourceDir, 'allowed.mov');
    const collectedRoot = path.join(destinationDir, 'Individual_Choice_Project', 'CollectedMedias');
    fs.mkdirSync(sourceDir, { recursive: true });
    fs.writeFileSync(copySourcePath, 'copy-choice-media');
    fs.writeFileSync(skipSourcePath, 'skip-choice-media');
    fs.writeFileSync(allowedSourcePath, 'allowed-media');
    t.after(() => fs.rmSync(tempRoot, { recursive: true, force: true }));

    const context = loadCollectorLogic(createTestDocument());
    vm.runInContext(`
        destination = ${JSON.stringify(destinationDir)};
        compareLocation = null;
        copyProjectFile = false;
        linkProjectAfterCollection = false;
        sequenceOnlyMode = false;
        createReducedProject = false;
        latestPlan = {
            projectName: 'Individual_Choice_Project',
            missingMedia: [],
            tasks: [{
                name: 'copy-this.mov',
                source: ${JSON.stringify(copySourcePath)},
                destination: 'Media/copy-this.mov',
                binPath: '',
                relativePath: 'Media/copy-this.mov'
            }, {
                name: 'skip-this.mov',
                source: ${JSON.stringify(skipSourcePath)},
                destination: 'Media/skip-this.mov',
                binPath: '',
                relativePath: 'Media/skip-this.mov'
            }, {
                name: 'allowed.mov',
                source: ${JSON.stringify(allowedSourcePath)},
                destination: 'Media/allowed.mov',
                binPath: '',
                relativePath: 'Media/allowed.mov'
            }]
        };
        sourceTree = [];
        buildCopyReadyContext = async function buildCopyReadyContextForIndividualChoiceTest() {
            const treeSelectedTaskSet = new Set(latestPlan.tasks);
            const ignoredMediaSet = new Set([
                normalizeMediaKey(latestPlan.tasks[0].source),
                normalizeMediaKey(latestPlan.tasks[1].source)
            ]);
            const includedMediaSet = new Set([
                normalizeMediaKey(latestPlan.tasks[0].source),
                normalizeMediaKey(latestPlan.tasks[1].source)
            ]);
            const trackRuleContext = buildTrackRuleContext(
                latestPlan.tasks,
                includedMediaSet,
                ignoredMediaSet
            );
            return {
                ok: true,
                copyRuleContext: createCopyRuleContext({
                    treeSelectedTaskSet,
                    trackRuleContext
                }),
                copyWarnings: [],
                sequenceScopeInfo: null,
                trackConflicts: trackRuleContext.conflicts
            };
        };
    `, context);

    const collectionPromise = vm.runInContext('collect()', context);
    await new Promise((resolve) => setImmediate(resolve));

    const stateBeforeContinue = vm.runInContext(`(() => {
        const list = document.getElementById('trackConflictList');
        list.onclick({ target: trackConflictPromptState.rows[0].copyButton });
        list.onclick({ target: trackConflictPromptState.rows[1].skipButton });
        return {
            decisions: trackConflictPromptState.decisions.slice(),
            continueDisabled: document.getElementById('trackConflictContinueButton').disabled
        };
    })()`, context);

    assert.deepEqual(Array.from(stateBeforeContinue.decisions), ['copy', 'skip']);
    assert.equal(stateBeforeContinue.continueDisabled, false);

    vm.runInContext('finishTrackConflictPrompt(true)', context);
    await collectionPromise;

    assert.equal(fs.existsSync(path.join(collectedRoot, 'copy-this.mov')), true);
    assert.equal(fs.readFileSync(path.join(collectedRoot, 'copy-this.mov'), 'utf8'), 'copy-choice-media');
    assert.equal(fs.existsSync(path.join(collectedRoot, 'skip-this.mov')), false);
    assert.equal(fs.existsSync(path.join(collectedRoot, 'allowed.mov')), true);
    assert.equal(vm.runInContext('isCopying', context), false);
});

test('Do not copy all skips only listed conflicts and still copies unrelated included media', async (t) => {
    const tempRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'project-collector-skip-all-'));
    const sourceDir = path.join(tempRoot, 'source');
    const destinationDir = path.join(tempRoot, 'backup');
    const fileNames = ['conflict-one.mov', 'conflict-two.mov', 'included-one.mov', 'included-two.mov'];
    const sourcePaths = fileNames.map((fileName) => path.join(sourceDir, fileName));
    const collectedRoot = path.join(destinationDir, 'Skip_All_Scope_Project', 'CollectedMedias');
    fs.mkdirSync(sourceDir, { recursive: true });
    sourcePaths.forEach((filePath, index) => {
        fs.writeFileSync(filePath, `media-${index + 1}`);
    });
    t.after(() => fs.rmSync(tempRoot, { recursive: true, force: true }));

    const context = loadCollectorLogic(createTestDocument());
    vm.runInContext(`
        destination = ${JSON.stringify(destinationDir)};
        compareLocation = null;
        copyProjectFile = false;
        linkProjectAfterCollection = false;
        sequenceOnlyMode = false;
        createReducedProject = false;
        latestPlan = {
            projectName: 'Skip_All_Scope_Project',
            missingMedia: [],
            tasks: ${JSON.stringify(sourcePaths.map((source, index) => ({
                name: fileNames[index],
                source,
                destination: `Media/${fileNames[index]}`,
                binPath: '',
                relativePath: `Media/${fileNames[index]}`
            })))}
        };
        sourceTree = [];
        buildCopyReadyContext = async function buildCopyReadyContextForSkipAllTest() {
            const treeSelectedTaskSet = new Set(latestPlan.tasks);
            const ignoredMediaSet = new Set([
                normalizeMediaKey(latestPlan.tasks[0].source),
                normalizeMediaKey(latestPlan.tasks[1].source)
            ]);
            const includedMediaSet = new Set(latestPlan.tasks.map((task) => normalizeMediaKey(task.source)));
            const trackRuleContext = buildTrackRuleContext(
                latestPlan.tasks,
                includedMediaSet,
                ignoredMediaSet
            );
            return {
                ok: true,
                copyRuleContext: createCopyRuleContext({
                    treeSelectedTaskSet,
                    trackRuleContext
                }),
                copyWarnings: [],
                sequenceScopeInfo: null,
                trackConflicts: trackRuleContext.conflicts
            };
        };
    `, context);

    const collectionPromise = vm.runInContext('collect()', context);
    await new Promise((resolve) => setImmediate(resolve));
    assert.equal(vm.runInContext('trackConflictPromptState.conflicts.length', context), 2);

    vm.runInContext(`
        setAllTrackConflictChoices('skip');
        finishTrackConflictPrompt(true);
    `, context);
    await collectionPromise;

    assert.equal(fs.existsSync(path.join(collectedRoot, 'conflict-one.mov')), false);
    assert.equal(fs.existsSync(path.join(collectedRoot, 'conflict-two.mov')), false);
    assert.equal(fs.existsSync(path.join(collectedRoot, 'included-one.mov')), true);
    assert.equal(fs.existsSync(path.join(collectedRoot, 'included-two.mov')), true);
    assert.equal(vm.runInContext('isCopying', context), false);
});

test('Copy all copies listed conflicts without selecting unrelated media', async (t) => {
    const tempRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'project-collector-copy-all-'));
    const sourceDir = path.join(tempRoot, 'source');
    const destinationDir = path.join(tempRoot, 'backup');
    const fileNames = ['conflict-one.mov', 'conflict-two.mov', 'unselected.mov'];
    const sourcePaths = fileNames.map((fileName) => path.join(sourceDir, fileName));
    const collectedRoot = path.join(destinationDir, 'Copy_All_Scope_Project', 'CollectedMedias');
    fs.mkdirSync(sourceDir, { recursive: true });
    sourcePaths.forEach((filePath, index) => {
        fs.writeFileSync(filePath, `copy-all-media-${index + 1}`);
    });
    t.after(() => fs.rmSync(tempRoot, { recursive: true, force: true }));

    const context = loadCollectorLogic(createTestDocument());
    vm.runInContext(`
        destination = ${JSON.stringify(destinationDir)};
        compareLocation = null;
        copyProjectFile = false;
        linkProjectAfterCollection = false;
        sequenceOnlyMode = false;
        createReducedProject = false;
        latestPlan = {
            projectName: 'Copy_All_Scope_Project',
            missingMedia: [],
            tasks: ${JSON.stringify(sourcePaths.map((source, index) => ({
                name: fileNames[index],
                source,
                destination: `Media/${fileNames[index]}`,
                binPath: '',
                relativePath: `Media/${fileNames[index]}`
            })))}
        };
        sourceTree = [];
        buildCopyReadyContext = async function buildCopyReadyContextForCopyAllTest() {
            const treeSelectedTaskSet = new Set();
            const ignoredMediaSet = new Set([
                normalizeMediaKey(latestPlan.tasks[0].source),
                normalizeMediaKey(latestPlan.tasks[1].source)
            ]);
            const includedMediaSet = new Set(latestPlan.tasks.map((task) => normalizeMediaKey(task.source)));
            const trackRuleContext = buildTrackRuleContext(
                latestPlan.tasks,
                includedMediaSet,
                ignoredMediaSet
            );
            return {
                ok: true,
                copyRuleContext: createCopyRuleContext({
                    treeSelectedTaskSet,
                    trackRuleContext
                }),
                copyWarnings: [],
                sequenceScopeInfo: null,
                trackConflicts: trackRuleContext.conflicts
            };
        };
    `, context);

    const collectionPromise = vm.runInContext('collect()', context);
    await new Promise((resolve) => setImmediate(resolve));
    assert.equal(vm.runInContext('trackConflictPromptState.conflicts.length', context), 2);

    vm.runInContext(`
        setAllTrackConflictChoices('copy');
        finishTrackConflictPrompt(true);
    `, context);
    await collectionPromise;

    assert.equal(fs.existsSync(path.join(collectedRoot, 'conflict-one.mov')), true);
    assert.equal(fs.existsSync(path.join(collectedRoot, 'conflict-two.mov')), true);
    assert.equal(fs.existsSync(path.join(collectedRoot, 'unselected.mov')), false);
    assert.equal(vm.runInContext('isCopying', context), false);
});

test('per-sequence track reset clears only the requested sequence', () => {
    const context = loadCollectorLogic();
    const result = JSON.parse(vm.runInContext(`JSON.stringify((() => {
        const filters = [
            {
                sequenceID: 'sequence-a',
                sequenceName: 'Sequence A',
                ignoredVideoTracks: [1, 3],
                ignoredAudioTracks: [2]
            },
            {
                sequenceID: 'sequence-b',
                sequenceName: 'Sequence B',
                ignoredVideoTracks: [4],
                ignoredAudioTracks: [1, 5]
            }
        ];
        const resetCount = clearTrackChoices(filters, 'sequence-b');
        return { filters, resetCount };
    })())`, context));

    assert.equal(result.resetCount, 1);
    assert.deepEqual(result.filters[0].ignoredVideoTracks, [1, 3]);
    assert.deepEqual(result.filters[0].ignoredAudioTracks, [2]);
    assert.deepEqual(result.filters[1].ignoredVideoTracks, []);
    assert.deepEqual(result.filters[1].ignoredAudioTracks, []);
});

test('global sequence refresh updates and resets every selected sequence without changing its scope', () => {
    const context = loadCollectorLogic();
    const result = JSON.parse(vm.runInContext(`JSON.stringify((() => {
        const filters = [
            createSequenceFilter('sequence-a', 'Sequence A', [
                { trackNumber: 1, clipCount: 1, hasItems: true }
            ], [], true),
            createSequenceFilter('sequence-b', 'Sequence B', [], [
                { trackNumber: 1, clipCount: 1, hasItems: true }
            ], false),
            createSequenceFilter('sequence-missing', 'Deleted Sequence', [
                { trackNumber: 9, clipCount: 9, hasItems: true }
            ], [], false)
        ];
        filters[0].ignoredVideoTracks = [1];
        filters[1].ignoredAudioTracks = [1];
        filters[2].ignoredVideoTracks = [9];

        const refreshResult = applySequenceTrackUsagePlan(filters, {
            sequences: [
                {
                    sequenceID: 'sequence-b',
                    sequenceName: 'Sequence B',
                    videoTrackUsage: [],
                    audioTrackUsage: [
                        { trackNumber: 2, clipCount: 7, hasItems: true }
                    ]
                },
                {
                    sequenceID: 'sequence-a',
                    sequenceName: 'Renamed Sequence A',
                    videoTrackUsage: [
                        { trackNumber: 3, clipCount: 5, hasItems: true }
                    ],
                    audioTrackUsage: []
                }
            ]
        }, true);

        return { filters, refreshResult };
    })())`, context));

    assert.equal(result.refreshResult.refreshedCount, 2);
    assert.equal(result.refreshResult.resetCount, 3);
    assert.deepEqual(result.refreshResult.missingSequences, ['Deleted Sequence']);
    assert.equal(result.filters.length, 3);
    assert.equal(result.filters[0].sequenceName, 'Renamed Sequence A');
    assert.equal(result.filters[0].videoTrackUsage[0].trackNumber, 3);
    assert.equal(result.filters[0].videoTrackUsage[0].clipCount, 5);
    assert.equal(result.filters[1].audioTrackUsage[0].trackNumber, 2);
    assert.equal(result.filters[1].audioTrackUsage[0].clipCount, 7);
    assert.equal(result.filters[2].videoTrackUsage[0].trackNumber, 9);
    result.filters.forEach((filter) => {
        assert.deepEqual(filter.ignoredVideoTracks, []);
        assert.deepEqual(filter.ignoredAudioTracks, []);
    });
});

test('Premiere host refresh plan reads every requested sequence, not only the active sequence', () => {
    const hostSource = fs.readFileSync(path.join(__dirname, '..', 'jsx', 'collector.jsx'), 'utf8');
    const hostContext = vm.createContext({
        JSON,
        app: {
            project: {
                activeSequence: null,
                sequences: null
            }
        }
    });

    vm.runInContext(`
        function createTestTracks(clipCounts, mediaPrefix) {
            const tracks = [];
            clipCounts.forEach((clipCount, trackIndex) => {
                const clips = [];
                for (let clipIndex = 0; clipIndex < clipCount; clipIndex += 1) {
                    const mediaPath = mediaPrefix + '-track-' + (trackIndex + 1) + '-clip-' + (clipIndex + 1) + '.mov';
                    clips.push({
                        projectItem: {
                            getMediaPath: function getMediaPath() {
                                return mediaPath;
                            }
                        }
                    });
                }
                clips.numItems = clips.length;
                tracks.push({ clips });
            });
            tracks.numTracks = tracks.length;
            return tracks;
        }

        const sequenceA = {
            sequenceID: 'sequence-a',
            name: 'Sequence A',
            videoTracks: createTestTracks([2, 0], 'A-video'),
            audioTracks: createTestTracks([1], 'A-audio')
        };
        const sequenceB = {
            sequenceID: 'sequence-b',
            name: 'Sequence B',
            videoTracks: createTestTracks([3], 'B-video'),
            audioTracks: createTestTracks([0, 4], 'B-audio')
        };
        const sequences = [sequenceA, sequenceB];
        sequences.numSequences = sequences.length;
        app.project.sequences = sequences;
        app.project.activeSequence = sequenceA;
    `, hostContext);
    vm.runInContext(hostSource, hostContext, { filename: 'collector.jsx' });

    const raw = vm.runInContext(`getSequenceTrackUsagePlan(${JSON.stringify(JSON.stringify([
        { sequenceID: 'sequence-b', sequenceName: 'Old B Name' },
        { sequenceID: '', sequenceName: 'Sequence A' },
        { sequenceID: 'missing-sequence', sequenceName: 'Missing Sequence' }
    ]))})`, hostContext);
    const result = JSON.parse(raw);

    assert.equal(result.sequences.length, 2);
    assert.equal(result.sequences[0].sequenceID, 'sequence-b');
    assert.equal(result.sequences[0].videoTrackUsage[0].clipCount, 3);
    assert.equal(result.sequences[0].audioTrackUsage[1].clipCount, 4);
    assert.equal(result.sequences[1].sequenceID, 'sequence-a');
    assert.equal(result.sequences[1].videoTrackUsage[0].clipCount, 2);
    assert.deepEqual(result.missingSequences, ['Missing Sequence']);
});

test('Refresh Project invokes the all-sequence refresh in reset mode', async () => {
    const testDocument = createTestDocument();
    const context = loadCollectorLogic(testDocument);

    const result = await vm.runInContext(`(async () => {
        let loadCount = 0;
        let allSequenceRefreshCount = 0;
        let receivedResetMode = null;
        loadProjectPlan = async function loadProjectPlanForRefreshTest() {
            loadCount += 1;
            return true;
        };
        refreshAllSelectedSequenceTracks = async function refreshAllForRefreshTest(resetChoices) {
            allSequenceRefreshCount += 1;
            receivedResetMode = resetChoices;
            return {
                refreshedCount: 2,
                missingSequences: [],
                resetCount: 2
            };
        };

        await refreshProject();
        return {
            loadCount,
            allSequenceRefreshCount,
            receivedResetMode,
            summary: document.getElementById('summaryText').textContent
        };
    })()`, context);

    assert.equal(result.loadCount, 1);
    assert.equal(result.allSequenceRefreshCount, 1);
    assert.equal(result.receivedResetMode, true);
    assert.match(result.summary, /all 2 selected sequences/i);
});

test('track presets persist per user and generate the next available default name', () => {
    const storage = createMemoryStorage();
    const context = loadCollectorLogic(null, { localStorage: storage });
    const result = JSON.parse(vm.runInContext(`JSON.stringify((() => {
        loadTrackPresets();
        const filter = createSequenceFilter(
            'sequence-a',
            'Sequence A',
            [
                { trackNumber: 1, clipCount: 1, hasItems: true },
                { trackNumber: 2, clipCount: 1, hasItems: true }
            ],
            [{ trackNumber: 1, clipCount: 1, hasItems: true }],
            true
        );
        filter.ignoredVideoTracks = [2];
        filter.ignoredAudioTracks = [1];

        const firstDefaultName = getNextTrackPresetName(trackPresets);
        const firstPreset = upsertTrackPresetFromFilter(filter, firstDefaultName);
        const secondDefaultName = getNextTrackPresetName(trackPresets);
        saveTrackPresets();

        trackPresets = [];
        trackPresetsLoaded = false;
        loadTrackPresets();
        const reloadedBeforeUpdate = JSON.parse(JSON.stringify(trackPresets));

        filter.ignoredVideoTracks = [1];
        filter.ignoredAudioTracks = [];
        const updatedPreset = upsertTrackPresetFromFilter(filter, 'project copy preset 1');

        return {
            firstDefaultName,
            secondDefaultName,
            firstPresetId: firstPreset.id,
            reloaded: reloadedBeforeUpdate,
            updatedPreset,
            finalCount: trackPresets.length
        };
    })())`, context));

    assert.equal(result.firstDefaultName, 'Project Copy Preset 1');
    assert.equal(result.secondDefaultName, 'Project Copy Preset 2');
    assert.equal(result.reloaded.length, 1);
    assert.equal(result.reloaded[0].id, result.firstPresetId);
    assert.deepEqual(result.reloaded[0].ignoredVideoTracks, [2]);
    assert.deepEqual(result.reloaded[0].ignoredAudioTracks, [1]);
    assert.equal(result.finalCount, 1);
    assert.equal(result.updatedPreset.id, result.firstPresetId);
    assert.deepEqual(result.updatedPreset.ignoredVideoTracks, [1]);
    assert.deepEqual(result.updatedPreset.ignoredAudioTracks, []);
});

test('manual cleanup preserves only track presets, destination, and unrelated application data', () => {
    const storage = createMemoryStorage({
        'projectcollector.trackPresets': '[{"id":"preset-1","name":"Keep Me"}]',
        'projectcollector.destination': 'D:/Backups',
        'projectcollector.sequenceOnlyMode': '1',
        'projectcollector.copyProjectFile': '1',
        'projectcollector.sequenceFilters:c:/project.prproj': '[{"sequenceName":"Old"}]',
        'another.extension.setting': 'keep'
    });
    const context = loadCollectorLogic(null, { localStorage: storage });
    const cleanupResult = JSON.parse(vm.runInContext(`JSON.stringify((() => {
        const deleted = deletePluginRelatedData();
        return {
            deleted,
            remainingProjectKeys: Array.from(
                { length: localStorage.length },
                (_, index) => localStorage.key(index)
            ).filter((key) => key && key.indexOf('projectcollector.') === 0).sort()
        };
    })())`, context));

    assert.equal(cleanupResult.deleted, true);
    assert.deepEqual(cleanupResult.remainingProjectKeys, [
        'projectcollector.destination',
        'projectcollector.trackPresets'
    ]);
    assert.equal(storage.getItem('projectcollector.trackPresets'), '[{"id":"preset-1","name":"Keep Me"}]');
    assert.equal(storage.getItem('projectcollector.destination'), 'D:/Backups');
    assert.equal(storage.getItem('another.extension.setting'), 'keep');
});

test('different sequences can apply different presets from the same shared list', () => {
    const context = loadCollectorLogic(createTestDocument());
    const result = JSON.parse(vm.runInContext(`JSON.stringify((() => {
        trackPresetsLoaded = true;
        trackPresets = [
            {
                id: 'dialogue-preset',
                name: 'Dialogue',
                ignoredVideoTracks: [2, 9],
                ignoredAudioTracks: [1]
            },
            {
                id: 'music-preset',
                name: 'Music',
                ignoredVideoTracks: [1],
                ignoredAudioTracks: [2, 8]
            }
        ];
        const usage = {
            video: [
                { trackNumber: 1, clipCount: 1, hasItems: true },
                { trackNumber: 2, clipCount: 1, hasItems: true }
            ],
            audio: [
                { trackNumber: 1, clipCount: 1, hasItems: true },
                { trackNumber: 2, clipCount: 1, hasItems: true }
            ]
        };
        selectedSequenceFilters = [
            createSequenceFilter('sequence-a', 'Sequence A', usage.video, usage.audio, true),
            createSequenceFilter('sequence-b', 'Sequence B', usage.video, usage.audio, false)
        ];
        renderSequenceFilters = function renderSequenceFiltersForPresetTest() {};

        applyTrackPresetToSequence('sequence-a', 'dialogue-preset');
        applyTrackPresetToSequence('sequence-b', 'music-preset');

        const firstSection = renderSequencePresetSection(selectedSequenceFilters[0]);
        const secondSection = renderSequencePresetSection(selectedSequenceFilters[1]);
        const firstSelect = firstSection.children.find((control) => String(control.className || '').indexOf('sequence-preset-select') !== -1);
        const secondSelect = secondSection.children.find((control) => String(control.className || '').indexOf('sequence-preset-select') !== -1);
        return {
            filters: selectedSequenceFilters,
            sharedFirstList: firstSelect.children.map((option) => option.textContent),
            sharedSecondList: secondSelect.children.map((option) => option.textContent)
        };
    })())`, context));

    assert.equal(result.filters[0].selectedPresetId, 'dialogue-preset');
    assert.deepEqual(result.filters[0].ignoredVideoTracks, [2]);
    assert.deepEqual(result.filters[0].ignoredAudioTracks, [1]);
    assert.equal(result.filters[1].selectedPresetId, 'music-preset');
    assert.deepEqual(result.filters[1].ignoredVideoTracks, [1]);
    assert.deepEqual(result.filters[1].ignoredAudioTracks, [2]);
    assert.deepEqual(result.sharedFirstList, ['Custom tracks', 'Dialogue', 'Music']);
    assert.deepEqual(result.sharedSecondList, result.sharedFirstList);
});

test('Save preset prompts with an automatic name and immediately selects the saved preset', () => {
    const storage = createMemoryStorage();
    let receivedDefaultName = '';
    const context = loadCollectorLogic(createTestDocument(), {
        localStorage: storage,
        prompt(message, defaultName) {
            receivedDefaultName = defaultName;
            return 'Interview Tracks';
        }
    });

    const result = JSON.parse(vm.runInContext(`JSON.stringify((() => {
        trackPresetsLoaded = true;
        trackPresets = [];
        selectedSequenceFilters = [
            createSequenceFilter(
                'sequence-a',
                'Sequence A',
                [
                    { trackNumber: 1, clipCount: 1, hasItems: true },
                    { trackNumber: 2, clipCount: 1, hasItems: true }
                ],
                [{ trackNumber: 1, clipCount: 1, hasItems: true }],
                true
            )
        ];
        selectedSequenceFilters[0].ignoredVideoTracks = [2];
        selectedSequenceFilters[0].ignoredAudioTracks = [1];
        renderSequenceFilters = function renderSequenceFiltersForSavePresetTest() {};
        updateSelectionSummary = function updateSelectionSummaryForSavePresetTest() {};

        saveTrackPresetForSequence('sequence-a');
        return {
            presets: trackPresets,
            filter: selectedSequenceFilters[0],
            storedPresets: JSON.parse(localStorage.getItem(TRACK_PRESETS_STORAGE_KEY))
        };
    })())`, context));

    assert.equal(receivedDefaultName, 'Project Copy Preset 1');
    assert.equal(result.presets.length, 1);
    assert.equal(result.presets[0].name, 'Interview Tracks');
    assert.deepEqual(result.presets[0].ignoredVideoTracks, [2]);
    assert.deepEqual(result.presets[0].ignoredAudioTracks, [1]);
    assert.equal(result.filter.selectedPresetId, result.presets[0].id);
    assert.deepEqual(result.storedPresets, result.presets);

    const reloadContext = loadCollectorLogic(null, { localStorage: storage });
    const reloaded = JSON.parse(vm.runInContext(`JSON.stringify((() => {
        latestPlan = {
            activeSequenceID: 'sequence-a',
            activeSequenceName: 'Sequence A',
            videoTrackUsage: [
                { trackNumber: 1, clipCount: 1, hasItems: true },
                { trackNumber: 2, clipCount: 1, hasItems: true }
            ],
            audioTrackUsage: [{ trackNumber: 1, clipCount: 1, hasItems: true }]
        };
        renderSequenceFilters = function renderSequenceFiltersForReloadPresetTest() {};
        loadSequenceFilters();
        return {
            presets: trackPresets,
            filter: selectedSequenceFilters[0]
        };
    })())`, reloadContext));

    assert.equal(reloaded.presets.length, 1);
    assert.equal(reloaded.filter.selectedPresetId, result.presets[0].id);
    assert.deepEqual(reloaded.filter.ignoredVideoTracks, [2]);
    assert.deepEqual(reloaded.filter.ignoredAudioTracks, [1]);
});

test('manual track edits and resets detach only that sequence from its preset', () => {
    const context = loadCollectorLogic();
    const result = JSON.parse(vm.runInContext(`JSON.stringify((() => {
        trackPresetsLoaded = true;
        trackPresets = [{
            id: 'shared-preset',
            name: 'Shared',
            ignoredVideoTracks: [2],
            ignoredAudioTracks: []
        }];
        selectedSequenceFilters = [
            createSequenceFilter(
                'sequence-a',
                'Sequence A',
                [
                    { trackNumber: 1, clipCount: 1, hasItems: true },
                    { trackNumber: 2, clipCount: 1, hasItems: true }
                ],
                [],
                true
            ),
            createSequenceFilter(
                'sequence-b',
                'Sequence B',
                [
                    { trackNumber: 1, clipCount: 1, hasItems: true },
                    { trackNumber: 2, clipCount: 1, hasItems: true }
                ],
                [],
                false
            )
        ];
        selectedSequenceFilters.forEach((filter) => applyTrackPresetToFilter(filter, trackPresets[0]));
        renderSequenceFilters = function renderSequenceFiltersForDetachTest() {};

        toggleIgnoredTrack('sequence-a', 'video', 1);
        const afterManualEdit = selectedSequenceFilters.map((filter) => ({
            selectedPresetId: filter.selectedPresetId,
            ignoredVideoTracks: filter.ignoredVideoTracks.slice()
        }));
        resetTrackSelection('sequence-b');

        return {
            afterManualEdit,
            finalFilters: selectedSequenceFilters
        };
    })())`, context));

    assert.equal(result.afterManualEdit[0].selectedPresetId, '');
    assert.deepEqual(result.afterManualEdit[0].ignoredVideoTracks, [1, 2]);
    assert.equal(result.afterManualEdit[1].selectedPresetId, 'shared-preset');
    assert.deepEqual(result.afterManualEdit[1].ignoredVideoTracks, [2]);
    assert.equal(result.finalFilters[1].selectedPresetId, '');
    assert.deepEqual(result.finalFilters[1].ignoredVideoTracks, []);
    assert.deepEqual(result.finalFilters[0].ignoredVideoTracks, [1, 2]);
});

test('every sequence is removable and an intentionally empty list stays saved', () => {
    const storage = createMemoryStorage();
    const firstContext = loadCollectorLogic(null, { localStorage: storage });
    const firstResult = JSON.parse(vm.runInContext(`JSON.stringify((() => {
        latestPlan = {
            activeSequenceID: 'sequence-a',
            activeSequenceName: 'Sequence A',
            videoTrackUsage: [{ trackNumber: 1, clipCount: 1, hasItems: true }],
            audioTrackUsage: []
        };
        renderSequenceFilters = function renderSequenceFiltersForRemovalTest() {};
        loadSequenceFilters();
        const initialCount = selectedSequenceFilters.length;
        removeSequenceFilter('sequence-a');
        loadSequenceFilters();
        return {
            initialCount,
            remainingCount: selectedSequenceFilters.length,
            storedFilters: JSON.parse(localStorage.getItem(getSequenceFiltersStorageKey()))
        };
    })())`, firstContext));

    assert.equal(firstResult.initialCount, 1);
    assert.equal(firstResult.remainingCount, 0);
    assert.deepEqual(firstResult.storedFilters, []);

    const reloadContext = loadCollectorLogic(null, { localStorage: storage });
    const reloadCount = vm.runInContext(`(() => {
        latestPlan = {
            activeSequenceID: 'sequence-a',
            activeSequenceName: 'Sequence A',
            videoTrackUsage: [{ trackNumber: 1, clipCount: 1, hasItems: true }],
            audioTrackUsage: []
        };
        renderSequenceFilters = function renderSequenceFiltersForEmptyReloadTest() {};
        loadSequenceFilters();
        return selectedSequenceFilters.length;
    })()`, reloadContext);
    assert.equal(reloadCount, 0);

    const testDocument = createTestDocument();
    const renderContext = loadCollectorLogic(testDocument);
    const actionLabels = JSON.parse(vm.runInContext(`JSON.stringify((() => {
        selectedSequenceFilters = [
            createSequenceFilter(
                'sequence-a',
                'Sequence A',
                [{ trackNumber: 1, clipCount: 1, hasItems: true }],
                [],
                true
            )
        ];
        updateSelectionSummary = function updateSelectionSummaryForRemoveButtonTest() {};
        renderSequenceFilters();
        const card = document.getElementById('sequenceFilters').children[0];
        const headerActions = card.children[0].children[1];
        const presetSection = card.children[1];
        return {
            header: headerActions.children.map((button) => button.textContent),
            preset: presetSection.children.map((control) => control.textContent)
        };
    })())`, renderContext));
    assert.equal(actionLabels.header.length, 1);
    assert.equal(actionLabels.header[0], '×');
    assert.deepEqual(actionLabels.preset.slice(1), ['Save preset', 'Delete', 'Reset tracks']);
});

test('sequence selections are saved separately for each Premiere project', () => {
    const storage = createMemoryStorage();
    const context = loadCollectorLogic(null, { localStorage: storage });
    const result = JSON.parse(vm.runInContext(`JSON.stringify((() => {
        renderSequenceFilters = function renderSequenceFiltersForProjectScopeTest() {};
        latestPlan = {
            projectName: 'Project A',
            projectPath: 'C:/Projects/Project A.prproj',
            availableSequences: [{
                sequenceID: 'sequence-a',
                sequenceName: 'Sequence A'
            }],
            activeSequenceID: 'sequence-a',
            activeSequenceName: 'Sequence A',
            videoTrackUsage: [{ trackNumber: 1, clipCount: 1, hasItems: true }],
            audioTrackUsage: []
        };
        loadSequenceFilters();
        selectedSequenceFilters[0].ignoredVideoTracks = [1];
        selectedSequenceFilters.push(createSequenceFilter(
            'deleted-sequence',
            'Deleted Sequence',
            [{ trackNumber: 1, clipCount: 1, hasItems: true }],
            [],
            false
        ));
        saveSequenceFilters();
        loadSequenceFilters();
        const sameProjectFilters = selectedSequenceFilters.map((filter) => ({
            sequenceName: filter.sequenceName,
            ignoredVideoTracks: filter.ignoredVideoTracks.slice()
        }));

        latestPlan = {
            projectName: 'Project B',
            projectPath: 'C:/Projects/Project B.prproj',
            availableSequences: [{
                sequenceID: 'sequence-b',
                sequenceName: 'Sequence B'
            }],
            activeSequenceID: 'sequence-b',
            activeSequenceName: 'Sequence B',
            videoTrackUsage: [{ trackNumber: 1, clipCount: 1, hasItems: true }],
            audioTrackUsage: []
        };
        loadSequenceFilters();
        const projectBFilters = selectedSequenceFilters.map((filter) => filter.sequenceName);

        latestPlan = {
            projectName: 'Project A',
            projectPath: 'C:/Projects/Project A.prproj',
            availableSequences: [{
                sequenceID: 'sequence-a',
                sequenceName: 'Sequence A'
            }],
            activeSequenceID: 'sequence-a',
            activeSequenceName: 'Sequence A',
            videoTrackUsage: [{ trackNumber: 1, clipCount: 1, hasItems: true }],
            audioTrackUsage: []
        };
        loadSequenceFilters();
        const restoredProjectAFilters = selectedSequenceFilters.map((filter) => ({
            sequenceName: filter.sequenceName,
            ignoredVideoTracks: filter.ignoredVideoTracks.slice()
        }));

        return {
            sameProjectFilters,
            projectBFilters,
            restoredProjectAFilters
        };
    })())`, context));

    assert.deepEqual(result.sameProjectFilters.map((filter) => filter.sequenceName), ['Sequence A']);
    assert.deepEqual(result.sameProjectFilters[0].ignoredVideoTracks, [1]);
    assert.deepEqual(result.projectBFilters, ['Sequence B']);
    assert.deepEqual(result.restoredProjectAFilters.map((filter) => filter.sequenceName), ['Sequence A']);
    assert.deepEqual(result.restoredProjectAFilters[0].ignoredVideoTracks, [1]);
});

test('Premiere project plan reports every sequence available in the open project', () => {
    const context = loadHostRelinkLogic();
    const sequenceA = {
        name: 'Sequence A',
        sequenceID: 'sequence-a',
        videoTracks: { numTracks: 0 },
        audioTracks: { numTracks: 0 }
    };
    const sequenceB = {
        name: 'Sequence B',
        sequenceID: 'sequence-b',
        videoTracks: { numTracks: 0 },
        audioTracks: { numTracks: 0 }
    };
    const sequences = {
        0: sequenceA,
        1: sequenceB,
        numSequences: 2
    };
    context.app = {
        project: {
            name: 'Scoped Project.prproj',
            path: 'C:/Projects/Scoped Project.prproj',
            rootItem: {
                type: 2,
                children: createHostChildren([])
            },
            sequences,
            activeSequence: sequenceB
        }
    };

    const plan = JSON.parse(context.getProjectCopyPlan('D:/Backup'));

    assert.deepEqual(
        Array.from(plan.availableSequences, (sequence) => ({
            sequenceID: sequence.sequenceID,
            sequenceName: sequence.sequenceName
        })),
        [
            { sequenceID: 'sequence-a', sequenceName: 'Sequence A' },
            { sequenceID: 'sequence-b', sequenceName: 'Sequence B' }
        ]
    );
    assert.equal(plan.activeSequenceID, 'sequence-b');
});

test('saving duplicate track settings selects the existing preset and does not prompt or duplicate', () => {
    let alertMessage = '';
    let promptCount = 0;
    const context = loadCollectorLogic(createTestDocument(), {
        alert(message) {
            alertMessage = message;
        },
        prompt() {
            promptCount += 1;
            return 'Should Not Save';
        }
    });

    const result = JSON.parse(vm.runInContext(`JSON.stringify((() => {
        trackPresetsLoaded = true;
        trackPresets = [{
            id: 'existing-preset',
            name: 'Existing Preset',
            ignoredVideoTracks: [2],
            ignoredAudioTracks: [1]
        }];
        selectedSequenceFilters = [
            createSequenceFilter(
                'sequence-a',
                'Sequence A',
                [
                    { trackNumber: 1, clipCount: 1, hasItems: true },
                    { trackNumber: 2, clipCount: 1, hasItems: true }
                ],
                [{ trackNumber: 1, clipCount: 1, hasItems: true }],
                false
            )
        ];
        selectedSequenceFilters[0].ignoredVideoTracks = [2];
        selectedSequenceFilters[0].ignoredAudioTracks = [1];
        renderSequenceFilters = function renderSequenceFiltersForDuplicatePresetTest() {};
        updateSelectionSummary = function updateSelectionSummaryForDuplicatePresetTest() {};

        saveTrackPresetForSequence('sequence-a');
        return {
            presets: trackPresets,
            filter: selectedSequenceFilters[0]
        };
    })())`, context));

    assert.equal(promptCount, 0);
    assert.match(alertMessage, /already exists/i);
    assert.match(alertMessage, /Existing Preset/);
    assert.equal(result.presets.length, 1);
    assert.equal(result.filter.selectedPresetId, 'existing-preset');
});

test('deleting a preset removes it from the shared list but preserves current track choices', () => {
    const storage = createMemoryStorage();
    const context = loadCollectorLogic(createTestDocument(), { localStorage: storage });
    const result = JSON.parse(vm.runInContext(`JSON.stringify((() => {
        trackPresetsLoaded = true;
        trackPresets = [
            {
                id: 'delete-me',
                name: 'Delete Me',
                ignoredVideoTracks: [2],
                ignoredAudioTracks: [1]
            },
            {
                id: 'keep-me',
                name: 'Keep Me',
                ignoredVideoTracks: [1],
                ignoredAudioTracks: []
            }
        ];
        selectedSequenceFilters = [
            createSequenceFilter(
                'sequence-a',
                'Sequence A',
                [
                    { trackNumber: 1, clipCount: 1, hasItems: true },
                    { trackNumber: 2, clipCount: 1, hasItems: true }
                ],
                [{ trackNumber: 1, clipCount: 1, hasItems: true }],
                false
            ),
            createSequenceFilter(
                'sequence-b',
                'Sequence B',
                [
                    { trackNumber: 1, clipCount: 1, hasItems: true },
                    { trackNumber: 2, clipCount: 1, hasItems: true }
                ],
                [{ trackNumber: 1, clipCount: 1, hasItems: true }],
                false
            )
        ];
        selectedSequenceFilters.forEach((filter) => applyTrackPresetToFilter(filter, trackPresets[0]));
        const sectionBeforeDelete = renderSequencePresetSection(selectedSequenceFilters[0]);
        const deleteButtonBefore = sectionBeforeDelete.children.find(
            (control) => String(control.className || '').indexOf('sequence-preset-delete') !== -1
        );
        const deleted = deleteTrackPreset('delete-me');

        return {
            deleted,
            deleteButtonLabel: deleteButtonBefore.textContent,
            deleteButtonDisabled: deleteButtonBefore.disabled,
            presets: trackPresets,
            filters: selectedSequenceFilters,
            storedPresets: JSON.parse(localStorage.getItem(TRACK_PRESETS_STORAGE_KEY))
        };
    })())`, context));

    assert.equal(result.deleteButtonLabel, 'Delete');
    assert.equal(result.deleteButtonDisabled, false);
    assert.equal(result.deleted, true);
    assert.deepEqual(result.presets.map((preset) => preset.id), ['keep-me']);
    assert.deepEqual(result.storedPresets.map((preset) => preset.id), ['keep-me']);
    result.filters.forEach((filter) => {
        assert.equal(filter.selectedPresetId, '');
        assert.deepEqual(filter.ignoredVideoTracks, [2]);
        assert.deepEqual(filter.ignoredAudioTracks, [1]);
    });
});

test('BACKUP relink plan targets collected, verified skip-location, and original paths correctly', () => {
    const context = loadCollectorLogic();
    const result = JSON.parse(vm.runInContext(`JSON.stringify((() => {
        const allTasks = [
            { source: 'C:/Original/copied.mov' },
            { source: 'C:/Original/ignored.mov' },
            { source: 'C:/Original/compare-skipped.mov' },
            { source: 'C:/Original/copy-failed.mov' }
        ];
        const copiedTasks = [{
            source: 'C:/Original/copied.mov',
            destinationPath: 'D:/Backup/CollectedMedias/copied.mov'
        }];
        const compareMatches = [{
            task: allTasks[2],
            match: {
                path: 'E:/Skip Library/compare-skipped.mov'
            }
        }];
        return buildLinkProjectTasks(allTasks, copiedTasks, compareMatches);
    })())`, context));

    assert.deepEqual(result, [
        {
            source: 'C:/Original/copied.mov',
            destination: 'D:/Backup/CollectedMedias/copied.mov',
            targetKind: 'collected'
        },
        {
            source: 'C:/Original/ignored.mov',
            destination: 'C:/Original/ignored.mov',
            targetKind: 'original'
        },
        {
            source: 'C:/Original/compare-skipped.mov',
            destination: 'E:/Skip Library/compare-skipped.mov',
            targetKind: 'skip-location'
        },
        {
            source: 'C:/Original/copy-failed.mov',
            destination: 'C:/Original/copy-failed.mov',
            targetKind: 'original'
        }
    ]);
});

test('skip-location matching requires identical SHA-256 content, not only name and size', async (t) => {
    const tempRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'project-collector-skip-hash-'));
    const sourceDir = path.join(tempRoot, 'source');
    const wrongDir = path.join(tempRoot, 'skip-wrong');
    const correctDir = path.join(tempRoot, 'skip-correct');
    fs.mkdirSync(sourceDir, { recursive: true });
    fs.mkdirSync(wrongDir, { recursive: true });
    fs.mkdirSync(correctDir, { recursive: true });
    const sourcePath = path.join(sourceDir, 'clip.mov');
    const wrongPath = path.join(wrongDir, 'clip.mov');
    const correctPath = path.join(correctDir, 'clip.mov');
    fs.writeFileSync(sourcePath, 'AAAA-same-size');
    fs.writeFileSync(wrongPath, 'BBBB-same-size');
    fs.writeFileSync(correctPath, 'AAAA-same-size');
    t.after(() => fs.rmSync(tempRoot, { recursive: true, force: true }));

    const context = loadCollectorLogic();
    const wrongOnlyLookup = context.buildCompareLookup([{
        name: 'clip.mov',
        path: wrongPath,
        size: fs.statSync(wrongPath).size
    }]);
    const wrongMatch = await context.findCompareMatchForTask({ source: sourcePath }, wrongOnlyLookup);
    assert.equal(wrongMatch, null);

    const mixedLookup = context.buildCompareLookup([
        {
            name: 'clip.mov',
            path: wrongPath,
            size: fs.statSync(wrongPath).size
        },
        {
            name: 'clip.mov',
            path: correctPath,
            size: fs.statSync(correctPath).size
        }
    ]);
    const verifiedMatch = await context.findCompareMatchForTask({ source: sourcePath }, mixedLookup);

    assert.equal(verifiedMatch.path, correctPath);
    assert.match(verifiedMatch.sha256, /^[a-f0-9]{64}$/);
});

test('copied project keeps the original project filename and appends BACKUP', () => {
    const context = loadCollectorLogic();
    const names = JSON.parse(vm.runInContext(`JSON.stringify((() => {
        sequenceOnlyMode = true;
        selectedSequenceFilters = [{
            sequenceID: 'different-sequence',
            sequenceName: 'This Sequence Name Must Not Be Used'
        }];
        latestPlan = {
            projectName: 'Different Collection Folder'
        };
        return {
            standard: getCollectedProjectFileName('D:/Projects/TEASER FN 20260721.prproj'),
            mixedCaseExtension: getCollectedProjectFileName('D:/Projects/Evening Show.PRPROJ')
        };
    })())`, context));

    assert.equal(names.standard, 'TEASER FN 20260721 BACKUP.prproj');
    assert.equal(names.mixedCaseExtension, 'Evening Show BACKUP.PRPROJ');
});

test('project copy physically creates the requested BACKUP filename even beside the source project', async (t) => {
    const tempRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'project-collector-project-name-'));
    const sourceProjectPath = path.join(tempRoot, 'temp.prproj');
    const expectedBackupPath = path.join(tempRoot, 'temp BACKUP.prproj');
    fs.writeFileSync(sourceProjectPath, 'premiere-project-data');
    t.after(() => fs.rmSync(tempRoot, { recursive: true, force: true }));

    const context = loadCollectorLogic();
    const result = await context.copyProjectFileIntoCollectedRoot(
        tempRoot,
        sourceProjectPath,
        context.getCollectedProjectFileName(sourceProjectPath)
    );

    assert.equal(result.success, true);
    assert.equal(result.destinationPath, expectedBackupPath);
    assert.equal(fs.existsSync(sourceProjectPath), true);
    assert.equal(fs.existsSync(expectedBackupPath), true);
    assert.equal(fs.readFileSync(expectedBackupPath, 'utf8'), 'premiere-project-data');
    assert.deepEqual(
        fs.readdirSync(tempRoot).filter((name) => name.startsWith('.projectcollector-copy-')),
        []
    );
});

test('collection sends the complete copied-versus-ignored relink plan to Premiere', async (t) => {
    const tempRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'project-collector-relink-plan-'));
    const sourceDir = path.join(tempRoot, 'source');
    const destinationDir = path.join(tempRoot, 'backup');
    const skipLocationDir = path.join(tempRoot, 'skip-location', 'extra', 'nested');
    const copiedSourcePath = path.join(sourceDir, 'copied.mov');
    const ignoredSourcePath = path.join(sourceDir, 'ignored.mov');
    const skipSourcePath = path.join(sourceDir, 'skip.mov');
    const skipMatchPath = path.join(skipLocationDir, 'skip.mov');
    const copiedDestinationPath = path.join(destinationDir, 'Relink_Project', 'CollectedMedias', 'copied.mov');
    const projectCopyPath = path.join(destinationDir, 'Relink_Project', 'Relink Project BACKUP.prproj');
    fs.mkdirSync(sourceDir, { recursive: true });
    fs.mkdirSync(skipLocationDir, { recursive: true });
    fs.writeFileSync(copiedSourcePath, 'copied-media');
    fs.writeFileSync(ignoredSourcePath, 'ignored-media');
    fs.writeFileSync(skipSourcePath, 'verified-skip-media');
    fs.writeFileSync(skipMatchPath, 'verified-skip-media');
    t.after(() => fs.rmSync(tempRoot, { recursive: true, force: true }));

    const context = loadCollectorLogic(createTestDocument());
    vm.runInContext(`
        destination = ${JSON.stringify(destinationDir)};
        compareLocation = ${JSON.stringify(path.join(tempRoot, 'skip-location'))};
        copyProjectFile = true;
        linkProjectAfterCollection = true;
        sequenceOnlyMode = false;
        createReducedProject = false;
        latestPlan = {
            projectName: 'Relink_Project',
            projectPath: 'C:/Projects/Relink Project.prproj',
            missingMedia: [],
            tasks: [{
                name: 'copied.mov',
                source: ${JSON.stringify(copiedSourcePath)},
                destination: 'CollectedMedias/copied.mov',
                binPath: 'CollectedMedias',
                relativePath: 'CollectedMedias/copied.mov'
            }, {
                name: 'ignored.mov',
                source: ${JSON.stringify(ignoredSourcePath)},
                destination: 'CollectedMedias/ignored.mov',
                binPath: 'CollectedMedias',
                relativePath: 'CollectedMedias/ignored.mov'
            }, {
                name: 'skip.mov',
                source: ${JSON.stringify(skipSourcePath)},
                destination: 'CollectedMedias/skip.mov',
                binPath: 'CollectedMedias',
                relativePath: 'CollectedMedias/skip.mov'
            }]
        };
        sourceTree = [];
        buildCopyReadyContext = async function buildCopyReadyContextForRelinkPlanTest() {
            const ignoredMediaSet = new Set([normalizeMediaKey(latestPlan.tasks[1].source)]);
            return {
                ok: true,
                copyRuleContext: createCopyRuleContext({
                    treeSelectedTaskSet: new Set(latestPlan.tasks),
                    trackRuleContext: buildTrackRuleContext(latestPlan.tasks, new Set(), ignoredMediaSet)
                }),
                copyWarnings: [],
                sequenceScopeInfo: null,
                trackConflicts: []
            };
        };
        copiedProjectDestinationRoot = '';
        copyProjectFileIntoCollectedRoot = async function copyProjectForRelinkPlanTest(rootPath) {
            copiedProjectDestinationRoot = rootPath;
            return {
                success: true,
                destinationPath: ${JSON.stringify(projectCopyPath)}
            };
        };
    `, context);

    let receivedRelinkTasks = null;
    context.callHost = async (script) => {
        if (script === 'saveCurrentProjectAndGetPath()') {
            return JSON.stringify({ projectPath: 'C:/Projects/Relink Project.prproj' });
        }
        if (script.startsWith('linkProjectCopyToCollectedMedia(')) {
            const match = script.match(/^linkProjectCopyToCollectedMedia\("(?:\\.|[^"])+"\,\"((?:\\.|[^"])*)"\)$/);
            assert.ok(match, `Unexpected relink host call: ${script}`);
            receivedRelinkTasks = JSON.parse(JSON.parse(`"${match[1]}"`));
            return JSON.stringify({
                success: true,
                offlineCount: 3,
                linkedCount: 3,
                linkedCollectedCount: 1,
                linkedSkipLocationCount: 1,
                linkedOriginalCount: 1,
                skippedNoPathCount: 0,
                failed: []
            });
        }
        throw new Error(`Unexpected host call: ${script}`);
    };

    await vm.runInContext('collect()', context);

    assert.equal(fs.existsSync(copiedDestinationPath), true);
    assert.equal(fs.existsSync(path.join(destinationDir, 'Relink_Project', 'CollectedMedias', 'ignored.mov')), false);
    assert.equal(fs.existsSync(path.join(destinationDir, 'Relink_Project', 'CollectedMedias', 'skip.mov')), false);
    assert.equal(
        vm.runInContext('copiedProjectDestinationRoot', context),
        path.join(destinationDir, 'Relink_Project')
    );
    assert.deepEqual(receivedRelinkTasks, [
        {
            source: copiedSourcePath,
            destination: copiedDestinationPath,
            targetKind: 'collected'
        },
        {
            source: ignoredSourcePath,
            destination: ignoredSourcePath,
            targetKind: 'original'
        },
        {
            source: skipSourcePath,
            destination: skipMatchPath,
            targetKind: 'skip-location'
        }
    ]);
});

function loadHostRelinkLogic() {
    const scriptPath = path.join(__dirname, '..', 'jsx', 'collector.jsx');
    const source = fs.readFileSync(scriptPath, 'utf8');
    const context = vm.createContext({
        File: function File(filePath) {
            this.fsName = filePath;
            this.exists = true;
        },
        JSON,
        ProjectItemType: { BIN: 2 },
        String
    });
    vm.runInContext(source, context);
    return context;
}

function createHostChildren(items) {
    const children = { numItems: items.length };
    items.forEach((item, index) => {
        children[index] = item;
    });
    return children;
}

function createHostMediaItem(name, mediaPath, events, options) {
    const settings = options || {};
    let currentPath = mediaPath;
    let offline = false;

    return {
        name,
        type: 1,
        getMediaPath() {
            return currentPath;
        },
        isOffline() {
            return offline;
        },
        setOffline() {
            events.push(`offline:${name}`);
            if (settings.offlineFails) {
                return false;
            }
            offline = true;
            currentPath = '';
            return true;
        },
        changeMediaPath(destination) {
            events.push(`link:${name}:${destination}`);
            if (settings.linkFails) {
                return 1;
            }
            currentPath = destination;
            offline = false;
            return 0;
        }
    };
}

test('Premiere host plan ignores pathless project structures but reports genuine offline media', () => {
    const context = loadHostRelinkLogic();
    let sequenceMediaPathRead = false;
    const sequenceItem = {
        name: 'Baby wakes mom up in the sweetest way possible@Daily Mail - nest',
        type: 1,
        isSequence() {
            return true;
        },
        getMediaPath() {
            sequenceMediaPathRead = true;
            return '';
        }
    };
    const generatedItem = {
        name: 'Adjustment Layer',
        type: 1,
        isSequence() {
            return false;
        },
        isOffline() {
            return false;
        },
        getMediaPath() {
            return '';
        }
    };
    const offlineItem = {
        name: 'Offline Interview',
        type: 1,
        isSequence() {
            return false;
        },
        isOffline() {
            return true;
        },
        getMediaPath() {
            return '';
        }
    };
    const unreadableItem = {
        name: 'Unreadable Host Item',
        type: 1,
        isSequence() {
            return false;
        },
        isOffline() {
            return false;
        },
        getMediaPath() {
            throw new Error('Host media-path read failed');
        }
    };
    const networkItem = {
        name: 'Network Clip',
        type: 1,
        isSequence() {
            return false;
        },
        isOffline() {
            return false;
        },
        getMediaPath() {
            return '\\\\media-server\\news\\network-clip.mov';
        }
    };

    context.app = {
        project: {
            name: 'Portable News Project.prproj',
            path: 'C:/Projects/Portable News Project.prproj',
            rootItem: {
                type: 2,
                children: createHostChildren([
                    sequenceItem,
                    generatedItem,
                    offlineItem,
                    unreadableItem,
                    networkItem
                ])
            }
        }
    };

    const plan = JSON.parse(context.getProjectCopyPlan('\\\\backup-server\\projects'));

    assert.equal(sequenceMediaPathRead, false);
    assert.equal(plan.tasks.length, 1);
    assert.equal(plan.tasks[0].source, '\\\\media-server\\news\\network-clip.mov');
    assert.equal(plan.missingMedia.length, 2);
    assert.equal(
        plan.missingMedia[0],
        'Offline Interview | Media is offline and no media path is available'
    );
    assert.match(
        plan.missingMedia[1],
        /^Unreadable Host Item \| Could not read media path: Error: Host media-path read failed$/
    );
});

test('Premiere host relink takes every file offline before linking copied and ignored media', () => {
    const context = loadHostRelinkLogic();
    const events = [];
    const copiedItem = createHostMediaItem('Copied', 'C:/Original/copied.mov', events);
    const ignoredItem = createHostMediaItem('Ignored', 'C:/Original/ignored.mov', events);
    const skipItem = createHostMediaItem('Skip', 'C:/Original/skip.mov', events);
    const generatedItem = {
        name: 'Generated',
        type: 1,
        getMediaPath() {
            return '';
        }
    };
    const copiedProject = {
        path: 'D:/Backup/My Project BACKUP.prproj',
        rootItem: {
            type: 2,
            children: createHostChildren([copiedItem, ignoredItem, skipItem, generatedItem])
        },
        save() {
            events.push('save:backup');
            return 0;
        },
        closeDocument(saveFirst, promptIfDirty) {
            events.push(`close:backup:${saveFirst}:${promptIfDirty}`);
            return 0;
        }
    };
    const originalProject = {
        path: 'C:/Projects/My Project.prproj',
        rootItem: { type: 2, children: createHostChildren([]) },
        save() {
            events.push('save:original');
            return 0;
        },
        closeDocument(saveFirst, promptIfDirty) {
            events.push(`close:original:${saveFirst}:${promptIfDirty}`);
            app.project = null;
            return 0;
        }
    };
    const app = {
        project: originalProject,
        openDocument(projectPath) {
            events.push(`open:${projectPath}`);
            this.project = projectPath === copiedProject.path ? copiedProject : originalProject;
            return true;
        }
    };
    context.app = app;

    const tasks = JSON.stringify([
        {
            source: 'C:/Original/copied.mov',
            destination: 'D:/Backup/CollectedMedias/copied.mov',
            targetKind: 'collected'
        },
        {
            source: 'C:/Original/ignored.mov',
            destination: 'C:/Original/ignored.mov',
            targetKind: 'original'
        },
        {
            source: 'C:/Original/skip.mov',
            destination: 'E:/Skip Library/skip.mov',
            targetKind: 'skip-location'
        }
    ]);
    const result = JSON.parse(context.linkProjectCopyToCollectedMedia(copiedProject.path, tasks));

    assert.equal(result.success, true);
    assert.equal(result.offlineCount, 3);
    assert.equal(result.linkedCount, 3);
    assert.equal(result.linkedCollectedCount, 1);
    assert.equal(result.linkedSkipLocationCount, 1);
    assert.equal(result.linkedOriginalCount, 1);
    assert.equal(result.skippedNoPathCount, 1);
    assert.deepEqual(result.failed, []);
    assert.ok(events.indexOf('save:original') < events.indexOf('close:original:0:0'));
    assert.ok(events.indexOf('close:original:0:0') < events.indexOf(`open:${copiedProject.path}`));
    assert.ok(events.indexOf('offline:Copied') < events.indexOf('link:Copied:D:/Backup/CollectedMedias/copied.mov'));
    assert.ok(events.indexOf('offline:Ignored') < events.indexOf('link:Copied:D:/Backup/CollectedMedias/copied.mov'));
    assert.ok(events.indexOf('link:Copied:D:/Backup/CollectedMedias/copied.mov') < events.indexOf('save:backup'));
    assert.equal(events.includes('close:backup:0:0'), false);
    assert.equal(copiedItem.getMediaPath(), 'D:/Backup/CollectedMedias/copied.mov');
    assert.equal(ignoredItem.getMediaPath(), 'C:/Original/ignored.mov');
    assert.equal(skipItem.getMediaPath(), 'E:/Skip Library/skip.mov');
    assert.equal(result.backupLeftOpen, true);
    assert.equal(app.project, copiedProject);
});

test('Premiere host finds the BACKUP in app.projects when the original remains the active project', () => {
    const context = loadHostRelinkLogic();
    const events = [];
    const copiedItem = createHostMediaItem('Copied', 'C:/Original/copied.mov', events);
    const copiedProject = {
        path: 'D:/Backup/My Project BACKUP.prproj',
        rootItem: {
            type: 2,
            children: createHostChildren([copiedItem])
        },
        save() {
            events.push('save:backup');
            return 0;
        },
        closeDocument() {
            events.push('close:backup');
            return 0;
        }
    };
    const originalProject = {
        path: 'C:/Projects/My Project.prproj',
        rootItem: { type: 2, children: createHostChildren([]) },
        save() {
            events.push('save:original');
            return 0;
        },
        closeDocument() {
            events.push('close:original');
            app.project = otherProject;
            projects[1] = otherProject;
            projects.numProjects = 1;
            return 0;
        }
    };
    const otherProject = {
        path: 'C:/Projects/Other Project.prproj',
        rootItem: { type: 2, children: createHostChildren([]) }
    };
    const projects = {
        1: originalProject,
        numProjects: 1
    };
    const app = {
        project: originalProject,
        projects,
        openDocument(projectPath) {
            events.push(`open:${projectPath}`);
            if (projectPath === copiedProject.path) {
                projects[2] = copiedProject;
                projects.numProjects = 2;
            } else if (projectPath === originalProject.path) {
                this.project = originalProject;
            }
            return true;
        }
    };
    context.app = app;

    const result = JSON.parse(context.linkProjectCopyToCollectedMedia(
        copiedProject.path,
        JSON.stringify([{
            source: 'C:/Original/copied.mov',
            destination: 'D:/Backup/CollectedMedias/copied.mov',
            targetKind: 'collected'
        }])
    ));

    assert.equal(result.success, true);
    assert.equal(result.linkedCollectedCount, 1);
    assert.equal(copiedItem.getMediaPath(), 'D:/Backup/CollectedMedias/copied.mov');
    assert.equal(events.includes('save:backup'), true);
    assert.equal(events.includes('close:backup'), false);
    assert.equal(events.includes('close:original'), true);
    assert.equal(result.backupLeftOpen, true);
    assert.equal(app.project, otherProject);
});

test('Premiere host does not relink or save a partially offlined BACKUP project', () => {
    const context = loadHostRelinkLogic();
    const events = [];
    const firstItem = createHostMediaItem('First', 'C:/Original/first.mov', events);
    const blockedItem = createHostMediaItem('Blocked', 'C:/Original/blocked.mov', events, {
        offlineFails: true
    });
    const copiedProject = {
        path: 'D:/Backup/My Project BACKUP.prproj',
        rootItem: {
            type: 2,
            children: createHostChildren([firstItem, blockedItem])
        },
        save() {
            events.push('save:backup');
            return 0;
        },
        closeDocument(saveFirst, promptIfDirty) {
            events.push(`close:backup:${saveFirst}:${promptIfDirty}`);
            return 0;
        }
    };
    const originalProject = {
        path: 'C:/Projects/My Project.prproj',
        rootItem: { type: 2, children: createHostChildren([]) },
        save() {
            events.push('save:original');
            return 0;
        },
        closeDocument() {
            events.push('close:original');
            app.project = null;
            return 0;
        }
    };
    const app = {
        project: originalProject,
        openDocument(projectPath) {
            this.project = projectPath === copiedProject.path ? copiedProject : originalProject;
            return true;
        }
    };
    context.app = app;

    const tasks = JSON.stringify([
        {
            source: 'C:/Original/first.mov',
            destination: 'D:/Backup/CollectedMedias/first.mov',
            targetKind: 'collected'
        },
        {
            source: 'C:/Original/blocked.mov',
            destination: 'C:/Original/blocked.mov',
            targetKind: 'original'
        }
    ]);
    const result = JSON.parse(context.linkProjectCopyToCollectedMedia(copiedProject.path, tasks));

    assert.equal(result.success, false);
    assert.equal(result.offlineCount, 1);
    assert.equal(result.linkedCount, 0);
    assert.match(result.failed.join('\n'), /could not take the item offline/i);
    assert.equal(events.some((event) => event.startsWith('link:')), false);
    assert.equal(events.includes('save:backup'), false);
    assert.equal(events.includes('close:backup:0:0'), true);
    assert.equal(app.project, originalProject);
});
