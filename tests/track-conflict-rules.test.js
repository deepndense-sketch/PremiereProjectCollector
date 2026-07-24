const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

function createTestElement() {
    return {
        children: [],
        classList: {
            add() {},
            remove() {},
            toggle() {}
        },
        disabled: false,
        innerHTML: '',
        style: {},
        textContent: '',
        appendChild(child) {
            this.children.push(child);
        },
        querySelector() {
            return null;
        },
        setAttribute() {}
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

function loadCollectorLogic(testDocument) {
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
        window: {
            addEventListener() {}
        },
        alert() {}
    });

    vm.runInContext(source, context, { filename: scriptPath });
    return context;
}

test('lists only media used by both included and ignored tracks', () => {
    const context = loadCollectorLogic();
    const result = vm.runInContext(`(() => {
        const shared = { name: 'shared.mov', source: 'C:/Media/shared.mov', binPath: '' };
        const ignoredOnly = { name: 'ignored.mov', source: 'C:/Media/ignored.mov', binPath: '' };
        const included = new Set([normalizeMediaKey(shared.source)]);
        const ignored = new Set([normalizeMediaKey(shared.source), normalizeMediaKey(ignoredOnly.source)]);
        return buildTrackSelectionConflicts([shared, ignoredOnly], included, ignored, new Set())
            .map((conflict) => conflict.name);
    })()`, context);

    assert.deepEqual(Array.from(result), ['shared.mov']);
});

test('Copy and Do not copy decisions control only the ignored-track rule', () => {
    const context = loadCollectorLogic();
    const result = vm.runInContext(`(() => {
        const task = { name: 'shared.mov', source: 'C:/Media/shared.mov', binPath: '' };
        const mediaKey = normalizeMediaKey(task.source);
        const ignored = new Set([mediaKey]);
        const conflicts = buildTrackSelectionConflicts([task], new Set([mediaKey]), ignored, new Set());
        const copyInfo = buildTrackConflictDecisionInfo(conflicts, new Map([[conflicts[0].id, 'copy']]));
        const skipInfo = buildTrackConflictDecisionInfo(conflicts, new Map([[conflicts[0].id, 'skip']]));
        const selected = new Set([task]);
        const scope = new Set([mediaKey]);

        return {
            copyReason: getCopySkipReason(task, selected, scope, ignored, new Set(), copyInfo),
            skipReason: getCopySkipReason(task, selected, scope, ignored, new Set(), skipInfo),
            sourceListReason: getCopySkipReason(task, new Set(), scope, ignored, new Set(), copyInfo)
        };
    })()`, context);

    assert.equal(result.copyReason, '');
    assert.equal(result.skipReason, 'skipped by your track conflict decision');
    assert.equal(result.sourceListReason, 'skipped by Source File List selection');
});

test('Continue Backup returns CEP-safe plain decisions that the copy filter consumes', () => {
    const context = loadCollectorLogic();
    const result = vm.runInContext(`(() => {
        const task = { name: 'shared.mov', source: 'C:/Media/shared.mov', binPath: '' };
        const mediaKey = normalizeMediaKey(task.source);
        const ignored = new Set([mediaKey]);
        const conflicts = buildTrackSelectionConflicts([task], new Set([mediaKey]), ignored, new Set());
        let resolvedDecisions = null;

        trackConflictPromptState = {
            conflicts,
            decisions: createTrackConflictDecisionStore()
        };
        trackConflictPromptResolver = (decisions) => {
            resolvedDecisions = decisions;
        };

        setAllTrackConflictChoices('copy');
        finishTrackConflictPrompt(true);

        const decisionInfo = buildTrackConflictDecisionInfo(conflicts, resolvedDecisions);
        return {
            copied: getCopySkipReason(task, new Set([task]), null, ignored, new Set(), decisionInfo) === '',
            isPlainDecisionStore: !!resolvedDecisions
                && typeof resolvedDecisions.get === 'undefined'
                && resolvedDecisions[conflicts[0].id] === 'copy',
            promptCleared: trackConflictPromptResolver === null && trackConflictPromptState === null
        };
    })()`, context);

    assert.equal(result.copied, true);
    assert.equal(result.isPlainDecisionStore, true);
    assert.equal(result.promptCleared, true);
});

test('Continue Backup resumes the full collection loop and copies chosen conflict media', async (t) => {
    const tempRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'project-collector-conflict-'));
    const sourceDir = path.join(tempRoot, 'source');
    const destinationDir = path.join(tempRoot, 'backup');
    const sourcePath = path.join(sourceDir, 'shared.mov');
    const copiedPath = path.join(destinationDir, 'Conflict_Project', 'CollectedMedias', 'shared.mov');
    fs.mkdirSync(sourceDir, { recursive: true });
    fs.writeFileSync(sourcePath, 'conflict-media-copy-test');
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
            projectName: 'Conflict_Project',
            missingMedia: [],
            tasks: [{
                name: 'shared.mov',
                source: ${JSON.stringify(sourcePath)},
                destination: 'Media/shared.mov',
                binPath: '',
                relativePath: 'Media/shared.mov'
            }]
        };
        sourceTree = [];
        buildCopyReadyContext = async function buildCopyReadyContextForTest() {
            const task = latestPlan.tasks[0];
            const mediaKey = normalizeMediaKey(task.source);
            const ignoredMediaSet = new Set([mediaKey]);
            const ignoredSignatureSet = new Set();
            return {
                ok: true,
                treeSelectedTaskSet: new Set([task]),
                ignoredMediaSet,
                ignoredSignatureSet,
                copyWarnings: [],
                sequenceScopedMediaSet: null,
                sequenceScopeInfo: null,
                trackConflicts: buildTrackSelectionConflicts(
                    latestPlan.tasks,
                    new Set([mediaKey]),
                    ignoredMediaSet,
                    ignoredSignatureSet
                )
            };
        };
    `, context);

    const collectionPromise = vm.runInContext('collect()', context);
    await new Promise((resolve) => setImmediate(resolve));
    assert.equal(vm.runInContext('!!trackConflictPromptResolver', context), true);

    vm.runInContext(`
        setAllTrackConflictChoices('copy');
        finishTrackConflictPrompt(true);
    `, context);
    await collectionPromise;

    assert.equal(fs.existsSync(copiedPath), true);
    assert.equal(fs.readFileSync(copiedPath, 'utf8'), 'conflict-media-copy-test');
    assert.equal(vm.runInContext('isCopying', context), false);
});

test('an ignored Premiere project folder still wins over a track Copy decision', () => {
    const context = loadCollectorLogic();
    const reason = vm.runInContext(`(() => {
        const task = { name: 'shared.mov', source: 'C:/Media/shared.mov', binPath: 'Do Not Back Up' };
        const mediaKey = normalizeMediaKey(task.source);
        const ignored = new Set([mediaKey]);
        const conflicts = buildTrackSelectionConflicts([task], new Set([mediaKey]), ignored, new Set());
        const copyInfo = buildTrackConflictDecisionInfo(conflicts, new Map([[conflicts[0].id, 'copy']]));
        ignoredProjectFolders = ['Do Not Back Up'];
        return getCopySkipReason(task, new Set([task]), new Set([mediaKey]), ignored, new Set(), copyInfo);
    })()`, context);

    assert.equal(reason, 'skipped because Premiere folder "Do Not Back Up" is ignored');
});
