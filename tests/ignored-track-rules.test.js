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
        'summaryText',
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

test('ignored track wins when the same media is also included', () => {
    const context = loadCollectorLogic();
    const reason = vm.runInContext(`(() => {
        const task = { name: 'shared.mov', source: 'C:/Media/shared.mov', binPath: '' };
        const mediaKey = normalizeMediaKey(task.source);
        return getCopySkipReason(
            task,
            new Set([task]),
            new Set([mediaKey]),
            new Set([mediaKey]),
            new Set()
        );
    })()`, context);

    assert.equal(reason, 'skipped by ignored track selection');
});

test('ignored Premiere project folder still wins', () => {
    const context = loadCollectorLogic();
    const reason = vm.runInContext(`(() => {
        const task = { name: 'shared.mov', source: 'C:/Media/shared.mov', binPath: 'Do Not Back Up' };
        ignoredProjectFolders = ['Do Not Back Up'];
        return getCopySkipReason(task, new Set([task]), null, new Set(), new Set());
    })()`, context);

    assert.equal(reason, 'skipped because Premiere folder "Do Not Back Up" is ignored');
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
            const ignoredSignatureSet = new Set();
            const selectedTasksBeforeCompare = latestPlan.tasks.filter((task) => !getCopySkipReason(
                task,
                treeSelectedTaskSet,
                null,
                ignoredMediaSet,
                ignoredSignatureSet
            ));
            return {
                ok: true,
                treeSelectedTaskSet,
                ignoredMediaSet,
                ignoredSignatureSet,
                copyWarnings: [],
                sequenceScopedMediaSet: null,
                sequenceScopeInfo: null,
                selectedTasksBeforeCompare
            };
        };
    `, context);

    await vm.runInContext('collect()', context);

    assert.equal(fs.existsSync(path.join(collectedRoot, 'shared.mov')), false);
    assert.equal(fs.existsSync(path.join(collectedRoot, 'allowed.mov')), true);
    assert.equal(fs.readFileSync(path.join(collectedRoot, 'allowed.mov'), 'utf8'), 'allowed-media');
    assert.equal(vm.runInContext('isCopying', context), false);
});
