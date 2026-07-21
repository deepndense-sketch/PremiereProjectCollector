const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

function loadCollectorLogic() {
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
        document: {
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
        }
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
