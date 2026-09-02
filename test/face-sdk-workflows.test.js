'use strict';

const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const repositoryRoot = path.resolve(__dirname, '..');
const updateWorkflowPath = '.github/workflows/face-sdk-update.yml';
const validationWorkflowPath = '.github/workflows/face-sdk-validation.yml';
const deploymentWorkflowPath = '.github/workflows/main_plataforma-backend-v3.yml';

function readRepositoryFile(relativePath) {
    return fs.readFileSync(path.join(repositoryRoot, relativePath), 'utf8');
}

function yamlMappingBlock(source, indentation, key) {
    const lines = source.replace(/\r\n/g, '\n').split('\n');
    const header = `${' '.repeat(indentation)}${key}:`;
    const start = lines.findIndex((line) => line === header);
    assert.notEqual(start, -1, `${key} must use block mapping syntax`);

    let end = start + 1;
    while (end < lines.length) {
        const line = lines[end];
        const leadingSpaces = line.match(/^ */)[0].length;
        if (line.trim() && leadingSpaces <= indentation) break;
        end += 1;
    }

    return lines.slice(start, end).join('\n');
}

function yamlDirectKeys(block, indentation) {
    const matcher = new RegExp(`^ {${indentation}}([A-Za-z0-9_-]+):(?:\\s|$)`);
    return block
        .split('\n')
        .map((line) => line.match(matcher)?.[1])
        .filter(Boolean);
}

function yamlNamedStep(source, name) {
    const lines = source.replace(/\r\n/g, '\n').split('\n');
    const header = `      - name: ${name}`;
    const start = lines.findIndex((line) => line === header);
    assert.notEqual(start, -1, `workflow must contain the step ${name}`);

    let end = start + 1;
    while (end < lines.length && !/^      - (?:name|uses):/.test(lines[end])) end += 1;
    return lines.slice(start, end).join('\n');
}

function referencedActions(source) {
    return Array.from(
        source.matchAll(/^\s*(?:-\s*)?uses:\s*(\S+)$/gm),
        ([, action]) => action,
    );
}

function uniqueMatches(source, expression) {
    return [...new Set(Array.from(source.matchAll(expression), ([, match]) => match))].sort();
}

const completeBackendCommands = [
    'node --check app.js',
    'node --check server.js',
    'node --check platform-row-authorization.js',
    'node --check domains/quote-requests.js',
    'node --check domains/conecta-recommendations.js',
    'node --check domains/client-onboarding.js',
    'node --check domains/learning-platform.js',
    'node --check domains/drm.js',
    'node --check domains/certificate-validation.js',
    'node --check integrations/microsoft-graph.js',
    'node --check integrations/azure-face.js',
    'node --check shared/retry.js',
    'node --check shared/escape-html.js',
    'node --check scripts/prepare-face-sdk-update.mjs',
    'npm test',
    'git diff --check',
];

test('trusted Face SDK updater is scheduled/manual, main-only, and least privilege', () => {
    const workflow = readRepositoryFile(updateWorkflowPath);
    const eventBlock = yamlMappingBlock(workflow, 0, 'on');
    const permissionsBlock = yamlMappingBlock(workflow, 0, 'permissions');
    const concurrencyBlock = yamlMappingBlock(workflow, 0, 'concurrency');
    const jobsBlock = yamlMappingBlock(workflow, 0, 'jobs');
    const proposeJob = yamlMappingBlock(workflow, 2, 'propose');

    assert.deepEqual(yamlDirectKeys(eventBlock, 2), ['schedule', 'workflow_dispatch']);
    assert.match(eventBlock, /^    - cron: '23 12 \* \* 2'$/m);
    assert.doesNotMatch(eventBlock, /\bpull_request(?:_target)?\b|\bpush\b|\bworkflow_(?:call|run)\b/);
    assert.deepEqual(yamlDirectKeys(permissionsBlock, 2), ['contents']);
    assert.match(permissionsBlock, /^  contents: read$/m);
    assert.deepEqual(yamlDirectKeys(concurrencyBlock, 2), ['group', 'cancel-in-progress']);
    assert.match(concurrencyBlock, /^  group: face-sdk-update$/m);
    assert.match(concurrencyBlock, /^  cancel-in-progress: false$/m);
    assert.deepEqual(yamlDirectKeys(jobsBlock, 2), ['propose']);
    assert.match(proposeJob, /github\.repository == 'IvyRoom\/backend'/);
    assert.match(proposeJob, /github\.ref == 'refs\/heads\/main'/);
    assert.match(proposeJob, /github\.event_name == 'schedule'/);
    assert.match(proposeJob, /github\.event_name == 'workflow_dispatch'/);
    assert.match(proposeJob, /^    environment:\n      name: face-sdk-automation\n      deployment: false$/m);
    assert.match(
        proposeJob,
        /^      FACE_SDK_WORKSPACE_NAME: face-sdk-\$\{\{ github\.run_id \}\}-\$\{\{ github\.run_attempt \}\}$/m,
    );
    assert.equal((workflow.match(/^permissions:/gm) || []).length, 1);
    assert.equal((workflow.match(/^[ \t]+permissions:/gm) || []).length, 0);

    assert.deepEqual(
        uniqueMatches(workflow, /\$\{\{\s*secrets\.([A-Z0-9_]+)\s*\}\}/g),
        ['AZURE_FACE_API_ENDPOINT', 'AZURE_FACE_API_KEY', 'FACE_SDK_APP_PRIVATE_KEY'],
    );
    assert.deepEqual(
        uniqueMatches(workflow, /\$\{\{\s*vars\.([A-Z0-9_]+)\s*\}\}/g),
        ['FACE_SDK_APP_CLIENT_ID'],
    );
    assert.doesNotMatch(workflow, /AZURE_FACE_API_NPM_TEMPORARY_TOKEN/);
    assert.doesNotMatch(workflow, /AZURE_AI_VISION_NPM_TOKEN_BASE64/);
    assert.doesNotMatch(workflow, /\benvironment:\s*(?:\n\s+name:\s*)?['"]?Production\b/i);
    assert.doesNotMatch(workflow, /^\s*id-token:/m);

    const allowedActionOwners = new Set([
        'actions/checkout',
        'actions/create-github-app-token',
        'actions/setup-node',
    ]);
    for (const action of referencedActions(workflow)) {
        const separator = action.lastIndexOf('@');
        assert.notEqual(separator, -1, `${action} must be pinned to a version`);
        assert.ok(allowedActionOwners.has(action.slice(0, separator)), `${action} is not allowed`);
        assert.match(action.slice(separator + 1), /^(?:v\d+(?:\.\d+\.\d+)?|[0-9a-f]{40})$/);
    }
    assert.doesNotMatch(workflow, /actions\/(?:upload|download)-artifact@/i);
    assert.doesNotMatch(workflow, /^\s*uses:\s*\.\/\.github\/workflows\//m);
});

test('private credentials are confined to preparation and publication token steps', () => {
    const workflow = readRepositoryFile(updateWorkflowPath);
    const prepareStep = yamlNamedStep(workflow, 'Download and inspect the private packages');
    const metadataStep = yamlNamedStep(workflow, 'Verify deterministic proposal metadata');
    const backendValidation = yamlNamedStep(workflow, 'Validate the complete Backend candidate');
    const sistemasValidation = yamlNamedStep(workflow, 'Validate the credential-free Sistemas candidate');
    const boundaryStep = yamlNamedStep(workflow, 'Verify candidate change boundaries');
    const cleanupStep = yamlNamedStep(workflow, 'Remove every temporary private package');
    const publicationTokenStep = yamlNamedStep(workflow, 'Create least-privilege publication token');

    assert.match(prepareStep, /^          AZURE_FACE_API_ENDPOINT: \$\{\{ secrets\.AZURE_FACE_API_ENDPOINT \}\}$/m);
    assert.match(prepareStep, /^          AZURE_FACE_API_KEY: \$\{\{ secrets\.AZURE_FACE_API_KEY \}\}$/m);
    assert.doesNotMatch(prepareStep, /FACE_SDK_APP_PRIVATE_KEY|publication-token\.outputs\.token/);

    for (const credentialFreeStep of [metadataStep, backendValidation, sistemasValidation, boundaryStep, cleanupStep]) {
        assert.doesNotMatch(credentialFreeStep, /\$\{\{\s*(?:secrets|vars)\./);
        assert.doesNotMatch(credentialFreeStep, /(?:GH_TOKEN|FACE_SDK_APP_TOKEN):/);
    }

    assert.match(publicationTokenStep, /^          repositories: \|\n            backend\n            sistemas$/m);
    assert.match(
        publicationTokenStep,
        /uses: actions\/create-github-app-token@bcd2ba49218906704ab6c1aa796996da409d3eb1/,
    );
    assert.match(publicationTokenStep, /^          client-id: \$\{\{ vars\.FACE_SDK_APP_CLIENT_ID \}\}$/m);
    assert.doesNotMatch(publicationTokenStep, /^\s+app-id:/m);
    assert.match(publicationTokenStep, /^          permission-contents: write$/m);
    assert.match(publicationTokenStep, /^          permission-pull-requests: write$/m);
    assert.doesNotMatch(publicationTokenStep, /permission-(?:actions|deployments|environments|workflows):/);
    assert.equal((workflow.match(/actions\/create-github-app-token@/g) || []).length, 1);
});

test('updater prepares, validates, cleans, and then publishes both proposals', () => {
    const workflow = readRepositoryFile(updateWorkflowPath);
    const checkoutSteps = [
        yamlNamedStep(workflow, 'Check out Backend main without persisted credentials'),
        yamlNamedStep(workflow, 'Check out Sistemas main without persisted credentials'),
    ];
    const prepareStep = yamlNamedStep(workflow, 'Download and inspect the private packages');
    const cleanupStep = yamlNamedStep(workflow, 'Remove every temporary private package');
    const commitStep = yamlNamedStep(workflow, 'Commit both validated candidates');
    const boundaryStep = yamlNamedStep(workflow, 'Verify candidate change boundaries');
    const publicationStep = yamlNamedStep(
        workflow,
        'Publish branches and idempotent cross-linked draft pull requests',
    );

    for (const expectedFragment of [
        'node scripts/prepare-face-sdk-update.mjs',
        '--backend-root "$GITHUB_WORKSPACE/backend"',
        '--sistemas-root "$GITHUB_WORKSPACE/sistemas"',
        '--workspace "$workspace"',
        '--temp-root "$RUNNER_TEMP"',
        '--requested-version "$requested_version"',
        '--github-output "$GITHUB_OUTPUT"',
    ]) {
        assert.ok(prepareStep.includes(expectedFragment), `${expectedFragment} must be passed to the preparer`);
    }
    for (const checkoutStep of checkoutSteps) {
        assert.match(checkoutStep, /^          fetch-depth: 0$/m);
        assert.match(checkoutStep, /^          persist-credentials: false$/m);
    }
    assert.match(workflow, /steps\.prepare\.outputs\['update-available'\] == 'true'/);
    assert.match(workflow, /steps\.prepare\.outputs\['current-version'\]/);
    assert.match(workflow, /steps\.prepare\.outputs\['target-version'\]/);
    assert.match(workflow, /expected_branch="chore\/update-face-liveness-sdk-\$\{PROPOSED_VERSION\/\/\.\/-\}"/);

    const backendValidation = yamlNamedStep(workflow, 'Validate the complete Backend candidate');
    for (const command of completeBackendCommands) {
        assert.ok(backendValidation.split('\n').some((line) => line.trim() === command));
    }
    const sistemasValidation = yamlNamedStep(workflow, 'Validate the credential-free Sistemas candidate');
    for (const command of [
        'node --check scripts/face-sdk-vendor-lib.mjs',
        'node --check scripts/update-face-sdk-vendor.mjs',
        'node --check scripts/check-face-sdk-vendor.mjs',
        'node scripts/check-face-sdk-vendor.mjs',
        'git diff --check',
    ]) {
        assert.ok(sistemasValidation.split('\n').some((line) => line.trim() === command));
    }
    assert.doesNotMatch(sistemasValidation, /node --test/);

    assert.match(cleanupStep, /^        if: always\(\)$/m);
    assert.match(cleanupStep, /path\.dirname\(workspace\) !== temporaryRoot/);
    assert.match(cleanupStep, /startsWith\("face-sdk-"\)/);
    assert.match(cleanupStep, /rm\(workspace, \{ recursive: true, force: true \}\)/);
    assert.ok(workflow.indexOf('Remove every temporary private package') < workflow.indexOf('Create least-privilege publication token'));

    assert.equal((boundaryStep.match(/status --porcelain=v1 -z --untracked-files=all --no-renames/g) || []).length, 2);
    assert.match(boundaryStep, /docs\/runbooks\/update-face-liveness-sdk\.md/);
    assert.match(boundaryStep, /apps\/learning-platform\/azure-ai-vision-face-ui\/\*/);
    assert.match(boundaryStep, /scripts\/face-sdk-vendor\.json/);
    assert.doesNotMatch(boundaryStep, /(?:package(?:-lock)?\.json|\.npmrc|\.github\/dependabot\.yml)/);

    assert.match(commitStep, /git -C backend add -- docs\/runbooks\/update-face-liveness-sdk\.md/);
    assert.match(commitStep, /git -C sistemas add --[\s\S]*apps\/learning-platform\/azure-ai-vision-face-ui[\s\S]*scripts\/face-sdk-vendor\.json/);
    assert.doesNotMatch(commitStep, /git\s+(?:-C\s+\S+\s+)?add\s+(?:-A|\.|--all)\b/);
    assert.doesNotMatch(commitStep, /(?:package(?:-lock)?\.json|\.npmrc|\.github\/dependabot\.yml)/);

    assert.match(publicationStep, /git -C "\$repository" diff --quiet/);
    assert.match(publicationStep, /assert_main_tips_unchanged/);
    assert.match(publicationStep, /git -C "\$repository" rev-parse main/);
    assert.match(publicationStep, /git -C "\$repository" ls-remote "\$remote" refs\/heads\/main/);
    assert.match(publicationStep, /main branch advanced while the proposal was prepared/);
    assert.match(publicationStep, /git -C "\$repository" merge-base main "\$remote_ref"/);
    assert.match(publicationStep, /git -C "\$repository" diff --no-renames --name-only -z/);
    assert.match(publicationStep, /"\$merge_base" "\$remote_ref" > "\$remote_changes_file"/);
    assert.match(publicationStep, /open_branches="\$\(/);
    assert.match(publicationStep, /remote_changes_file="\$\(mktemp/);
    assert.match(publicationStep, /mapfile -d '' -t remote_changes < "\$remote_changes_file"/);
    assert.match(publicationStep, /closed_record="\$\(find_closed_pull_request/);
    assert.doesNotMatch(publicationStep, /done < <\(\s*gh api/);
    assert.doesNotMatch(publicationStep, /mapfile -d '' -t remote_changes < <\(/);
    assert.match(publicationStep, /Existing \$repository proposal branch contains an unexpected path/);
    assert.match(publicationStep, /"\$remote_ref" "\$PROPOSED_BRANCH" -- "\$\{comparison_paths\[@\]\}"/);
    assert.match(publicationStep, /-F draft=true/);
    assert.match(publicationStep, /state=open/);
    assert.match(publicationStep, /state=closed/);
    assert.match(publicationStep, /\.head\.repo\.full_name ==/);
    assert.match(publicationStep, /An earlier Face SDK proposal remains open/);
    assert.match(publicationStep, /Both matching proposals are already ready for review and retain the validated candidates/);
    assert.match(publicationStep, /companion proposals have inconsistent draft states/);
    assert.match(publicationStep, /was closed or merged; it will not be recreated/);
    assert.match(publicationStep, /Companion Sistemas proposal/);
    assert.match(publicationStep, /Companion Backend proposal/);
    assert.match(publicationStep, /<!-- face-sdk-automation:start -->/);
    assert.match(publicationStep, /<!-- face-sdk-automation:end -->/);
    assert.match(publicationStep, /current_body/);
    assert.doesNotMatch(publicationStep, /body=\$backend_body|body=\$sistemas_body/);
    assert.match(publicationStep, /Private registry\/API credentials were confined/);
    assert.match(publicationStep, /temporary packages were cleaned on success or failure/);
    assert.match(publicationStep, /excluded from production deployment, preview creation\/closure, Azure\/OIDC credentials, and auto-merge/);
    assert.match(publicationStep, /Merge and verify the Sistemas proposal first/);
    assert.doesNotMatch(publicationStep, /\b(?:gh\s+pr\s+(?:merge|close|ready|review)|git\s+push\s+--(?:force|delete)|git\s+branch\s+-[dD]|--auto)\b/);
    assert.doesNotMatch(publicationStep, /gh api --method DELETE/);

    const prepareIndex = workflow.indexOf('Download and inspect the private packages');
    const backendValidationIndex = workflow.indexOf('Validate the complete Backend candidate');
    const sistemasValidationIndex = workflow.indexOf('Validate the credential-free Sistemas candidate');
    const cleanupIndex = workflow.indexOf('Remove every temporary private package');
    const publicationTokenIndex = workflow.indexOf('Create least-privilege publication token');
    const pushIndex = workflow.indexOf('git -C "$repository" push');
    assert.ok(prepareIndex < backendValidationIndex);
    assert.ok(prepareIndex < sistemasValidationIndex);
    assert.ok(backendValidationIndex < cleanupIndex);
    assert.ok(sistemasValidationIndex < cleanupIndex);
    assert.ok(cleanupIndex < publicationTokenIndex);
    assert.ok(publicationTokenIndex < pushIndex);
});

test('preparer protects dependency configuration and invokes Sistemas without credentials', () => {
    const preparer = readRepositoryFile('scripts/prepare-face-sdk-update.mjs');
    const npmConfiguration = readRepositoryFile('.npmrc');

    assert.match(
        preparer,
        /const PROTECTED_BACKEND_FILES = \['package\.json', 'package-lock\.json', '\.npmrc'\]/,
    );
    assert.match(preparer, /const protectedBefore = await fileDigests\(resolvedBackendRoot, PROTECTED_BACKEND_FILES\)/);
    assert.match(preparer, /const protectedAfter = await fileDigests\(resolvedBackendRoot, PROTECTED_BACKEND_FILES\)/);
    assert.match(preparer, /assertDigestsUnchanged\(protectedBefore, protectedAfter\)/);
    assert.match(
        preparer,
        /const SISTEMAS_MANIFEST_RELATIVE_PATH = 'scripts\/face-sdk-vendor\.json'/,
    );
    assert.match(preparer, /Backend and Sistemas record different Face SDK versions/);
    assert.match(preparer, /endpointUrl\.protocol !== 'https:'/);
    assert.match(preparer, /cognitiveservices\\\.azure\\\.com/);
    assert.match(preparer, /'scripts', 'update-face-sdk-vendor\.mjs'/);
    assert.match(
        preparer,
        /sistemasUpdater,[\s\S]*'--package-root', installRoot,[\s\S]*'--target-version', targetVersion,[\s\S]*'--repository-root', resolvedSistemasRoot/,
    );
    assert.match(
        preparer,
        /\{ cwd: resolvedSistemasRoot, env: sanitizeEnvironment\(environment\) \}/,
    );
    assert.match(preparer, /'--ignore-scripts'/);
    assert.match(preparer, /'--package-lock=false'/);
    assert.match(preparer, /'--no-save'/);
    assert.match(preparer, /'--loglevel=silent'/);
    assert.doesNotMatch(preparer, /AZURE_FACE_API_NPM_TEMPORARY_TOKEN/);

    assert.match(npmConfiguration, /_password=\$\{AZURE_AI_VISION_NPM_TOKEN_BASE64\}/);
    assert.doesNotMatch(npmConfiguration, /_password=(?!\$\{AZURE_AI_VISION_NPM_TOKEN_BASE64\})\S+/);
});

test('Face SDK proposal pull requests receive complete secret-free Backend validation', () => {
    const workflow = readRepositoryFile(validationWorkflowPath);
    const eventBlock = yamlMappingBlock(workflow, 0, 'on');
    const jobsBlock = yamlMappingBlock(workflow, 0, 'jobs');
    const validationJob = yamlMappingBlock(workflow, 2, 'validate');

    assert.deepEqual(yamlDirectKeys(eventBlock, 2), ['pull_request']);
    assert.deepEqual(yamlDirectKeys(jobsBlock, 2), ['validate']);
    assert.match(validationJob, /github\.event\.pull_request\.head\.repo\.full_name == github\.repository/);
    assert.match(validationJob, /startsWith\(github\.head_ref, 'chore\/update-face-liveness-sdk-'\)/);
    assert.match(workflow, /^permissions:\r?\n  contents: read\r?\n\r?\njobs:/m);
    assert.equal((workflow.match(/^permissions:/gm) || []).length, 1);
    assert.equal((workflow.match(/^[ \t]+permissions:/gm) || []).length, 0);
    assert.match(validationJob, /^          persist-credentials: false$/m);
    assert.match(validationJob, /^          package-manager-cache: false$/m);
    assert.match(validationJob, /^      - name: Install locked dependencies\n        run: npm ci$/m);

    for (const command of completeBackendCommands) {
        assert.ok(validationJob.split('\n').some((line) => line.trim() === command));
    }
    assert.deepEqual(
        referencedActions(workflow).map((action) => action.slice(0, action.lastIndexOf('@'))),
        ['actions/checkout', 'actions/setup-node'],
    );

    for (const forbidden of [
        /\bsecrets\b/,
        /\bvars\b/,
        /\bpull_request_target\b/,
        /\bworkflow_dispatch\b/,
        /\bworkflow_call\b/,
        /\bAzure\//i,
        /actions\/(?:upload|download)-artifact@/,
        /^\s*environment:/m,
        /^\s*id-token:/m,
        /^\s*needs:/m,
        /^\s*uses:\s*\.\/\.github\/workflows\//m,
        /^\s*run: npm start$/m,
        /\b(?:gh\s+pr|git\s+push|auto-merge)\b/,
    ]) {
        assert.doesNotMatch(workflow, forbidden);
    }
});

test('setup-only Backend files cannot trigger production deployment', () => {
    const workflow = readRepositoryFile(deploymentWorkflowPath);
    const pushBlock = yamlMappingBlock(yamlMappingBlock(workflow, 0, 'on'), 2, 'push');
    const ignoredPaths = Array.from(
        pushBlock.matchAll(/^      - '([^']+)'$/gm),
        ([, ignoredPath]) => ignoredPath,
    );

    for (const requiredPath of [
        updateWorkflowPath,
        validationWorkflowPath,
        deploymentWorkflowPath,
        'docs/runbooks/update-face-liveness-sdk.md',
        'scripts/prepare-face-sdk-update.mjs',
        'test/face-sdk-acquisition.test.js',
        'test/face-sdk-workflows.test.js',
    ]) {
        assert.ok(ignoredPaths.includes(requiredPath), `${requiredPath} must be deployment-neutral`);
    }
    for (const runtimePath of [
        'app.js',
        'server.js',
        'package.json',
        'package-lock.json',
        '.npmrc',
        'domains/**',
        'integrations/**',
        'shared/**',
    ]) {
        assert.ok(!ignoredPaths.includes(runtimePath), `${runtimePath} must retain the production deployment path`);
    }
    assert.doesNotMatch(workflow, /\bpull_request(?:_target)?\b/);
});
