'use strict';

const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const repositoryRoot = path.resolve(__dirname, '..');

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

function workflowActionReferences(source) {
    return Array.from(
        source.matchAll(/^\s*(?:-\s*)?uses:\s*(\S+)(?:[^\S\r\n]+# ([^\r\n]+))?$/gm),
        ([, reference, version]) => ({
            action: reference.slice(0, reference.lastIndexOf('@')),
            ref: reference.slice(reference.lastIndexOf('@') + 1),
            version,
        }),
    );
}

function dependabotUpdateBlocks(configuration) {
    const lines = configuration.replace(/\r\n/g, '\n').split('\n');
    const starts = lines.flatMap((line, index) => (
        /^  - package-ecosystem: "[^"]+"$/.test(line) ? [index] : []
    ));

    return starts.map((start, index) => {
        const end = starts[index + 1] ?? lines.length;
        const name = lines[start].match(/^  - package-ecosystem: "([^"]+)"$/)[1];
        return { name, source: lines.slice(start, end).join('\n') };
    });
}

test('Dependabot proposes grouped non-major npm updates without changing security updates', () => {
    const configuration = readRepositoryFile('.github/dependabot.yml');
    const updateBlocks = dependabotUpdateBlocks(configuration);
    const actionsBlock = updateBlocks.find(({ name }) => name === 'github-actions').source;
    const npmBlock = updateBlocks.find(({ name }) => name === 'npm').source;

    assert.match(configuration, /^version: 2\r?\nupdates:\r?\n/);
    assert.deepEqual(updateBlocks.map(({ name }) => name), ['github-actions', 'npm']);
    assert.match(actionsBlock, /^    directory: "\/"$/m);
    assert.match(actionsBlock, /^      interval: "weekly"$/m);
    assert.doesNotMatch(actionsBlock, /^\s*groups:/m);
    assert.match(npmBlock, /^    directory: "\/"$/m);
    assert.match(npmBlock, /^      interval: "weekly"$/m);
    assert.equal((configuration.match(/^\s*groups:/gm) || []).length, 1);
    assert.match(
        npmBlock,
        /npm-minor-and-patch:[\s\S]*?applies-to: "version-updates"[\s\S]*?patterns:[\s\S]*?- "\*"[\s\S]*?update-types:[\s\S]*?- "minor"[\s\S]*?- "patch"/,
    );

    for (const forbidden of [
        /^registries:/m,
        /^\s*target-branch:/m,
        /^\s*open-pull-requests-limit:/m,
        /applies-to: "security-updates"/,
    ]) {
        assert.doesNotMatch(configuration, forbidden);
    }
});

test('repository workflows pin reviewed actions and bound each job with explicit authority', () => {
    const expectedActions = new Map([
        ['main_plataforma-backend-v3.yml', [
            'actions/checkout',
            'actions/setup-node',
            'actions/upload-artifact',
            'actions/download-artifact',
            'azure/login',
            'azure/webapps-deploy',
        ]],
        ['dependabot-validation.yml', ['actions/checkout', 'actions/setup-node']],
        ['face-sdk-version-check.yml', ['actions/checkout', 'actions/setup-node']],
    ]);
    const workflowNames = fs.readdirSync(path.join(repositoryRoot, '.github/workflows'))
        .filter((name) => /\.ya?ml$/u.test(name)).sort();
    assert.deepEqual(workflowNames, [...expectedActions.keys()].sort());

    for (const [name, actions] of expectedActions) {
        const workflow = readRepositoryFile(`.github/workflows/${name}`);
        const references = workflowActionReferences(workflow);
        assert.match(workflow, /^permissions: \{\}$/m);
        assert.equal((workflow.match(/^permissions:/gm) || []).length, 1);
        assert.deepEqual(references.map(({ action }) => action), actions);
        for (const { action, ref, version } of references) {
            assert.match(ref, /^[0-9a-f]{40}$/, `${action} must use a full commit SHA`);
            assert.match(version || '', /^v\d+\.\d+\.\d+$/, `${action} must retain its release comment`);
        }

        const jobs = yamlDirectKeys(yamlMappingBlock(workflow, 0, 'jobs'), 2);
        for (const jobName of jobs) {
            const job = yamlMappingBlock(workflow, 2, jobName);
            const timeout = job.match(/^    timeout-minutes: (\d+)$/m);
            assert.ok(timeout, `${name}/${jobName} must have a timeout`);
            assert.ok(Number(timeout[1]) > 0 && Number(timeout[1]) <= 20);
            yamlMappingBlock(job, 4, 'permissions');
        }
        assert.equal((workflow.match(/^          persist-credentials: false$/gm) || []).length, 1);
        assert.equal((workflow.match(/^          node-version: '24\.x'$/gm) || []).length, 1);
        assert.equal((workflow.match(/^          package-manager-cache: false$/gm) || []).length, 1);
        assert.doesNotMatch(workflow, /^\s*(?:cache|cache-dependency-path):/m);
        assert.doesNotMatch(workflow, /\bpull_request_target\b|\bwrite-all\b/);
    }
});

test('production deployment serializes complete runs and installs the unchanged lock deterministically', () => {
    const workflow = readRepositoryFile('.github/workflows/main_plataforma-backend-v3.yml');
    const eventBlock = yamlMappingBlock(workflow, 0, 'on');
    const concurrency = yamlMappingBlock(workflow, 0, 'concurrency');
    const buildJob = yamlMappingBlock(workflow, 2, 'build');
    const deployJob = yamlMappingBlock(workflow, 2, 'deploy');
    const buildPermissions = yamlMappingBlock(buildJob, 4, 'permissions');
    const deployPermissions = yamlMappingBlock(deployJob, 4, 'permissions');

    assert.deepEqual(yamlDirectKeys(eventBlock, 2), ['push', 'workflow_dispatch']);
    assert.match(eventBlock, /^      - main$/m);
    assert.deepEqual(
        Array.from(eventBlock.matchAll(/^      - '([^']+)'$/gm), ([, value]) => value),
        [
            '**/*.md', '.agents/**', '.github/dependabot.yml',
            '.github/workflows/face-sdk-version-check.yml',
            '.github/workflows/main_plataforma-backend-v3.yml',
            '.gitignore', 'docs/**', 'test/**',
        ],
    );
    assert.deepEqual(yamlDirectKeys(concurrency, 2), ['group', 'cancel-in-progress']);
    assert.match(concurrency, /^  group: backend-production-deployment$/m);
    assert.match(concurrency, /^  cancel-in-progress: false$/m);
    assert.equal((workflow.match(/^[ \t]*concurrency:/gm) || []).length, 1);
    assert.deepEqual(yamlDirectKeys(buildPermissions, 6), ['contents']);
    assert.match(buildPermissions, /^      contents: read$/m);
    assert.deepEqual(yamlDirectKeys(deployPermissions, 6), ['id-token']);
    assert.match(deployPermissions, /^      id-token: write(?: #.*)?$/m);
    assert.match(buildJob, /^    runs-on: windows-latest$/m);
    assert.match(deployJob, /^    runs-on: ubuntu-latest$/m);
    assert.match(deployJob, /^    needs: build$/m);
    assert.match(deployJob, /^      name: 'Production'$/m);
    assert.match(deployJob, /^          app-name: 'Plataforma-Backend-v3'$/m);
    assert.match(deployJob, /^          slot-name: 'Production'$/m);
    assert.deepEqual(
        Array.from(buildJob.matchAll(/^        run: (.+)$/gm), ([, command]) => command),
        ['npm ci', 'npm run build --if-present', 'npm test'],
    );
    assert.doesNotMatch(buildJob, /^        run: [|>]|\bnpm install\b/m);
    assert.doesNotMatch(buildJob, /\bsecrets\b|\bid-token\b|\bAzure\//i);
    assert.match(buildJob, /^          name: node-app$/m);
    assert.match(buildJob, /^          path: \.$/m);
    assert.match(deployJob, /^          name: node-app$/m);
    assert.match(deployJob, /^          package: \.$/m);

    const packageManifest = JSON.parse(readRepositoryFile('package.json'));
    const lock = JSON.parse(readRepositoryFile('package-lock.json'));
    assert.equal(packageManifest.engines.node, '24.x');
    assert.equal(lock.packages[''].engines.node, '24.x');
    assert.equal(packageManifest.scripts.test, 'node --require ./test/deny-production-network.js --test');
});

test('Dependabot pull requests are validated without a production deployment path', () => {
    const validationWorkflow = readRepositoryFile('.github/workflows/dependabot-validation.yml');
    const deploymentWorkflow = readRepositoryFile(
        '.github/workflows/main_plataforma-backend-v3.yml',
    );
    const eventBlock = yamlMappingBlock(validationWorkflow, 0, 'on');
    const jobsBlock = yamlMappingBlock(validationWorkflow, 0, 'jobs');
    const validationJob = yamlMappingBlock(validationWorkflow, 2, 'validate');
    const deploymentJob = yamlMappingBlock(deploymentWorkflow, 2, 'deploy');
    const validationPermissions = yamlMappingBlock(validationJob, 4, 'permissions');
    const concurrency = yamlMappingBlock(validationWorkflow, 0, 'concurrency');

    assert.deepEqual(yamlDirectKeys(eventBlock, 2), ['pull_request']);
    assert.deepEqual(yamlDirectKeys(jobsBlock, 2), ['validate']);
    assert.doesNotMatch(validationWorkflow, /\bpull_request_target\b/);
    assert.doesNotMatch(validationWorkflow, /\bworkflow_dispatch\b/);
    assert.match(validationWorkflow, /^permissions: \{\}$/m);
    assert.equal((validationWorkflow.match(/^permissions:/gm) || []).length, 1);
    assert.equal((validationWorkflow.match(/^[ \t]+permissions:/gm) || []).length, 1);
    assert.deepEqual(yamlDirectKeys(validationPermissions, 6), ['contents']);
    assert.match(validationPermissions, /^      contents: read$/m);
    assert.match(concurrency, /^  group: dependabot-validation-\$\{\{ github\.event\.pull_request\.number \}\}$/m);
    assert.match(concurrency, /^  cancel-in-progress: true$/m);
    assert.match(
        validationJob,
        /^\s*if: github\.event\.pull_request\.user\.login == 'dependabot\[bot\]'$/m,
    );
    for (const requiredLine of [
        'persist-credentials: false',
        "node-version: '24.x'",
        'package-manager-cache: false',
        'run: npm ci',
        'shell: bash',
        'run: npm test',
    ]) {
        assert.ok(validationJob.split('\n').some((line) => line.trim() === requiredLine));
    }

    for (const command of [
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
    ]) {
        assert.ok(
            validationJob.split('\n').some((line) => line.trim() === command),
            `${command} must run for bot PRs`,
        );
    }

    const validationActions = workflowActionReferences(validationJob);
    assert.deepEqual(
        validationActions.map(({ action }) => action),
        ['actions/checkout', 'actions/setup-node'],
    );

    for (const forbidden of [
        /\bsecrets\b/,
        /\bwrite-all\b/,
        /\bAzure\//i,
        /actions\/(?:upload|download)-artifact@/,
        /^\s*environment:/m,
        /^\s*id-token:/m,
        /^\s*needs:/m,
        /^\s*uses:\s*\.\/\.github\/workflows\//m,
        /^\s*run: npm start$/m,
    ]) {
        assert.doesNotMatch(validationWorkflow, forbidden);
    }

    assert.doesNotMatch(deploymentWorkflow, /\bpull_request(?:_target)?\b/);
    assert.match(deploymentJob, /^\s*environment:/m);
    assert.match(deploymentJob, /^\s*id-token: write/m);
    assert.match(
        deploymentJob,
        /^[ \t]*(?:-[ \t]*)?uses: azure\/login@[0-9a-f]{40} # v\d+\.\d+\.\d+$/im,
    );
    assert.match(
        deploymentJob,
        /^[ \t]*(?:-[ \t]*)?uses: azure\/webapps-deploy@[0-9a-f]{40} # v\d+\.\d+\.\d+$/im,
    );
});

test('Face SDK checker reads public metadata and only opens a review issue', () => {
    const workflowPath = '.github/workflows/face-sdk-version-check.yml';
    const workflow = readRepositoryFile(workflowPath);
    const deploymentWorkflow = readRepositoryFile(
        '.github/workflows/main_plataforma-backend-v3.yml',
    );
    const eventBlock = yamlMappingBlock(workflow, 0, 'on');
    const jobsBlock = yamlMappingBlock(workflow, 0, 'jobs');
    const checkJob = yamlMappingBlock(workflow, 2, 'check');
    const notifyJob = yamlMappingBlock(workflow, 2, 'notify');
    const checkPermissions = yamlMappingBlock(checkJob, 4, 'permissions');
    const notifyPermissions = yamlMappingBlock(notifyJob, 4, 'permissions');

    assert.deepEqual(yamlDirectKeys(eventBlock, 2), ['schedule', 'workflow_dispatch']);
    assert.match(eventBlock, /^    - cron: '23 12 \* \* 2'$/m);
    assert.deepEqual(yamlDirectKeys(jobsBlock, 2), ['check', 'notify']);
    assert.match(workflow, /^permissions: \{\}$/m);
    assert.deepEqual(yamlDirectKeys(checkPermissions, 6), ['contents']);
    assert.match(checkPermissions, /^      contents: read$/m);
    assert.deepEqual(yamlDirectKeys(notifyPermissions, 6), ['issues']);
    assert.match(notifyPermissions, /^      issues: write$/m);

    assert.match(checkJob, /github\.repository == 'IvyRoom\/backend'/);
    assert.match(checkJob, /github\.ref == 'refs\/heads\/main'/);
    assert.match(checkJob, /github\.event_name == 'schedule'/);
    assert.match(checkJob, /github\.event_name == 'workflow_dispatch'/);
    assert.match(checkJob, /^          persist-credentials: false$/m);
    assert.match(checkJob, /^          package-manager-cache: false$/m);
    assert.match(
        checkJob,
        /npm view '@azure\/ai-vision-face-ui@latest' version --userconfig=\/dev\/null --registry=https:\/\/registry\.npmjs\.org\/ --ignore-scripts --silent/,
    );
    assert.ok(checkJob.indexOf('cd "$RUNNER_TEMP"') < checkJob.indexOf('npm view'));
    assert.match(checkJob, /latest_major > current_major/);
    assert.match(checkJob, /Public latest version .* is older than vendored version/);
    assert.match(notifyJob, /^    needs: check$/m);
    assert.match(notifyJob, /^    if: needs\.check\.outputs\.update_available == 'true'$/m);
    assert.match(notifyJob, /^          GH_TOKEN: \$\{\{ github\.token \}\}$/m);
    assert.match(notifyJob, /^          ISSUE_ASSIGNEE: IvyRoom$/m);
    assert.doesNotMatch(notifyJob, /github\.actor/);
    assert.match(
        notifyJob,
        /gh api --paginate --method GET[\s\S]*repos\/\$GITHUB_REPOSITORY\/issues[\s\S]*-f state=all/,
    );
    assert.doesNotMatch(notifyJob, /gh issue list/);
    assert.match(notifyJob, /gh issue create[\s\S]*--assignee "\$ISSUE_ASSIGNEE"/);

    const actions = workflowActionReferences(workflow).map(({ action }) => action);
    assert.deepEqual(actions, ['actions/checkout', 'actions/setup-node']);

    for (const forbidden of [
        /\bsecrets\b/,
        /\bvars\b/,
        /^\s*pull_request(?:_target)?:/m,
        /^\s*environment:/m,
        /^\s*id-token:/m,
        /^\s*(?:-\s*)?uses:\s*azure\//im,
        /actions\/create-github-app-token@/,
        /\bnpm (?:ci|install|pack)\b/,
        /\bgit push\b/,
        /\bgh pr\b/,
        /^\s*repository:\s*IvyRoom\/sistemas\s*$/im,
        /\bdeploy(?:ment)?\b/i,
    ]) {
        assert.doesNotMatch(workflow, forbidden);
    }

    assert.match(
        deploymentWorkflow,
        /^      - '\.github\/workflows\/face-sdk-version-check\.yml'$/m,
    );
    for (const runtimePath of ['app.js', 'server.js', 'package.json', 'package-lock.json']) {
        assert.doesNotMatch(
            deploymentWorkflow,
            new RegExp("^      - '" + runtimePath.replace('.', '\\.') + "'$", 'm'),
        );
    }
});
