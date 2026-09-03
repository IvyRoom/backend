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

test('Dependabot pull requests are validated without a production deployment path', () => {
    const validationWorkflow = readRepositoryFile('.github/workflows/dependabot-validation.yml');
    const deploymentWorkflow = readRepositoryFile(
        '.github/workflows/main_plataforma-backend-v3.yml',
    );
    const eventBlock = yamlMappingBlock(validationWorkflow, 0, 'on');
    const jobsBlock = yamlMappingBlock(validationWorkflow, 0, 'jobs');
    const validationJob = yamlMappingBlock(validationWorkflow, 2, 'validate');
    const deploymentJob = yamlMappingBlock(deploymentWorkflow, 2, 'deploy');

    assert.deepEqual(yamlDirectKeys(eventBlock, 2), ['pull_request']);
    assert.deepEqual(yamlDirectKeys(jobsBlock, 2), ['validate']);
    assert.doesNotMatch(validationWorkflow, /\bpull_request_target\b/);
    assert.doesNotMatch(validationWorkflow, /\bworkflow_dispatch\b/);
    assert.match(
        validationWorkflow,
        /^permissions:\r?\n  contents: read\r?\n\r?\njobs:/m,
    );
    assert.equal((validationWorkflow.match(/^permissions:/gm) || []).length, 1);
    assert.equal((validationWorkflow.match(/^[ \t]+permissions:/gm) || []).length, 0);
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

    const validationActions = Array.from(
        validationJob.matchAll(/^\s*(?:-\s*)?uses:\s*(\S+)$/gm),
        ([, action]) => action,
    );
    assert.deepEqual(
        validationActions.map((action) => action.slice(0, action.lastIndexOf('@'))),
        ['actions/checkout', 'actions/setup-node'],
    );
    for (const action of validationActions) {
        assert.match(action, /^(?:actions\/checkout|actions\/setup-node)@(?:v\d+(?:\.\d+\.\d+)?|[0-9a-f]{40})$/);
    }

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
        /^[ \t]*(?:-[ \t]*)?uses: azure\/login@(?:v\d+(?:\.\d+\.\d+)?|[0-9a-f]{40})$/im,
    );
    assert.match(
        deploymentJob,
        /^[ \t]*(?:-[ \t]*)?uses: azure\/webapps-deploy@(?:v\d+(?:\.\d+\.\d+)?|[0-9a-f]{40})$/im,
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
    assert.match(
        notifyJob,
        /gh api --paginate --method GET[\s\S]*repos\/\$GITHUB_REPOSITORY\/issues[\s\S]*-f state=all/,
    );
    assert.doesNotMatch(notifyJob, /gh issue list/);
    assert.match(notifyJob, /gh issue create[\s\S]*--assignee "\$ISSUE_ASSIGNEE"/);

    const actions = Array.from(
        workflow.matchAll(/^\s*(?:-\s*)?uses:\s*(\S+)$/gm),
        ([, action]) => action.slice(0, action.lastIndexOf('@')),
    );
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
