'use strict';

process.env.NODE_ENV = 'test';

const test = require('node:test');
const assert = require('node:assert/strict');
const {
    TEST_NOW_MS,
    createDeferred,
    createTestApp,
    startLoopback,
} = require('./app-test-support');

const USER_PREFIX = '/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be';
const SEND_MAIL_PATH = `${USER_PREFIX}/sendMail`;
const RECOMMENDATIONS_TABLE = `${USER_PREFIX}/drive/items/01OSXVECRAQXJDB7TBYFGKA5YQJXO3YAOS/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}`;
const PLATFORM_TABLE = `${USER_PREFIX}/drive/items/01OSXVECSBYCZNYGEWFFDLEOZ36WI2PDWO/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}`;
const CLIENTS_TABLE = `${USER_PREFIX}/drive/items/01OSXVECQNNRY4S7VCKBF2SOETFSLESSLH/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}`;
const RECOMMENDATIONS_ROWS = `${RECOMMENDATIONS_TABLE}/rows`;
const PLATFORM_ROWS = `${PLATFORM_TABLE}/rows`;
const CLIENTS_ROWS = `${CLIENTS_TABLE}/rows`;
const BRAZIL_NOW_SERIAL = 46402 + (5 / 24);
const ACCESS_DEADLINE_SERIAL = 46462;
const RETRY_DELAYS = [500, 1_000, 1_500, 2_000];

function workbookRow(cells) {
    return { values: [cells] };
}

function workbookRows(rows) {
    return { value: rows.map(workbookRow) };
}

function sleepDelays(ledger) {
    return ledger
        .filter(({ type }) => type === 'sleep')
        .map(({ delay }) => delay);
}

async function postJson(origin, route, body) {
    return fetch(`${origin}${route}`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(body),
    });
}

async function assertJsonResponse(response, status, body) {
    assert.equal(response.status, status);
    assert.equal(response.headers.get('content-type'), 'application/json; charset=utf-8');
    assert.deepEqual(await response.json(), body);
}

function validRecommendationPayload(overrides = {}) {
    return {
        recommenderFullName: 'Alice & Silva',
        benefitedCompany: 'Acme <Holding>',
        recommendedCompany: 'Beta Labs',
        recommendedProfessional: 'Bruno Souza',
        recommendedWhatsapp: '+55 41 99999-0000',
        ...overrides,
    };
}

function recommendationCells(overrides = {}) {
    const cells = new Array(13).fill('-');
    cells[0] = 'Acme <Holding>';
    cells[1] = 'Alice & Silva';
    cells[2] = '-';
    cells[3] = 'alice@example.test';

    for (const [index, value] of Object.entries(overrides)) cells[Number(index)] = value;
    return cells;
}

function expectedRecommendationCells(payload, copiedCells = null) {
    const cells = new Array(13).fill(null);
    cells[4] = BRAZIL_NOW_SERIAL;
    cells[5] = payload.recommendedCompany.trim();
    cells[6] = payload.recommendedProfessional.trim();
    cells[7] = payload.recommendedWhatsapp.trim();
    cells[8] = '1. REALIZAR CONTATO INICIAL';
    cells[9] = 'A INICIAR';
    cells[10] = BRAZIL_NOW_SERIAL;
    cells[11] = BRAZIL_NOW_SERIAL;

    if (copiedCells) {
        cells[0] = copiedCells[0];
        cells[1] = copiedCells[1];
        cells[3] = copiedCells[3];
        cells[12] = '-';
    }

    return cells;
}

function quoteMailBody(payload) {
    return {
        message: {
            subject: 'Machado - Nova Solicitação de Orçamento',
            body: {
                contentType: 'HTML',
                content: `<p><b>Dados do Solicitante:</b></p><p>${payload.Solicitante_NomeCompleto}</p><p>${payload.Solicitante_Email}</p><p>${payload.Solicitante_Telefone}</p><p>${payload.Solicitante_Cargo}</p><p><b>Dados da Empresa:</b></p><p>${payload.Solicitante_NomeEmpresa}</p><p>${payload.Solicitante_CNPJ}</p><p>${payload.Solicitante_NúmerodeParticipantes}</p><p>${payload.Solicitante_Observações}</p><p><img width="500" height="auto" src="https://plataforma-backend-v3.azurewebsites.net/img/ASSINATURA_E-MAIL.jpg"/></p>`,
            },
            toRecipients: [{ emailAddress: { address: 'contato@machadogestao.com' } }],
        },
    };
}

function conectaMailBodies(payload, recommenderCells) {
    const signature = '<p><img width="600" height="auto" src="https://plataforma-backend-v3.azurewebsites.net/img/ASSINATURA_E-MAIL.jpg"/></p>';
    const recommenderEmail = String(recommenderCells[3]).trim();
    const recommenderFirstNameCell = String(recommenderCells[2]).trim();
    const recommenderFirstName = recommenderFirstNameCell && recommenderFirstNameCell !== '-'
        ? recommenderFirstNameCell
        : String(recommenderCells[1]).trim().split(/\s+/)[0];
    const escapeHtml = (value) => String(value == null ? '' : value).replace(
        /[&<>"']/g,
        (character) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[character]),
    );
    const internalContent = `<p><b>Dados do Recomendante:</b></p><p>Nome Completo: ${escapeHtml(recommenderCells[1])}</p><p>E-mail: ${escapeHtml(recommenderEmail)}</p><p>Empresa Beneficiada: ${escapeHtml(recommenderCells[0])}</p><p><b>Dados da Recomendação:</b></p><p>Empresa Recomendada: ${escapeHtml(payload.recommendedCompany.trim())}</p><p>Profissional Contatado: ${escapeHtml(payload.recommendedProfessional.trim())}</p><p>WhatsApp do Profissional: ${escapeHtml(payload.recommendedWhatsapp.trim())}</p>${signature}`;
    const confirmationContent = `<p>Olá ${escapeHtml(recommenderFirstName)},</p><p>Recebemos sua recomendação da Machado para a empresa <b>${escapeHtml(payload.recommendedCompany.trim())}</b>. Obrigado pela confiança.</p><p>Logo entraremos em contato com ${escapeHtml(payload.recommendedProfessional.trim())}. Assim que houver atualizações relevantes, sinalizaremos a você.</p><p>Atenciosamente,</p>${signature}`;

    return [
        {
            message: {
                subject: 'Machado Conecta - Nova Recomendação Recebida',
                body: { contentType: 'HTML', content: internalContent },
                toRecipients: [{ emailAddress: { address: 'contato@machadogestao.com' } }],
            },
        },
        {
            message: {
                subject: 'Machado Conecta - Recomendação Registrada',
                body: { contentType: 'HTML', content: confirmationContent },
                toRecipients: [{ emailAddress: { address: recommenderEmail } }],
            },
        },
    ];
}

function validIntakePayload() {
    return {
        company: {
            legalName: 'Empresa <Legal> & Cia',
            cnpj: '12.345.678/0001-90',
            address: {
                postalCode: '80000-000',
                street: 'Rua <Sede>',
                number: '10 & 12',
                complement: "Sala 'A'",
                neighborhood: 'Centro > Norte',
                city: 'Curitiba & Região',
                state: 'PR',
            },
        },
        shippingAddress: {
            postalCode: '81000-000',
            street: 'Rua de Entrega & Filhos',
            number: '007',
            complement: '',
            neighborhood: 'Bairro <Sul>',
            city: 'Curitiba',
            state: 'PR',
            useCompanyAddress: false,
        },
        legalRepresentative: {
            fullName: 'Lia <Legal>',
            cpf: '111.111.111-11',
            role: 'Representante & Sócia',
            areaCode: '41',
            whatsapp: '99999-1111',
            email: 'lia@example.test',
        },
        adminAssistant: {
            fullName: 'Caio & Costa',
            cpf: '222.222.222-22',
            role: 'Financeiro > Administrativo',
            areaCode: '41',
            whatsapp: '99999-2222',
            email: 'caio@example.test',
        },
        participants: [{
            fullName: 'Ana <Souza>',
            email: 'ana@example.test',
            cpf: '333.333.333-33',
            role: 'Diretora & Sócia',
            areaCode: '41',
            whatsapp: '99999-3333',
        }],
    };
}

function clone(value) {
    return JSON.parse(JSON.stringify(value));
}

function expectedPlatformCells(
    participant,
    {
        password = 100000000000,
        certificateId = 'FMG-0000-0000',
    } = {},
) {
    const cells = new Array(22).fill(null);
    cells[0] = participant.fullName;
    cells[2] = participant.email;
    cells[3] = password;
    cells[4] = 'Ativo';
    cells[5] = 'Não';
    cells[6] = ACCESS_DEADLINE_SERIAL;
    cells[8] = 0;
    for (let module = 10; module <= 19; module++) cells[module] = 0;
    cells[21] = certificateId;
    return cells;
}

function expectedClientCells(payload, participant) {
    const cells = new Array(13).fill(null);
    const shipping = payload.shippingAddress;
    cells[0] = payload.company.legalName;
    cells[3] = participant.fullName;
    cells[4] = participant.cpf;
    cells[5] = shipping.street;
    cells[6] = /^\d+$/.test(shipping.number) ? Number(shipping.number) : shipping.number;
    cells[7] = shipping.complement || '-';
    cells[8] = shipping.neighborhood;
    cells[9] = shipping.city;
    cells[10] = shipping.state;
    cells[12] = shipping.postalCode;
    return cells;
}

function intakeMailBody(payload) {
    const escapeHtml = (value) => String(value == null ? '' : value).replace(
        /[&<>"']/g,
        (character) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[character]),
    );
    const companyAddress = payload.company.address || {};
    const pessoaHtml = (label, person) => `<p><b>${label}</b></p><p>Nome Completo: ${escapeHtml(person.fullName)}</p><p>CPF: ${escapeHtml(person.cpf)}</p><p>Cargo: ${escapeHtml(person.role)}</p><p>DDD: ${escapeHtml(person.areaCode)}</p><p>WhatsApp: ${escapeHtml(person.whatsapp)}</p><p>E-mail: ${escapeHtml(person.email)}</p>`;
    const participantsHtml = payload.participants
        .map((participant, index) => `<p>${index + 1}. ${escapeHtml(participant.fullName)} — Cargo: ${escapeHtml(participant.role)} · DDD: ${escapeHtml(participant.areaCode)} · WhatsApp: ${escapeHtml(participant.whatsapp)}</p>`)
        .join('');
    const content = `<p>Um novo Formulário de Informações Iniciais foi preenchido.</p><p><b>Pessoa Jurídica Contratante</b></p><p>Razão Social: ${escapeHtml(payload.company.legalName)}</p><p>CNPJ: ${escapeHtml(payload.company.cnpj)}</p><p>CEP: ${escapeHtml(companyAddress.postalCode)}</p><p>Rua: ${escapeHtml(companyAddress.street)}</p><p>Número: ${escapeHtml(companyAddress.number)}</p><p>Complemento: ${escapeHtml(companyAddress.complement)}</p><p>Bairro: ${escapeHtml(companyAddress.neighborhood)}</p><p>Cidade: ${escapeHtml(companyAddress.city)}</p><p>Estado: ${escapeHtml(companyAddress.state)}</p>${pessoaHtml('Representante Jurídico', payload.legalRepresentative)}${pessoaHtml('Auxiliar Administrativo Financeiro', payload.adminAssistant)}<p><b>Participantes</b></p>${participantsHtml}<p><img width="500" height="auto" src="https://plataforma-backend-v3.azurewebsites.net/img/ASSINATURA_E-MAIL.jpg"/></p>`;

    return {
        message: {
            subject: 'Machado: novo Formulário de Informações Iniciais preenchido',
            body: { contentType: 'HTML', content },
            toRecipients: [{ emailAddress: { address: 'contato@machadogestao.com' } }],
        },
    };
}

function accessReleaseRows() {
    const recipients = Array.from({ length: 7 }, (_, index) => ({
        firstName: `Cliente ${index + 1}`,
        email: `cliente${index + 1}@example.test`,
        accessCode: `senha-${index + 1}`,
    }));
    const rows = Array.from({ length: 42 }, (_, index) => {
        const cells = new Array(22).fill(null);
        cells[0] = `Unselected ${index}`;
        return workbookRow(cells);
    });

    recipients.forEach((recipient, index) => {
        const cells = rows[index + 35].values[0];
        cells[1] = recipient.firstName;
        cells[2] = recipient.email;
        cells[3] = recipient.accessCode;
    });

    return { data: { value: rows }, recipients };
}

function accessReleaseMailBody({ firstName, email, accessCode }) {
    const content = [
        '',
        `                            <p>Bom dia ${firstName},</p>`,
        '                            <p>Escrevemos do suporte da Machado | Método Gerencial para Empresas. Tudo bem?</p>',
        '                            <p>Recentemente a Engefy contratou a nova versão de nossa Solução em Método Gerencial, para auxiliarmos no amadurecimento do Sistema de Gestão da empresa. E você foi um dos profissionais selecionados para participar do trabalho!</p>',
        '                            <p>A Solução possui duas grandes porções:</p>',
        '                            <p><b>• Formação em Método Gerencial:</b> acontece em nossa plataforma de ensino, de maneira online e assíncrona, durante 5 semanas. Esta é a etapa que estamos começando agora.</p>',
        '                            <p><b>• Encontros ao Vivo:</b> posteriormente, nosso fundador (Lucas Machado) irá até a Engefy para conduzir junto a vocês o choque de Gestão na empresa, durante 2 dias.</p>',
        '                            <p>Dito isto, compartilhamos as instruções de acesso à Formação:</p>',
        '                            <span><b>Link:</b> <a href="https://machadogestao.com/plataforma_v2/login">https://machadogestao.com/plataforma_v2/login</a><br></span>',
        `                            <span><b>Login:</b> ${email}<br></span>`,
        `                            <span><b>Senha:</b> ${accessCode}<br></span>`,
        '                            <p>*Suas credenciais de acesso são individuais e instransferíveis.</p>',
        '                            <p>**Nossa plataforma possui várias camadas de segurança e monitoramento. Por isto, o acesso deve ser realizado exclusivamente pelo navegador <b>Microsoft Edge</b>, via laptop ou desktop com <b>sistema Windows</b>. Computadores Apple/Mac são incompatíveis com nossos sistemas.</p>',
        '                            <p>Orientações Adicionais:</p>',
        '                            <p>• Sua caixa personalizada com materiais impressos (apostilas, cases, documentos auxiliares, etc.) já foi enviada à Engefy. Favor alinhar recebimento junto ao Luan Mannes.</p>',
        '                            <p>• A meta de início dos estudos será encaminhada pelo grupo do WhatsApp ainda hoje, logo após a reunião de kick-off. Importante: sugerimos que você tenha sua caixa de materiais impressos em mãos antes de iniciar os estudos.</p>',
        '                            <p>• Porém sugerimos também que você faça seu primeiro login, incluindo cadastramento no sistema de reconhecimento facial e familiarização inicial com a plataforma desde já.</p>',
        '                            <p>Em caso de dúvidas / dificuldades:</p>',
        '                            <p>• <b>Técnicas</b> (relacionadas ao acesso à plataforma ou eventuais bugs): sinalize para nós via inbox ao WhatsApp +55 41 99679 9092. Iremos auxiliá-lo(a) prontamente.</p>',
        '                            <p>• <b>Conceituais</b> (relacionadas à compreensão ou aplicação do Método Gerencial no dia a dia da Engefy): anote em seus materiais impressos de forma organizada e traga nos Encontros ao Vivo para discussão conjunta.</p>',
        '                            <p>Qualquer dúvida ou insegurança, sempre à disposição.</p>',
        '                            <p>Atenciosamente,</p>',
        '                            <p><img src="https://plataforma-backend-v3.azurewebsites.net/img/ASSINATURA_E-MAIL.jpg" width="600" /></p>',
        '                        ',
    ].join('\n');

    return {
        message: {
            subject: 'Machado | Método Gerencial para Empresas - Instruções de Acesso à Plataforma',
            body: { contentType: 'HTML', content },
            toRecipients: [{ emailAddress: { address: email } }],
        },
    };
}

test('quote request sends the exact unescaped internal mail and repeats without dedupe', async (t) => {
    const payload = {
        Solicitante_NomeCompleto: 'Ana <Silva>',
        Solicitante_Email: 'ana@example.test',
        Solicitante_Telefone: '+55 41 90000-0000',
        Solicitante_Cargo: 'Diretora & Sócia',
        Solicitante_NomeEmpresa: 'Empresa <Beta>',
        Solicitante_CNPJ: '12.345.678/0001-90',
        Solicitante_NúmerodeParticipantes: 17,
        Solicitante_Observações: 'Sem restrições & sem escape',
    };
    const { app, graphClient, ledger } = createTestApp();
    graphClient.enqueue('POST', SEND_MAIL_PATH, {}, {});
    const { origin } = await startLoopback(app, t);

    for (let repeat = 0; repeat < 2; repeat++) {
        const response = await postJson(origin, '/landingpage/solicitacaoorcamento', payload);
        await assertJsonResponse(response, 200, {});
    }

    assert.equal(graphClient.calls.length, 2);
    assert.ok(graphClient.calls.every(({ method, path }) => method === 'POST' && path === SEND_MAIL_PATH));
    assert.deepEqual(graphClient.calls.map(({ body }) => body), [
        quoteMailBody(payload),
        quoteMailBody(payload),
    ]);
    assert.match(graphClient.calls[0].body.message.body.content, /Ana <Silva>/);
    assert.deepEqual(sleepDelays(ledger), []);
    graphClient.assertExhausted();
});

test('quote request succeeds on the fifth mail attempt with exact retry sleeps', async (t) => {
    const { app, graphClient, ledger } = createTestApp();
    graphClient.enqueue(
        'POST',
        SEND_MAIL_PATH,
        new Error('attempt 1'),
        new Error('attempt 2'),
        new Error('attempt 3'),
        new Error('attempt 4'),
        {},
    );
    const { origin } = await startLoopback(app, t);

    const response = await postJson(origin, '/landingpage/solicitacaoorcamento', {});

    await assertJsonResponse(response, 200, {});
    assert.equal(graphClient.calls.length, 5);
    assert.deepEqual(sleepDelays(ledger), RETRY_DELAYS);
    assert.match(graphClient.calls[0].body.message.body.content, /<p>undefined<\/p>/);
    graphClient.assertExhausted();
});

test('quote request returns bare 500 after five failed mail attempts', async (t) => {
    const { app, graphClient, ledger } = createTestApp();
    graphClient.enqueue(
        'POST',
        SEND_MAIL_PATH,
        ...Array.from({ length: 5 }, (_, index) => new Error(`failure ${index + 1}`)),
    );
    const { origin } = await startLoopback(app, t);

    const response = await postJson(origin, '/landingpage/solicitacaoorcamento', {});

    await assertJsonResponse(response, 500, {});
    assert.equal(graphClient.calls.length, 5);
    assert.deepEqual(sleepDelays(ledger), RETRY_DELAYS);
    graphClient.assertExhausted();
});

test('Conecta rejects every required-field branch and malformed WhatsApp before Graph', async (t) => {
    const { app, graphClient } = createTestApp();
    const { origin } = await startLoopback(app, t);
    const invalidPayloads = [
        {},
        ...[
            'recommenderFullName',
            'benefitedCompany',
            'recommendedCompany',
            'recommendedProfessional',
            'recommendedWhatsapp',
        ].map((field) => validRecommendationPayload({ [field]: '   ' })),
        validRecommendationPayload({ recommendedWhatsapp: '+55 41 9999-0000' }),
        validRecommendationPayload({ recommendedWhatsapp: '+55 41 99999-0000 extra' }),
        validRecommendationPayload({ recommendedWhatsapp: 5541999990000 }),
    ];

    for (const payload of invalidPayloads) {
        const response = await postJson(origin, '/conecta/processa-recomendacao', payload);
        await assertJsonResponse(response, 400, { error: 'Erro_014' });
    }

    assert.equal(graphClient.calls.length, 0);
});

test('Conecta normalizes the recommender match and updates the first free slot exactly', async (t) => {
    const payload = validRecommendationPayload({
        recommenderFullName: '  alice   &   SILVA ',
        benefitedCompany: ' acme   <holding> ',
        recommendedCompany: '  Beta <Labs> & Co  ',
        recommendedProfessional: '  Bruno & <Equipe>  ',
        recommendedWhatsapp: '  +55 41 99999-0000  ',
    });
    const unrelated = recommendationCells({ 0: 'Other Company', 1: 'Other Person' });
    const slot = recommendationCells({ 0: ' Acme <Holding> ', 1: ' Alice & Silva ', 3: ' alice@example.test ' });
    const updatePath = `${RECOMMENDATIONS_ROWS}/itemAt(index=1)`;
    const expectedCells = expectedRecommendationCells(payload);
    const expectedMails = conectaMailBodies(payload, slot);
    const { app, graphClient, ledger } = createTestApp();
    graphClient.enqueue('GET', RECOMMENDATIONS_ROWS, workbookRows([unrelated, slot]));
    graphClient.enqueue('UPDATE', updatePath, {});
    graphClient.enqueue('POST', SEND_MAIL_PATH, {}, {});
    const { origin } = await startLoopback(app, t);

    const response = await postJson(origin, '/conecta/processa-recomendacao', payload);

    await assertJsonResponse(response, 200, {});
    assert.deepEqual(
        graphClient.calls.map(({ method, path }) => [method, path]),
        [
            ['GET', RECOMMENDATIONS_ROWS],
            ['UPDATE', updatePath],
            ['POST', SEND_MAIL_PATH],
            ['POST', SEND_MAIL_PATH],
        ],
    );
    assert.deepEqual(graphClient.calls[1].body, { values: [expectedCells] });
    assert.equal(graphClient.calls[1].body.values[0].length, 13);
    assert.deepEqual(graphClient.calls.slice(2).map(({ body }) => body), expectedMails);
    assert.deepEqual(sleepDelays(ledger), []);
    graphClient.assertExhausted();
});

test('Conecta appends an exact 13-cell row when no free matched slot exists', async (t) => {
    const payload = validRecommendationPayload();
    const occupied = recommendationCells({
        0: 'Acme <Holding>',
        1: 'Alice & Silva',
        2: 'Alicia',
        3: 'alice@example.test',
        4: 45000,
        5: 'Earlier Company',
        6: 'Earlier Person',
        7: '+55 41 98888-0000',
        8: '2. CONTATO',
        9: 'EM ANDAMENTO',
        10: 45001,
        11: 45002,
        12: 3,
    });
    const { app, graphClient } = createTestApp();
    graphClient.enqueue('GET', RECOMMENDATIONS_ROWS, workbookRows([occupied]));
    graphClient.enqueue('POST', `${RECOMMENDATIONS_ROWS}/add`, {});
    graphClient.enqueue('POST', SEND_MAIL_PATH, {}, {});
    const { origin } = await startLoopback(app, t);

    const response = await postJson(origin, '/conecta/processa-recomendacao', payload);

    await assertJsonResponse(response, 200, {});
    assert.deepEqual(graphClient.calls[1], {
        type: 'external-call',
        client: 'graph',
        method: 'POST',
        path: `${RECOMMENDATIONS_ROWS}/add`,
        body: { values: [expectedRecommendationCells(payload, occupied)] },
        payload: { values: [expectedRecommendationCells(payload, occupied)] },
    });
    const appendedCells = graphClient.calls[1].body.values[0];
    assert.equal(appendedCells.length, 13);
    assert.equal(appendedCells[2], null);
    assert.equal(appendedCells[12], '-');
    assert.deepEqual(
        graphClient.calls.slice(2).map(({ body }) => body),
        conectaMailBodies(payload, occupied),
    );
    graphClient.assertExhausted();
});

test('Conecta duplicate detection skips writes but repeats both mails on every request', async (t) => {
    const payload = validRecommendationPayload({
        recommendedCompany: ' beta   labs ',
        recommendedProfessional: ' BRUNO   SOUZA ',
        recommendedWhatsapp: ' +55 41 99999-0000 ',
    });
    const duplicate = recommendationCells({
        2: 'Alice',
        4: 45000,
        5: 'Beta Labs',
        6: 'Bruno Souza',
        7: '+55 41 99999-0000',
        8: '1. REALIZAR CONTATO INICIAL',
        9: 'A INICIAR',
        10: 45000,
        11: 45000,
        12: '-',
    });
    const { app, graphClient } = createTestApp();
    graphClient.enqueue('GET', RECOMMENDATIONS_ROWS, workbookRows([duplicate]), workbookRows([duplicate]));
    graphClient.enqueue('POST', SEND_MAIL_PATH, {}, {}, {}, {});
    const { origin } = await startLoopback(app, t);

    for (let repeat = 0; repeat < 2; repeat++) {
        const response = await postJson(origin, '/conecta/processa-recomendacao', payload);
        await assertJsonResponse(response, 200, {});
    }

    assert.deepEqual(
        graphClient.calls.map(({ method, path }) => [method, path]),
        [
            ['GET', RECOMMENDATIONS_ROWS],
            ['POST', SEND_MAIL_PATH],
            ['POST', SEND_MAIL_PATH],
            ['GET', RECOMMENDATIONS_ROWS],
            ['POST', SEND_MAIL_PATH],
            ['POST', SEND_MAIL_PATH],
        ],
    );
    assert.deepEqual(
        graphClient.calls.filter(({ method }) => method === 'POST').map(({ body }) => body),
        [
            ...conectaMailBodies(payload, duplicate),
            ...conectaMailBodies(payload, duplicate),
        ],
    );
    graphClient.assertExhausted();
});

test('Conecta retry after an ambiguous committed mutation skips a visible duplicate but mails again', async (t) => {
    t.mock.method(console, 'error', () => {});
    const payload = validRecommendationPayload();
    const slot = recommendationCells();
    const committed = recommendationCells({
        4: BRAZIL_NOW_SERIAL,
        5: payload.recommendedCompany,
        6: payload.recommendedProfessional,
        7: payload.recommendedWhatsapp,
        8: '1. REALIZAR CONTATO INICIAL',
        9: 'A INICIAR',
        10: BRAZIL_NOW_SERIAL,
        11: BRAZIL_NOW_SERIAL,
        12: '-',
    });
    const updatePath = `${RECOMMENDATIONS_ROWS}/itemAt(index=0)`;
    let mutationCommitted = false;
    const { app, graphClient } = createTestApp();
    graphClient.enqueue(
        'GET',
        RECOMMENDATIONS_ROWS,
        workbookRows([slot]),
        () => {
            assert.equal(mutationCommitted, true);
            return workbookRows([committed]);
        },
    );
    graphClient.enqueue('UPDATE', updatePath, () => {
        mutationCommitted = true;
        throw new Error('ambiguous mutation failure');
    });
    graphClient.enqueue('POST', SEND_MAIL_PATH, {}, {});
    const { origin } = await startLoopback(app, t);

    const firstResponse = await postJson(origin, '/conecta/processa-recomendacao', payload);
    await assertJsonResponse(firstResponse, 500, { error: 'Erro_017' });
    const retryResponse = await postJson(origin, '/conecta/processa-recomendacao', payload);
    await assertJsonResponse(retryResponse, 200, {});

    assert.equal(graphClient.calls.filter(({ path }) => path === updatePath).length, 1);
    assert.equal(graphClient.calls.filter(({ path }) => path === SEND_MAIL_PATH).length, 2);
    graphClient.assertExhausted();
});

test('Conecta read failures retry five times and no match is Erro_016', async (t) => {
    t.mock.method(console, 'error', () => {});

    await t.test('exhausted read is Erro_015', async (subtest) => {
        const { app, graphClient, ledger } = createTestApp();
        graphClient.enqueue(
            'GET',
            RECOMMENDATIONS_ROWS,
            ...Array.from({ length: 5 }, () => new Error('read failed')),
        );
        const { origin } = await startLoopback(app, subtest);

        const response = await postJson(origin, '/conecta/processa-recomendacao', validRecommendationPayload());

        await assertJsonResponse(response, 500, { error: 'Erro_015' });
        assert.equal(graphClient.calls.length, 5);
        assert.deepEqual(sleepDelays(ledger), RETRY_DELAYS);
        graphClient.assertExhausted();
    });

    await t.test('normalized no-match is Erro_016', async (subtest) => {
        const { app, graphClient, ledger } = createTestApp();
        graphClient.enqueue('GET', RECOMMENDATIONS_ROWS, workbookRows([
            recommendationCells({ 0: 'Different', 1: 'Person' }),
        ]));
        const { origin } = await startLoopback(app, subtest);

        const response = await postJson(origin, '/conecta/processa-recomendacao', validRecommendationPayload());

        await assertJsonResponse(response, 404, { error: 'Erro_016' });
        assert.equal(graphClient.calls.length, 1);
        assert.deepEqual(sleepDelays(ledger), []);
        graphClient.assertExhausted();
    });
});

test('Conecta free-slot and append mutations are each attempted only once on Erro_017', async (t) => {
    t.mock.method(console, 'error', () => {});

    for (const contractCase of [
        {
            name: 'free-slot update',
            cells: recommendationCells(),
            method: 'UPDATE',
            path: `${RECOMMENDATIONS_ROWS}/itemAt(index=0)`,
        },
        {
            name: 'append',
            cells: recommendationCells({ 4: 45000, 5: 'Occupied' }),
            method: 'POST',
            path: `${RECOMMENDATIONS_ROWS}/add`,
        },
    ]) {
        await t.test(contractCase.name, async (subtest) => {
            const { app, graphClient, ledger } = createTestApp();
            graphClient.enqueue('GET', RECOMMENDATIONS_ROWS, workbookRows([contractCase.cells]));
            graphClient.enqueue(contractCase.method, contractCase.path, new Error('ambiguous mutation failure'));
            const { origin } = await startLoopback(app, subtest);

            const response = await postJson(origin, '/conecta/processa-recomendacao', validRecommendationPayload());

            await assertJsonResponse(response, 500, { error: 'Erro_017' });
            assert.deepEqual(
                graphClient.calls.map(({ method, path }) => [method, path]),
                [['GET', RECOMMENDATIONS_ROWS], [contractCase.method, contractCase.path]],
            );
            assert.deepEqual(sleepDelays(ledger), []);
            graphClient.assertExhausted();
        });
    }
});

test('Conecta preserves workbook and mail partial successes on Erro_018', async (t) => {
    t.mock.method(console, 'error', () => {});

    for (const contractCase of [
        { name: 'internal mail fails', successfulMailCount: 0 },
        { name: 'confirmation mail fails after internal success', successfulMailCount: 1 },
    ]) {
        await t.test(contractCase.name, async (subtest) => {
            const slot = recommendationCells();
            const updatePath = `${RECOMMENDATIONS_ROWS}/itemAt(index=0)`;
            const { app, graphClient, ledger } = createTestApp();
            graphClient.enqueue('GET', RECOMMENDATIONS_ROWS, workbookRows([slot]));
            graphClient.enqueue('UPDATE', updatePath, {});
            graphClient.enqueue(
                'POST',
                SEND_MAIL_PATH,
                ...Array.from({ length: contractCase.successfulMailCount }, () => ({})),
                ...Array.from({ length: 5 }, () => new Error('mail failure')),
            );
            const { origin } = await startLoopback(app, subtest);

            const response = await postJson(origin, '/conecta/processa-recomendacao', validRecommendationPayload());

            await assertJsonResponse(response, 500, { error: 'Erro_018' });
            assert.deepEqual(
                graphClient.calls.slice(0, 2).map(({ method, path }) => [method, path]),
                [['GET', RECOMMENDATIONS_ROWS], ['UPDATE', updatePath]],
            );
            assert.equal(
                graphClient.calls.filter(({ path }) => path === SEND_MAIL_PATH).length,
                contractCase.successfulMailCount + 5,
            );
            assert.deepEqual(sleepDelays(ledger), RETRY_DELAYS);
            graphClient.assertExhausted();
        });
    }
});

test('client intake writes exact positional rows, mails exact content, and dedupes a visible repeat', async (t) => {
    const payload = validIntakePayload();
    const participant = payload.participants[0];
    const platformCells = expectedPlatformCells(participant);
    const clientCells = expectedClientCells(payload, participant);
    const { app, graphClient, ledger } = createTestApp();

    graphClient.enqueue('GET', PLATFORM_ROWS, workbookRows([]), workbookRows([platformCells]));
    graphClient.enqueue('GET', CLIENTS_ROWS, workbookRows([]), workbookRows([clientCells]));
    graphClient.enqueue('POST', `${PLATFORM_ROWS}/add`, {});
    graphClient.enqueue('POST', `${CLIENTS_ROWS}/add`, {});
    graphClient.enqueue('POST', SEND_MAIL_PATH, {}, {});
    const { origin } = await startLoopback(app, t);

    for (let repeat = 0; repeat < 2; repeat++) {
        const response = await postJson(origin, '/clientes/processa-formulario', payload);
        await assertJsonResponse(response, 200, {});
    }

    assert.deepEqual(
        graphClient.calls.map(({ method, path }) => [method, path]),
        [
            ['GET', PLATFORM_ROWS],
            ['GET', CLIENTS_ROWS],
            ['POST', `${PLATFORM_ROWS}/add`],
            ['POST', `${CLIENTS_ROWS}/add`],
            ['POST', SEND_MAIL_PATH],
            ['GET', PLATFORM_ROWS],
            ['GET', CLIENTS_ROWS],
            ['POST', SEND_MAIL_PATH],
        ],
    );
    assert.deepEqual(graphClient.calls[2].body, { values: [platformCells] });
    assert.deepEqual(graphClient.calls[3].body, { values: [clientCells] });
    assert.equal(platformCells.length, 22);
    assert.deepEqual(
        [platformCells[1], platformCells[7], platformCells[9], platformCells[20]],
        [null, null, null, null],
    );
    assert.equal(platformCells[3], 100000000000);
    assert.equal(platformCells[6], ACCESS_DEADLINE_SERIAL);
    assert.equal(platformCells[21], 'FMG-0000-0000');
    assert.equal(clientCells.length, 13);
    assert.deepEqual([clientCells[1], clientCells[2], clientCells[11]], [null, null, null]);
    assert.equal(clientCells[6], 7);
    assert.equal(clientCells[7], '-');
    assert.deepEqual(
        graphClient.calls.filter(({ path }) => path === SEND_MAIL_PATH).map(({ body }) => body),
        [intakeMailBody(payload), intakeMailBody(payload)],
    );
    assert.deepEqual(sleepDelays(ledger), []);
    graphClient.assertExhausted();
});

test('client intake dedupes platform email and client CPF independently, including within-request duplicates', async (t) => {
    const payload = validIntakePayload();
    payload.shippingAddress.number = '12A';
    payload.shippingAddress.complement = 'Fundos';
    payload.participants = [
        {
            fullName: 'Email Existing',
            email: ' Existing@Example.Test ',
            cpf: '111.111.111-11',
            role: 'One',
            areaCode: '41',
            whatsapp: '90000-0001',
        },
        {
            fullName: 'CPF Existing',
            email: 'new@example.test',
            cpf: '222.222.222-22',
            role: 'Two',
            areaCode: '41',
            whatsapp: '90000-0002',
        },
        {
            fullName: 'Request Email Duplicate',
            email: ' NEW@example.test ',
            cpf: '333.333.333-33',
            role: 'Three',
            areaCode: '41',
            whatsapp: '90000-0003',
        },
    ];
    const existingPlatformCells = new Array(22).fill(null);
    existingPlatformCells[2] = 'existing@example.test';
    existingPlatformCells[21] = 'FMG-ZZZZ-ZZZZ';
    const existingClientCells = new Array(13).fill(null);
    existingClientCells[4] = '22222222222';
    const expectedPlatformRows = [expectedPlatformCells(payload.participants[1])];
    const expectedClientRows = [
        expectedClientCells(payload, payload.participants[0]),
        expectedClientCells(payload, payload.participants[2]),
    ];
    const { app, graphClient } = createTestApp();
    graphClient.enqueue('GET', PLATFORM_ROWS, workbookRows([existingPlatformCells]));
    graphClient.enqueue('GET', CLIENTS_ROWS, workbookRows([existingClientCells]));
    graphClient.enqueue('POST', `${PLATFORM_ROWS}/add`, {});
    graphClient.enqueue('POST', `${CLIENTS_ROWS}/add`, {});
    graphClient.enqueue('POST', SEND_MAIL_PATH, {});
    const { origin } = await startLoopback(app, t);

    const response = await postJson(origin, '/clientes/processa-formulario', payload);

    await assertJsonResponse(response, 200, {});
    assert.deepEqual(graphClient.calls[2].body, { values: expectedPlatformRows });
    assert.deepEqual(graphClient.calls[3].body, { values: expectedClientRows });
    assert.equal(graphClient.calls[2].body.values[0].length, 22);
    assert.ok(graphClient.calls[3].body.values.every((cells) => cells.length === 13));
    assert.equal(graphClient.calls[3].body.values[0][6], '12A');
    assert.equal(graphClient.calls[3].body.values[0][7], 'Fundos');
    graphClient.assertExhausted();
});

test('client intake regenerates certificate IDs that collide with observed and same-request IDs', async (t) => {
    const payload = validIntakePayload();
    payload.participants.push({
        fullName: 'Bruno Certificado',
        email: 'bruno@example.test',
        cpf: '444.444.444-44',
        role: 'Gerente',
        areaCode: '41',
        whatsapp: '99999-4444',
    });
    const existingPlatformCells = new Array(22).fill(null);
    existingPlatformCells[2] = 'existing@example.test';
    existingPlatformCells[21] = 'FMG-0000-0000';
    const passwords = [111111111111, 222222222222];
    const certificateCharacters = [
        ...Array(8).fill(0),
        ...Array(8).fill(1),
        ...Array(8).fill(1),
        ...Array(8).fill(2),
    ];
    const passwordBounds = [];
    function randomInt(min, max) {
        if (max !== undefined) {
            passwordBounds.push([min, max]);
            return passwords.shift();
        }
        assert.equal(min, 32);
        return certificateCharacters.shift();
    }
    const expectedRows = [
        expectedPlatformCells(payload.participants[0], {
            password: 111111111111,
            certificateId: 'FMG-1111-1111',
        }),
        expectedPlatformCells(payload.participants[1], {
            password: 222222222222,
            certificateId: 'FMG-2222-2222',
        }),
    ];
    const { app, graphClient } = createTestApp({ randomInt });
    graphClient.enqueue('GET', PLATFORM_ROWS, workbookRows([existingPlatformCells]));
    graphClient.enqueue('GET', CLIENTS_ROWS, workbookRows([]));
    graphClient.enqueue('POST', `${PLATFORM_ROWS}/add`, {});
    graphClient.enqueue('POST', `${CLIENTS_ROWS}/add`, {});
    graphClient.enqueue('POST', SEND_MAIL_PATH, {});
    const { origin } = await startLoopback(app, t);

    const response = await postJson(origin, '/clientes/processa-formulario', payload);

    await assertJsonResponse(response, 200, {});
    assert.deepEqual(graphClient.calls[2].body, { values: expectedRows });
    assert.deepEqual(passwords, []);
    assert.deepEqual(passwordBounds, [
        [100000000000, 1000000000000],
        [100000000000, 1000000000000],
    ]);
    assert.deepEqual(certificateCharacters, []);
    graphClient.assertExhausted();
});

test('client intake certificate generation uses the complete exact 32-character alphabet', async (t) => {
    const payload = validIntakePayload();
    payload.participants = Array.from({ length: 4 }, (_, index) => ({
        fullName: `Pessoa ${index + 1}`,
        email: `pessoa-${index + 1}@example.test`,
        cpf: `000.000.000-0${index + 1}`,
        role: 'Participante',
        areaCode: '41',
        whatsapp: `99999-000${index + 1}`,
    }));
    const certificateIndexes = Array.from({ length: 32 }, (_, index) => index);
    function randomInt(min, max) {
        if (max !== undefined) {
            assert.deepEqual([min, max], [100000000000, 1000000000000]);
            return min;
        }
        assert.equal(min, 32);
        return certificateIndexes.shift();
    }
    const { app, graphClient } = createTestApp({ randomInt });
    graphClient.enqueue('GET', PLATFORM_ROWS, workbookRows([]));
    graphClient.enqueue('GET', CLIENTS_ROWS, workbookRows([]));
    graphClient.enqueue('POST', `${PLATFORM_ROWS}/add`, {});
    graphClient.enqueue('POST', `${CLIENTS_ROWS}/add`, {});
    graphClient.enqueue('POST', SEND_MAIL_PATH, {});
    const { origin } = await startLoopback(app, t);

    const response = await postJson(origin, '/clientes/processa-formulario', payload);

    await assertJsonResponse(response, 200, {});
    assert.deepEqual(
        graphClient.calls[2].body.values.map((cells) => cells[21]),
        ['FMG-0123-4567', 'FMG-89AB-CDEF', 'FMG-GHJK-MNPQ', 'FMG-RSTV-WXYZ'],
    );
    assert.deepEqual(certificateIndexes, []);
    graphClient.assertExhausted();
});

test('client intake rejects every defined validation branch before Graph', async (t) => {
    const valid = validIntakePayload();
    const blankCompany = clone(valid);
    blankCompany.company.legalName = '   ';
    const notArray = clone(valid);
    notArray.participants = {};
    const emptyParticipants = clone(valid);
    emptyParticipants.participants = [];
    const tooManyParticipants = clone(valid);
    tooManyParticipants.participants = Array.from(
        { length: 26 },
        (_, index) => ({ ...valid.participants[0], email: `${index}@example.test`, cpf: String(index) }),
    );
    const nullParticipant = clone(valid);
    nullParticipant.participants = [null];
    const invalidPayloads = [blankCompany, notArray, emptyParticipants, tooManyParticipants, nullParticipant];

    for (const field of ['fullName', 'email', 'cpf']) {
        const invalidParticipant = clone(valid);
        invalidParticipant.participants[0][field] = '   ';
        invalidPayloads.push(invalidParticipant);
    }

    const { app, graphClient } = createTestApp();
    const { origin } = await startLoopback(app, t);

    for (const payload of invalidPayloads) {
        const response = await postJson(origin, '/clientes/processa-formulario', payload);
        await assertJsonResponse(response, 400, { error: 'Erro_013' });
    }

    assert.equal(graphClient.calls.length, 0);
});

test('client intake read failures preserve Erro_001 and Erro_011 retry boundaries', async (t) => {
    await t.test('platform read exhaustion is Erro_001', async (subtest) => {
        const { app, graphClient, ledger } = createTestApp();
        graphClient.enqueue(
            'GET',
            PLATFORM_ROWS,
            ...Array.from({ length: 5 }, () => new Error('platform read failure')),
        );
        const { origin } = await startLoopback(app, subtest);

        const response = await postJson(origin, '/clientes/processa-formulario', validIntakePayload());

        await assertJsonResponse(response, 500, { error: 'Erro_001' });
        assert.equal(graphClient.calls.length, 5);
        assert.deepEqual(sleepDelays(ledger), RETRY_DELAYS);
        graphClient.assertExhausted();
    });

    await t.test('client read exhaustion after platform success is Erro_011', async (subtest) => {
        const { app, graphClient, ledger } = createTestApp();
        graphClient.enqueue('GET', PLATFORM_ROWS, workbookRows([]));
        graphClient.enqueue(
            'GET',
            CLIENTS_ROWS,
            ...Array.from({ length: 5 }, () => new Error('client read failure')),
        );
        const { origin } = await startLoopback(app, subtest);

        const response = await postJson(origin, '/clientes/processa-formulario', validIntakePayload());

        await assertJsonResponse(response, 500, { error: 'Erro_011' });
        assert.deepEqual(
            graphClient.calls.map(({ method, path }) => [method, path]),
            [
                ['GET', PLATFORM_ROWS],
                ...Array.from({ length: 5 }, () => ['GET', CLIENTS_ROWS]),
            ],
        );
        assert.deepEqual(sleepDelays(ledger), RETRY_DELAYS);
        graphClient.assertExhausted();
    });
});

test('client intake writes are not application-retried and retain sequential partial success', async (t) => {
    await t.test('platform write failure is one attempt and Erro_008', async (subtest) => {
        const { app, graphClient, ledger } = createTestApp();
        graphClient.enqueue('GET', PLATFORM_ROWS, workbookRows([]));
        graphClient.enqueue('GET', CLIENTS_ROWS, workbookRows([]));
        graphClient.enqueue('POST', `${PLATFORM_ROWS}/add`, new Error('ambiguous platform write'));
        const { origin } = await startLoopback(app, subtest);

        const response = await postJson(origin, '/clientes/processa-formulario', validIntakePayload());

        await assertJsonResponse(response, 500, { error: 'Erro_008' });
        assert.deepEqual(
            graphClient.calls.map(({ method, path }) => [method, path]),
            [['GET', PLATFORM_ROWS], ['GET', CLIENTS_ROWS], ['POST', `${PLATFORM_ROWS}/add`]],
        );
        assert.deepEqual(sleepDelays(ledger), []);
        graphClient.assertExhausted();
    });

    await t.test('client write failure follows committed platform write and is Erro_010', async (subtest) => {
        const { app, graphClient, ledger } = createTestApp();
        graphClient.enqueue('GET', PLATFORM_ROWS, workbookRows([]));
        graphClient.enqueue('GET', CLIENTS_ROWS, workbookRows([]));
        graphClient.enqueue('POST', `${PLATFORM_ROWS}/add`, {});
        graphClient.enqueue('POST', `${CLIENTS_ROWS}/add`, new Error('ambiguous client write'));
        const { origin } = await startLoopback(app, subtest);

        const response = await postJson(origin, '/clientes/processa-formulario', validIntakePayload());

        await assertJsonResponse(response, 500, { error: 'Erro_010' });
        assert.deepEqual(
            graphClient.calls.map(({ method, path }) => [method, path]),
            [
                ['GET', PLATFORM_ROWS],
                ['GET', CLIENTS_ROWS],
                ['POST', `${PLATFORM_ROWS}/add`],
                ['POST', `${CLIENTS_ROWS}/add`],
            ],
        );
        assert.deepEqual(sleepDelays(ledger), []);
        graphClient.assertExhausted();
    });
});

test('client intake mail exhaustion is Erro_012 after both writes persist', async (t) => {
    const { app, graphClient, ledger } = createTestApp();
    graphClient.enqueue('GET', PLATFORM_ROWS, workbookRows([]));
    graphClient.enqueue('GET', CLIENTS_ROWS, workbookRows([]));
    graphClient.enqueue('POST', `${PLATFORM_ROWS}/add`, {});
    graphClient.enqueue('POST', `${CLIENTS_ROWS}/add`, {});
    graphClient.enqueue(
        'POST',
        SEND_MAIL_PATH,
        ...Array.from({ length: 5 }, () => new Error('mail failure')),
    );
    const { origin } = await startLoopback(app, t);

    const response = await postJson(origin, '/clientes/processa-formulario', validIntakePayload());

    await assertJsonResponse(response, 500, { error: 'Erro_012' });
    assert.deepEqual(
        graphClient.calls.slice(0, 4).map(({ method, path }) => [method, path]),
        [
            ['GET', PLATFORM_ROWS],
            ['GET', CLIENTS_ROWS],
            ['POST', `${PLATFORM_ROWS}/add`],
            ['POST', `${CLIENTS_ROWS}/add`],
        ],
    );
    assert.equal(graphClient.calls.filter(({ path }) => path === SEND_MAIL_PATH).length, 5);
    assert.deepEqual(sleepDelays(ledger), RETRY_DELAYS);
    graphClient.assertExhausted();
});

test('access release responds before deferred Graph read and sends rows 35 through 41 exactly', async (t) => {
    t.mock.method(console, 'log', () => {});
    const deferredRead = createDeferred();
    const { data, recipients } = accessReleaseRows();
    const { app, graphClient, ledger, scheduler } = createTestApp();
    graphClient.enqueue('GET', PLATFORM_ROWS, () => deferredRead.promise);
    graphClient.enqueue('POST', SEND_MAIL_PATH, ...Array.from({ length: 7 }, () => ({})));
    const { origin } = await startLoopback(app, t);

    const response = await fetch(`${origin}/clientes/liberacao-acesso-plataforma`, { method: 'POST' });

    assert.equal(response.status, 200);
    assert.equal(response.headers.get('content-type'), null);
    assert.equal(await response.text(), '');
    assert.deepEqual(
        graphClient.calls.map(({ method, path }) => [method, path]),
        [['GET', PLATFORM_ROWS]],
    );
    assert.deepEqual(scheduler.pending, []);

    deferredRead.resolve(data);
    await new Promise((resolve) => setImmediate(resolve));

    assert.deepEqual(scheduler.pending, [{ id: 1, delay: 1_000 }]);
    await scheduler.runNext();

    const mailCalls = graphClient.calls.filter(({ path }) => path === SEND_MAIL_PATH);
    assert.equal(mailCalls.length, 7);
    assert.deepEqual(mailCalls.map(({ body }) => body), recipients.map(accessReleaseMailBody));
    assert.deepEqual(
        mailCalls.map(({ body }) => body.message.toRecipients[0].emailAddress.address),
        recipients.map(({ email }) => email),
    );
    assert.deepEqual(sleepDelays(ledger), Array(7).fill(2_000));
    assert.deepEqual(
        ledger
            .filter((entry) => entry.type === 'sleep' || (
                entry.type === 'external-call' && entry.path === SEND_MAIL_PATH
            ))
            .map((entry) => (
                entry.type === 'sleep'
                    ? ['sleep', entry.delay]
                    : ['mail', entry.body.message.toRecipients[0].emailAddress.address]
            )),
        recipients.flatMap(({ email }) => [['mail', email], ['sleep', 2_000]]),
    );
    assert.deepEqual(
        ledger.filter(({ type }) => type === 'timer-scheduled' || type === 'timer-run'),
        [
            { type: 'timer-scheduled', id: 1, delay: 1_000 },
            { type: 'timer-run', id: 1, delay: 1_000 },
        ],
    );
    graphClient.assertExhausted();
});

test('access release mail failure is not retried and stops the remaining suffix', async (t) => {
    t.mock.method(console, 'log', () => {});
    const { data, recipients } = accessReleaseRows();
    const { app, graphClient, ledger, scheduler } = createTestApp();
    graphClient.enqueue('GET', PLATFORM_ROWS, data);
    graphClient.enqueue('POST', SEND_MAIL_PATH, {}, {}, new Error('third mail failed'));
    const { origin } = await startLoopback(app, t);

    const response = await fetch(`${origin}/clientes/liberacao-acesso-plataforma`, { method: 'POST' });
    assert.equal(response.status, 200);
    assert.equal(await response.text(), '');
    await new Promise((resolve) => setImmediate(resolve));

    assert.deepEqual(scheduler.pending, [{ id: 1, delay: 1_000 }]);
    await assert.rejects(scheduler.runNext(), /third mail failed/);

    const mailCalls = graphClient.calls.filter(({ path }) => path === SEND_MAIL_PATH);
    assert.equal(mailCalls.length, 3);
    assert.deepEqual(
        mailCalls.map(({ body }) => body),
        recipients.slice(0, 3).map(accessReleaseMailBody),
    );
    assert.deepEqual(sleepDelays(ledger), [2_000, 2_000]);
    assert.deepEqual(
        ledger
            .filter((entry) => entry.type === 'sleep' || (
                entry.type === 'external-call' && entry.path === SEND_MAIL_PATH
            ))
            .map(({ type, delay, body }) => (
                type === 'sleep'
                    ? ['sleep', delay]
                    : ['mail', body.message.toRecipients[0].emailAddress.address]
            )),
        [
            ['mail', recipients[0].email],
            ['sleep', 2_000],
            ['mail', recipients[1].email],
            ['sleep', 2_000],
            ['mail', recipients[2].email],
        ],
    );
    assert.deepEqual(scheduler.pending, []);
    graphClient.assertExhausted();
});

test('repeated access-release requests schedule and resend the same seven recipients', async (t) => {
    t.mock.method(console, 'log', () => {});
    const { data, recipients } = accessReleaseRows();
    const { app, graphClient, ledger, scheduler } = createTestApp();
    graphClient.enqueue('GET', PLATFORM_ROWS, data, data);
    graphClient.enqueue('POST', SEND_MAIL_PATH, ...Array.from({ length: 14 }, () => ({})));
    const { origin } = await startLoopback(app, t);

    for (let repeat = 0; repeat < 2; repeat++) {
        const response = await fetch(`${origin}/clientes/liberacao-acesso-plataforma`, { method: 'POST' });
        assert.equal(response.status, 200);
        assert.equal(await response.text(), '');
    }
    await new Promise((resolve) => setImmediate(resolve));

    assert.deepEqual(scheduler.pending, [{ id: 1, delay: 1_000 }, { id: 2, delay: 1_000 }]);
    await scheduler.runNext();
    await scheduler.runNext();

    const recipientAddresses = graphClient.calls
        .filter(({ path }) => path === SEND_MAIL_PATH)
        .map(({ body }) => body.message.toRecipients[0].emailAddress.address);
    assert.deepEqual(recipientAddresses, [
        ...recipients.map(({ email }) => email),
        ...recipients.map(({ email }) => email),
    ]);
    assert.deepEqual(sleepDelays(ledger), Array(14).fill(2_000));
    graphClient.assertExhausted();
});
