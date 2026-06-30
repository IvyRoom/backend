///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////// IMPORTAÇÃO DE BIBLIOTECAS, CRIAÇÃO DE FUNÇÕES E VARIÁVEIS /////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Configura comunicação com variáveis de ambiente.
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

const dotenv = require('dotenv');
dotenv.config();

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Configura comunicação com HTTP Requests.
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

const express = require('express');
const cors = require('cors');
const app = express();
app.use(cors());
app.use(express.json());
app.use('/img', express.static('img'));
app.listen(process.env.PORT || 3000);

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Configura comunicação com o Microsoft Graph API.
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

const { Client } = require('@microsoft/microsoft-graph-client');
const { ConfidentialClientApplication } = require('@azure/msal-node');
const cca = new ConfidentialClientApplication({ auth: { clientId: process.env.CLIENT_ID, authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`, clientSecret: process.env.CLIENT_SECRET } });
let Microsoft_Graph_API_AccessToken;
let Microsoft_Graph_API_Client = Client.init({authProvider:(done)=>{done(null, Microsoft_Graph_API_AccessToken)}});
let Microsoft_Graph_API_SetTimeout;
let Microsoft_Graph_API_Delay = 2000;

async function refreshMicrosoftGraphAccessToken() {
    
    try {

        const response = await cca.acquireTokenByClientCredential({scopes: ['https://graph.microsoft.com/.default']});
        Microsoft_Graph_API_AccessToken = response.accessToken;

        if (Microsoft_Graph_API_SetTimeout) clearTimeout(Microsoft_Graph_API_SetTimeout);
        Microsoft_Graph_API_SetTimeout = setTimeout(refreshMicrosoftGraphAccessToken, Math.max(new Date(response.expiresOn).getTime() - Date.now() - 5 * 60 * 1000, 60000));
        
        Microsoft_Graph_API_Delay = 2000;

    } catch (err) {

        if (Microsoft_Graph_API_SetTimeout) clearTimeout(Microsoft_Graph_API_SetTimeout);
        Microsoft_Graph_API_SetTimeout = setTimeout(refreshMicrosoftGraphAccessToken, Microsoft_Graph_API_Delay);
        Microsoft_Graph_API_Delay = Math.min(Microsoft_Graph_API_Delay * 2, 60000);
    
    }

}

refreshMicrosoftGraphAccessToken();

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Configura comunicação com o Azure Face API.
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

const { AzureKeyCredential } = require("@azure/core-auth");
const FaceClient = require("@azure-rest/ai-vision-face").default;

const Azure_Face_API_Credential = new AzureKeyCredential(process.env.AZURE_FACE_API_KEY);
const Azure_Face_API_Client = FaceClient(process.env.AZURE_FACE_API_ENDPOINT, Azure_Face_API_Credential);

const multer = require('multer');
const { v4: uuidv4 } = require('uuid');

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Função que transforma datas no formato Excel em datas no formato DD/MMM/AAAA.
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

function ConverteData(DataExcel) {
    const date = new Date((DataExcel - 25569) * 86400 * 1000);
    return date.toLocaleDateString('pt-BR', { day: '2-digit', month: 'short', year: 'numeric' }).replace(/\bde\b|\./g, '').replace(/\s+/g, '/');
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Função de retry.
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

async function retry(fn, retries = 5) {
    for (let i = 0; i < retries; i++) {
        try { return await fn(); } 
        catch (err) {
            if (i === retries - 1) throw err;
            await new Promise(r => setTimeout(r, 500 * (i + 1)));
        }
    }
}

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////// LANDING PAGE: PROCESSA SUBMISSÃO DO FORMULÁRIO DE SOLICITAÇÃO DE ORÇAMENTO /////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

app.post('/landingpage/solicitacaoorcamento', async (req, res) => {
    
    let { Solicitante_NomeCompleto, Solicitante_Email, Solicitante_Telefone, Solicitante_Cargo, Solicitante_NomeEmpresa, Solicitante_CNPJ, Solicitante_NúmerodeParticipantes, Solicitante_Observações } = req.body;
    
    try { await retry(() => Microsoft_Graph_API_Client.api('/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/sendMail').post({ message: { subject: 'Machado - Nova Solicitação de Orçamento', body: { contentType: 'HTML', content: `<p><b>Dados do Solicitante:</b></p><p>${Solicitante_NomeCompleto}</p><p>${Solicitante_Email}</p><p>${Solicitante_Telefone}</p><p>${Solicitante_Cargo}</p><p><b>Dados da Empresa:</b></p><p>${Solicitante_NomeEmpresa}</p><p>${Solicitante_CNPJ}</p><p>${Solicitante_NúmerodeParticipantes}</p><p>${Solicitante_Observações}</p><p><img width="500" height="auto" src="https://plataforma-backend-v3.azurewebsites.net/img/ASSINATURA_E-MAIL.jpg"/></p>`}, toRecipients: [{ emailAddress: { address: 'contato@machadogestao.com' } }] } })) }
    catch (err) { return res.status(500).json({}) }

    return res.status(200).json({});

});

function accessDeadlineSerial(daysFromToday) {
    const today = new Date();
    const utcMidnight = Date.UTC(today.getFullYear(), today.getMonth(), today.getDate());
    return Math.floor(utcMidnight / 86400000) + 25569 + daysFromToday;
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Função que gera um Certificado ID# aleatório e não-sequencial (resistente a enumeração).
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

const crypto = require('node:crypto');
const Alfabeto_CertificadoID = '0123456789ABCDEFGHJKMNPQRSTVWXYZ';

function GeraCertificadoID() {
    let Sufixo = '';
    for (let i = 0; i < 8; i++) Sufixo += Alfabeto_CertificadoID[crypto.randomInt(Alfabeto_CertificadoID.length)];
    return `FMG-${Sufixo.slice(0, 4)}-${Sufixo.slice(4)}`;
}


///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////// VALIDAÇÃO PÚBLICA DE CERTIFICADOS PELO CERTIFICADO ID# ///////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


app.get('/landingpage/validacaocertificados/:Solicitante_CertificadoID', async (req, res) => {

    const Solicitante_CertificadoID = String(req.params.Solicitante_CertificadoID || '').trim().toUpperCase();

    let BD_Plataforma;
    try { BD_Plataforma = await retry(() => Microsoft_Graph_API_Client.api('/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECSBYCZNYGEWFFDLEOZ36WI2PDWO/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows').get()) }
    catch (err) { return res.status(500).json({ error: 'Erro_001' }) }

    const Linha = Solicitante_CertificadoID
        ? BD_Plataforma.value.find((row) => String(row.values[0][21] == null ? '' : row.values[0][21]).trim().toUpperCase() === Solicitante_CertificadoID)
        : undefined;

    if (!Linha) return res.status(200).json({ Certificado_Válido: false });

    const Acumulado_Bruto = Number(Linha.values[0][20]);
    const Acumulado_Percentual = !isFinite(Acumulado_Bruto) ? 0 : (Acumulado_Bruto <= 1 ? Acumulado_Bruto * 100 : Acumulado_Bruto);

    if (Acumulado_Percentual < 70) return res.status(200).json({ Certificado_Válido: false });

    return res.status(200).json({
        Certificado_Válido: true,
        Titular_NomeCompleto: Linha.values[0][0],
        Acumulado_Percentual: Math.round(Acumulado_Percentual),
        Certificado_ID: String(Linha.values[0][21]).trim().toUpperCase()
    });

});

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////// FORMULÁRIO DE INFORMAÇÕES INICIAIS: PROCESSA SUBMISSÃO DO FORMULÁRIO ////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


app.post('/clientes/processa-formulario', async (req, res) => {

    const participants = Array.isArray(req.body && req.body.participants) ? req.body.participants : [];
    const company = (req.body && req.body.company) || {};
    const shipping = (req.body && req.body.shippingAddress) || {};
    const legalRep = (req.body && req.body.legalRepresentative) || {};
    const adminAssistant = (req.body && req.body.adminAssistant) || {};

    const plataformaTable = '/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECSBYCZNYGEWFFDLEOZ36WI2PDWO/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}';
    const clientesTable = '/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECQNNRY4S7VCKBF2SOETFSLESSLH/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}';

    let plataformaData, clientesData;
    try { plataformaData = await retry(() => Microsoft_Graph_API_Client.api(`${plataformaTable}/rows`).get()); }
    catch (err) { return res.status(500).json({ error: 'Erro_001' }); }
    try { clientesData = await retry(() => Microsoft_Graph_API_Client.api(`${clientesTable}/rows`).get()); }
    catch (err) { return res.status(500).json({ error: 'Erro_011' }); }

    const onlyDigits = (value) => String(value == null ? '' : value).replace(/\D/g, '');
    const existingEmails = new Set(plataformaData.value.map((row) => String(row.values[0][2] == null ? '' : row.values[0][2]).trim().toLowerCase()));
    const existingCpfs = new Set(clientesData.value.map((row) => onlyDigits(row.values[0][4])));
    const existingCertificadoIDs = new Set(plataformaData.value.map((row) => String(row.values[0][21] == null ? '' : row.values[0][21]).trim().toUpperCase()).filter(Boolean));

    const deadline = accessDeadlineSerial(60);
    const addressNumber = /^\d+$/.test(shipping.number) ? Number(shipping.number) : shipping.number;

    const plataformaRows = participants
        .filter((participant) => {
            const email = String(participant.email || '').trim().toLowerCase();
            if (existingEmails.has(email)) return false;
            existingEmails.add(email);
            return true;
        })
        .map((participant) => {
            const cells = new Array(22).fill(null);
            cells[0] = participant.fullName;
            cells[2] = participant.email;
            cells[3] = Math.floor(100000000000 + Math.random() * 900000000000);
            cells[4] = 'Ativo';
            cells[5] = 'Não';
            cells[6] = deadline;
            cells[8] = 0;
            for (let module = 10; module <= 19; module++) cells[module] = 0;
            let CertificadoID;
            do { CertificadoID = GeraCertificadoID(); } while (existingCertificadoIDs.has(CertificadoID));
            existingCertificadoIDs.add(CertificadoID);
            cells[21] = CertificadoID;
            return cells;
        });

    const clientesRows = participants
        .filter((participant) => {
            const cpf = onlyDigits(participant.cpf);
            if (existingCpfs.has(cpf)) return false;
            existingCpfs.add(cpf);
            return true;
        })
        .map((participant) => {
            const cells = new Array(13).fill(null);
            cells[0] = company.legalName;
            cells[3] = participant.fullName;
            cells[4] = participant.cpf;
            cells[5] = shipping.street;
            cells[6] = addressNumber;
            cells[7] = shipping.complement || '-';
            cells[8] = shipping.neighborhood;
            cells[9] = shipping.city;
            cells[10] = shipping.state;
            cells[12] = shipping.postalCode;
            return cells;
        });

    if (plataformaRows.length > 0) {
        try { await Microsoft_Graph_API_Client.api(`${plataformaTable}/rows/add`).post({ values: plataformaRows }); }
        catch (err) { return res.status(500).json({ error: 'Erro_008' }); }
    }

    if (clientesRows.length > 0) {
        try { await Microsoft_Graph_API_Client.api(`${clientesTable}/rows/add`).post({ values: clientesRows }); }
        catch (err) { return res.status(500).json({ error: 'Erro_010' }); }
    }

    const companyAddress = company.address || {};
    const pessoaHTML = (rotulo, p) => `<p><b>${rotulo}</b></p><p>Nome Completo: ${p.fullName}</p><p>CPF: ${p.cpf}</p><p>Cargo: ${p.role}</p><p>DDD: ${p.areaCode}</p><p>WhatsApp: ${p.whatsapp}</p><p>E-mail: ${p.email}</p>`;
    const participantesHTML = participants.map((p, i) => `<p>${i + 1}. ${p.fullName} — Cargo: ${p.role} · DDD: ${p.areaCode} · WhatsApp: ${p.whatsapp}</p>`).join('');
    const emailContent = `<p>Um novo Formulário de Informações Iniciais foi preenchido.</p><p><b>Pessoa Jurídica Contratante</b></p><p>Razão Social: ${company.legalName}</p><p>CNPJ: ${company.cnpj}</p><p>CEP: ${companyAddress.postalCode}</p><p>Rua: ${companyAddress.street}</p><p>Número: ${companyAddress.number}</p><p>Complemento: ${companyAddress.complement}</p><p>Bairro: ${companyAddress.neighborhood}</p><p>Cidade: ${companyAddress.city}</p><p>Estado: ${companyAddress.state}</p>${pessoaHTML('Representante Jurídico', legalRep)}${pessoaHTML('Auxiliar Administrativo Financeiro', adminAssistant)}<p><b>Participantes</b></p>${participantesHTML}<p><img width="500" height="auto" src="https://plataforma-backend-v3.azurewebsites.net/img/ASSINATURA_E-MAIL.jpg"/></p>`;

    try { await retry(() => Microsoft_Graph_API_Client.api('/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/sendMail').post({ message: { subject: 'Machado: novo Formulário de Informações Iniciais preenchido', body: { contentType: 'HTML', content: emailContent }, toRecipients: [{ emailAddress: { address: 'contato@machadogestao.com' } }] } })); }
    catch (err) { return res.status(500).json({ error: 'Erro_012' }); }

    return res.status(200).json({});

});

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////// E-MAIL: LIBERA ACESSOS À PLATAFORMA //////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

app.post('/clientes/liberacao-acesso-plataforma', async (req, res) => {

    res.status(200).send();
    console.log(`1. Request recebida.`);

    const BD_Plataforma = await Microsoft_Graph_API_Client.api('/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECSBYCZNYGEWFFDLEOZ36WI2PDWO/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows').get();
    if (BD_Plataforma !== null) console.log(`2. BD_Plataforma obtida.`);

    let Número_Email_Enviado = 0;
    let Linha_Inicial = 39;
    let Linha_Final = 45;

    async function Envia_Email_Clientes() {

        for (let LinhaAtual = (Linha_Inicial - 4); LinhaAtual <= (Linha_Final - 4); LinhaAtual++) {

            let Cliente_PrimeiroNome = BD_Plataforma.value[LinhaAtual].values[0][1];
            let Cliente_Email = BD_Plataforma.value[LinhaAtual].values[0][2];
            let Cliente_Senha = BD_Plataforma.value[LinhaAtual].values[0][3];

            Número_Email_Enviado++;

            console.log(`3. E-mail #${Número_Email_Enviado} enviado para: ${Cliente_PrimeiroNome}`);

            if (LinhaAtual === (Linha_Final - 4)) console.log(`--- fim ---`);

            await Microsoft_Graph_API_Client.api('/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/sendMail').post({

                message: {
                    subject: 'Machado | Método Gerencial para Empresas - Instruções de Acesso à Plataforma',
                    body: {
                        contentType: 'HTML',
                        content: `
                            <p>Bom dia ${Cliente_PrimeiroNome},</p>
                            <p>Escrevemos do suporte da Machado | Método Gerencial para Empresas. Tudo bem?</p>
                            <p>Recentemente a Engefy contratou a nova versão de nossa Solução em Método Gerencial, para auxiliarmos no amadurecimento do Sistema de Gestão da empresa. E você foi um dos profissionais selecionados para participar do trabalho!</p>
                            <p>A Solução possui duas grandes porções:</p>
                            <p><b>• Formação em Método Gerencial:</b> acontece em nossa plataforma de ensino, de maneira online e assíncrona, durante 5 semanas. Esta é a etapa que estamos começando agora.</p>
                            <p><b>• Encontros ao Vivo:</b> posteriormente, nosso fundador (Lucas Machado) irá até a Engefy para conduzir junto a vocês o choque de Gestão na empresa, durante 2 dias.</p>
                            <p>Dito isto, compartilhamos as instruções de acesso à Formação:</p>
                            <span><b>Link:</b> <a href="https://machadogestao.com/plataforma_v2/login">https://machadogestao.com/plataforma_v2/login</a><br></span>
                            <span><b>Login:</b> ${Cliente_Email}<br></span>
                            <span><b>Senha:</b> ${Cliente_Senha}<br></span>
                            <p>*Suas credenciais de acesso são individuais e instransferíveis.</p>
                            <p>**Nossa plataforma possui várias camadas de segurança e monitoramento. Por isto, o acesso deve ser realizado exclusivamente pelo navegador <b>Microsoft Edge</b>, via laptop ou desktop com <b>sistema Windows</b>. Computadores Apple/Mac são incompatíveis com nossos sistemas.</p>
                            <p>Orientações Adicionais:</p>
                            <p>• Sua caixa personalizada com materiais impressos (apostilas, cases, documentos auxiliares, etc.) já foi enviada à Engefy. Favor alinhar recebimento junto ao Luan Mannes.</p>
                            <p>• A meta de início dos estudos será encaminhada pelo grupo do WhatsApp ainda hoje, logo após a reunião de kick-off. Importante: sugerimos que você tenha sua caixa de materiais impressos em mãos antes de iniciar os estudos.</p>
                            <p>• Porém sugerimos também que você faça seu primeiro login, incluindo cadastramento no sistema de reconhecimento facial e familiarização inicial com a plataforma desde já.</p>
                            <p>Em caso de dúvidas / dificuldades:</p>
                            <p>• <b>Técnicas</b> (relacionadas ao acesso à plataforma ou eventuais bugs): sinalize para nós via inbox ao WhatsApp +55 41 99679 9092. Iremos auxiliá-lo(a) prontamente.</p>
                            <p>• <b>Conceituais</b> (relacionadas à compreensão ou aplicação do Método Gerencial no dia a dia da Engefy): anote em seus materiais impressos de forma organizada e traga nos Encontros ao Vivo para discussão conjunta.</p>
                            <p>Qualquer dúvida ou insegurança, sempre à disposição.</p>
                            <p>Atenciosamente,</p>
                            <p><img src="https://plataforma-backend-v3.azurewebsites.net/img/ASSINATURA_E-MAIL.jpg" width="600" /></p>
                        `
                    },
                    toRecipients: [{ emailAddress: { address: Cliente_Email } }]
                }

            });

            await new Promise(resolve => setTimeout(resolve, 2000));

        }

    }

    setTimeout(Envia_Email_Clientes, 1000);

});

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////// PROCESSAMENTO DA PLATAFORMA_v2 ///////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

app.post('/plataforma_v2/login-FaceID', async (req, res) => {
    
    let { Usuário_Login, Usuário_Senha } = req.body;

    let BD_Plataforma;
    try { BD_Plataforma = await retry(() => Microsoft_Graph_API_Client.api('/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECSBYCZNYGEWFFDLEOZ36WI2PDWO/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows').get()) }
    catch (err) { return res.status(500).json({ error: 'Erro_001' }) }
    
    for (let i = 0; i < BD_Plataforma.value.length; i++) {
        let LinhaVerificada = BD_Plataforma.value[i].values[0];
        if (Usuário_Login === LinhaVerificada[2] && Usuário_Senha === LinhaVerificada[3].toString()) { 
            return res.status(200).json({ IndexVerificado: i, Usuário_Status_FaceID: LinhaVerificada[4], Usuário_Foto_Cadastrada: LinhaVerificada[5], Usuário_PrazoAcesso: ConverteData(LinhaVerificada[6]), Usuário_Status_Login: LinhaVerificada[7] })
        }
    }

    return res.status(401).json({error: 'credenciais_inválidas'});

});

app.post('/plataforma_v2/CadastroFoto_e_FaceID', multer().single('file'), async (req, res) => {
    
    let IndexVerificado = req.body.IndexVerificado;
    let FotoReferência = req.file.buffer;
        
    try { await retry(() => Microsoft_Graph_API_Client.api(`/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/root:/2. ENTREGA/1. CONTROLAR PLATAFORMA/PG - FOTOS DE REFERÊNCIA/${IndexVerificado}.jpg:/content`).put(FotoReferência))}
    catch (err) { return res.status(500).json({ error: 'Erro_002' }) }  

    try { await retry(() => Microsoft_Graph_API_Client.api('/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECSBYCZNYGEWFFDLEOZ36WI2PDWO/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows/itemAt(index=' + IndexVerificado + ')').update({values: [[null, null, null, null, null, 'Sim', null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null]]}))}
    catch (err) { return res.status(500).json({ error: 'Erro_003' }) }
        
    let Azure_Face_API_LivenessSession;
    try { Azure_Face_API_LivenessSession = await retry(() => Azure_Face_API_Client.path("/detectLivenessWithVerify-sessions").post({ contentType: "multipart/form-data", body: [{ name: "VerifyImage", body: FotoReferência }, { name: "livenessOperationMode", body: "Passive" }, { name: "deviceCorrelationId", body: uuidv4() }] })) }
    catch (err) { return res.status(500).json({ error: 'Erro_004' }) }

    let Azure_Face_API_LivenessSession_authToken = Azure_Face_API_LivenessSession.body.authToken;
    let Azure_Face_API_LivenessSession_sessionID = Azure_Face_API_LivenessSession.body.sessionId;

    return res.status(200).json({ Azure_Face_API_LivenessSession_authToken, Azure_Face_API_LivenessSession_sessionID });

});

app.post('/plataforma_v2/FaceID', async (req, res) => {

    let { IndexVerificado } = req.body;
    
    let FotoReferência;
    try { FotoReferência = await retry(() => Microsoft_Graph_API_Client.api(`/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/root:/2. ENTREGA/1. CONTROLAR PLATAFORMA/PG - FOTOS DE REFERÊNCIA/${IndexVerificado}.jpg:/content`).get())}
    catch (err) { return res.status(500).json({ error: 'Erro_005' }) }

    let Azure_Face_API_LivenessSession;
    try { Azure_Face_API_LivenessSession = await retry(() => Azure_Face_API_Client.path("/detectLivenessWithVerify-sessions").post({ contentType: "multipart/form-data", body: [{name: "VerifyImage", body: FotoReferência}, {name: "livenessOperationMode", body: "Passive"}, {name: "deviceCorrelationId", body: uuidv4()}]}))}
    catch (err) { return res.status(500).json({ error: 'Erro_004' }) }

    let Azure_Face_API_LivenessSession_authToken = Azure_Face_API_LivenessSession.body.authToken;
    let Azure_Face_API_LivenessSession_sessionID = Azure_Face_API_LivenessSession.body.sessionId;

    return res.status(200).json({ Azure_Face_API_LivenessSession_authToken, Azure_Face_API_LivenessSession_sessionID });

});

app.get('/plataforma_v2/FaceID_resultado/:Azure_Face_API_LivenessSession_sessionID', async (req, res) => {

    let Azure_Face_API_LivenessSession_sessionID = req.params.Azure_Face_API_LivenessSession_sessionID;

    let Azure_Face_API_LivenessSession;
    try { Azure_Face_API_LivenessSession = await retry(() => Azure_Face_API_Client.path('/detectLivenessWithVerify-sessions/{sessionId}', Azure_Face_API_LivenessSession_sessionID).get()) }
    catch (err) { return res.status(500).json({ error: 'Erro_007' }) }
    
    let Azure_Face_API_LivenessSession_LivenessDecision = Azure_Face_API_LivenessSession.body.results.attempts[0].result.livenessDecision;
    let Azure_Face_API_LivenessSession_MatchConfidence = Azure_Face_API_LivenessSession.body.results.attempts[0].result.verifyResult.matchConfidence;
    let Azure_Face_API_LivenessSession_MatchDecision = Azure_Face_API_LivenessSession.body.results.attempts[0].result.verifyResult.isIdentical;

    return res.status(200).json({ Azure_Face_API_LivenessSession_LivenessDecision, Azure_Face_API_LivenessSession_MatchConfidence, Azure_Face_API_LivenessSession_MatchDecision });

});

app.post('/plataforma_v2/refresh', async (req, res) => {
    
    let { IndexVerificado } = req.body;
    
    let BD_Plataforma;
    try { BD_Plataforma = await retry(() => Microsoft_Graph_API_Client.api('/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECSBYCZNYGEWFFDLEOZ36WI2PDWO/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows').get()) }
    catch (err) { return res.status(500).json({ error: 'Erro_001' }) }

    let Usuário_NomeCompleto = BD_Plataforma.value[IndexVerificado].values[0][0];
    let Usuário_PrimeiroNome = BD_Plataforma.value[IndexVerificado].values[0][1];
    let Usuário_Email = BD_Plataforma.value[IndexVerificado].values[0][2];
    let Usuário_PrazoAcesso = ConverteData(BD_Plataforma.value[IndexVerificado].values[0][6]);
    let Usuário_Status_Login = BD_Plataforma.value[IndexVerificado].values[0][7];
    let Usuário_Formação_NúmeroTópicosConcluídos = BD_Plataforma.value[IndexVerificado].values[0][8];
    let Usuário_Formação_NotaMódulo1 = BD_Plataforma.value[IndexVerificado].values[0][10];
    let Usuário_Formação_NotaMódulo2 = BD_Plataforma.value[IndexVerificado].values[0][11];
    let Usuário_Formação_NotaMódulo3 = BD_Plataforma.value[IndexVerificado].values[0][12];
    let Usuário_Formação_NotaMódulo4 = BD_Plataforma.value[IndexVerificado].values[0][13];
    let Usuário_Formação_NotaMódulo5 = BD_Plataforma.value[IndexVerificado].values[0][14];
    let Usuário_Formação_NotaMódulo6 = BD_Plataforma.value[IndexVerificado].values[0][15];
    let Usuário_Formação_NotaMódulo7 = BD_Plataforma.value[IndexVerificado].values[0][16];
    let Usuário_Formação_NotaMódulo8 = BD_Plataforma.value[IndexVerificado].values[0][17];
    let Usuário_Formação_NotaMódulo9 = BD_Plataforma.value[IndexVerificado].values[0][18];
    let Usuário_Formação_NotaMódulo10 = BD_Plataforma.value[IndexVerificado].values[0][19];
    let Usuário_Formação_NotaAcumulado = BD_Plataforma.value[IndexVerificado].values[0][20];
    let Usuário_Formação_CertificadoID = BD_Plataforma.value[IndexVerificado].values[0][21];
                    
    return res.status(200).json({ Usuário_NomeCompleto, Usuário_PrimeiroNome, Usuário_Email, Usuário_PrazoAcesso, Usuário_Status_Login, Usuário_Formação_NúmeroTópicosConcluídos, Usuário_Formação_NotaMódulo1, Usuário_Formação_NotaMódulo2, Usuário_Formação_NotaMódulo3, Usuário_Formação_NotaMódulo4, Usuário_Formação_NotaMódulo5, Usuário_Formação_NotaMódulo6, Usuário_Formação_NotaMódulo7, Usuário_Formação_NotaMódulo8, Usuário_Formação_NotaMódulo9, Usuário_Formação_NotaMódulo10, Usuário_Formação_NotaAcumulado, Usuário_Formação_CertificadoID });

});

app.post('/plataforma_v2/updates', async (req,res) => {
    
    let { TipoAtualização, IndexVerificado, NúmeroTópicosConcluídos, NúmeroMódulo, NotaTeste } = req.body;

    let DadosaInserir = new Array(22).fill(null);
    DadosaInserir[8] = NúmeroTópicosConcluídos;

    if(TipoAtualização === 'NúmeroTópicosConcluídos-e-NotaTeste'){ DadosaInserir[NúmeroMódulo + 9] = NotaTeste; }

    try { await retry(() => Microsoft_Graph_API_Client.api(`/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECSBYCZNYGEWFFDLEOZ36WI2PDWO/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows/itemAt(index=${IndexVerificado})`).update({ values: [DadosaInserir] })) }
    catch (err) { return res.status(500).json({ error: 'Erro_008' }) }

    return res.status(200).json({});

});

app.post('/plataforma_v2/processa-feedback', async (req, res) => {

    let { IndexVerificado, NúmeroTópicosConcluídos, Usuário_NomeCompleto, Usuário_Email, Feedback_DataPreenchimento, NúmeroMódulo, Feedback_TamanhoMódulo, Feedback_QualidadeConteúdo, Feedback_QualidadePlataforma, Feedback_QualidadeMateriaisImpressos, Feedback_Comentários } = req.body;

    try { await retry(() => Microsoft_Graph_API_Client.api('/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECSBYCZNYGEWFFDLEOZ36WI2PDWO/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows/itemAt(index=' + IndexVerificado + ')').update({ values: [[null, null, null, null, null, null, null, null, NúmeroTópicosConcluídos, null, null, null, null, null, null, null, null, null, null, null, null, null]] })) }
    catch (err) { return res.status(500).json({ error: 'Erro_008' }) }
    
    try { await retry(() => Microsoft_Graph_API_Client.api('/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECXO7I5R6LKLXJD3VWXORUAF7J37/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows/add').post({ values: [[Usuário_NomeCompleto, Usuário_Email, Feedback_DataPreenchimento, NúmeroMódulo, Feedback_TamanhoMódulo, Feedback_QualidadeConteúdo, Feedback_QualidadePlataforma, Feedback_QualidadeMateriaisImpressos, Feedback_Comentários]] })) }
    catch (err) { return res.status(500).json({ error: 'Erro_009' }) }

    return res.status(200).json({});

});

app.get('/ezdrm-playready-authorization-url', (req, res) => {
  
  const token = req.query.token || "";
  const customData = req.query.CustomData || "";
  const response = "p1=5&p2=&p3=&p4=1&p5=0&p6=1&p7=0&p8=0" + "&token=" + encodeURIComponent(token) + "&CustomData=" + encodeURIComponent(customData);
  res.set("Content-Type", "text/html");
  res.status(200).send(response);

});

app.post('/plataforma_v2/statusreport', async (req, res) => {

    let { linha_inicial, linha_final } = req.body;
    
    let BD_Plataforma;
    try { BD_Plataforma = await retry(() => Microsoft_Graph_API_Client.api('/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECSBYCZNYGEWFFDLEOZ36WI2PDWO/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows').get()) }
    catch (err) { return res.status(500).json({ error: 'Erro_001' }) }

    const Dados_Extraídos_BD_Plataforma = BD_Plataforma.value.slice(linha_inicial, linha_final + 1).map(({ values }) => [ values[0][0], values[0][8], ...values[0].slice(10, 22) ]);

    return res.status(200).json({ Dados_Extraídos_BD_Plataforma });

});