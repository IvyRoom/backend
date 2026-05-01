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
    let Linha_Inicial = 23;
    let Linha_Final = 24;

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
                            <p>Boa tarde ${Cliente_PrimeiroNome},</p>
                            <p>Escrevemos do suporte da Machado | Método Gerencial para Empresas. Tudo bem?</p>
                            <p>Recentemente, a Sion contratou a nova versão de nossa Solução em Método Gerencial, para auxiliarmos no amadurecimento do Sistema de Gestão da empresa. E você foi um dos profissionais selecionados para participar do trabalho!</p>
                            <p>A Solução possui duas grandes porções:</p>
                            <p><b>• Formação em Método Gerencial:</b> acontece em nossa plataforma de ensino, de maneira online e assíncrona, durante 5 a 10 semanas. Esta é a etapa que estamos começando agora.</p>
                            <p><b>• Encontros ao Vivo:</b> posteriormente, nosso fundador (Lucas Machado) irá até a Sion para conduzir junto a vocês o choque de Gestão na empresa, durante 3 dias.</p>
                            <p>Dito isto, compartilhamos as instruções de acesso à Formação:</p>
                            <span><b>Link:</b> <a href="https://machadogestao.com/plataforma_v2/login">https://machadogestao.com/plataforma_v2/login</a><br></span>
                            <span><b>Login:</b> ${Cliente_Email}<br></span>
                            <span><b>Senha:</b> ${Cliente_Senha}<br></span>
                            <p>*Suas credenciais de acesso são individuais e instransferíveis.</p>
                            <p>**Nossa plataforma possui várias camadas de segurança e monitoramento. Por isto, o acesso deve ser realizado exclusivamente pelo navegador <b>Microsoft Edge</b>, via laptop ou desktop.</p>
                            <p>A meta de início dos estudos será encaminhada pelo grupo do WhatsApp assim que os materiais impressos de vocês chegarem à Sion (data prevista: quarta, 25/mar/2026). <b>Sugerimos fortemente que você aguarde a chegada dos materiais para avançar nos estudos.</b></p>
                            <p>Porém, <b>sugerimos também que você já faça seu primeiro login na plataforma</b>, incluindo cadastramento no sistema de reconhecimento facial e familiarização inicial com nossos sistemas.</p>
                            <p>Observações Importantes:</p>
                            <p>• Como esta é uma versão nova de nosso serviço, caso você encontre qualquer dificuldade de acesso ou observe eventuais falhas/bugs, sinalize para nós via inbox ao WhatsApp +55 41 99679 9092. Iremos auxiliá-lo(a) prontamente.</p>
                            <p>• Além disso, se tiver dúvidas sobre a estrutura do serviço em si ou sobre as metas de estudos semanais, encaminhe-as ao grupo de WhatsApp da turma.</p>
                            <p>Qualquer dúvida ou insegurança, sempre à disposição.</p>
                            <p>Atenciosamente,</p>
                            <p><img src="https://plataforma-backend-v3.azurewebsites.net/img/ASSINATURA_E-MAIL.jpg" width="600" /></p>
                        `
                    },
                    toRecipients: [{ emailAddress: { address: 'contato@machadogestao.com' } }]
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