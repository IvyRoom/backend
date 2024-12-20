///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////// IMPORTAÇÃO DE BIBLIOTECAS E CRIAÇÃO DE FUNÇÕES /////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////
// Importa todas as bibliotecas, faz as configurações iniciais necessárias
// e cria as funções auxiliares.
////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////

////////////////////////////////////////////////////////////////////////////////////////
// Importa a biblioteca para comunicação com variáveis de ambiente.

const dotenv = require('dotenv');
dotenv.config();

//////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Importa a biblioteca para HTTP Posts, a Cors para receber sinais inclusive do Localhost, e cria o terminal.

const express = require('express');
const cors = require('cors');
const app = express();
const port = process.env.PORT || 3000;
app.use(cors());
app.use(express.json());
app.listen(port);

////////////////////////////////////////////////////////////////////////////////////////
// Importa a biblioteca para criação de arquivos iCalendar.

const ical = require('ical-generator');
const { ICalCalendar } = require('ical-generator');

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Importa as bibliotecas de comunicação com o Microsoft Graph API e cria a função para renovação do acesso.

const { Client } = require('@microsoft/microsoft-graph-client');
const { ConfidentialClientApplication } = require('@azure/msal-node');
var Microsoft_Graph_API_Client;

async function Conecta_ao_Microsoft_Graph_API() {
    const cca = new ConfidentialClientApplication({ auth: { clientId: process.env.CLIENT_ID, authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`, clientSecret: process.env.CLIENT_SECRET } });
    response = await cca.acquireTokenByClientCredential({scopes: ['https://graph.microsoft.com/.default']});
    Microsoft_Graph_API_Client = Client.init({authProvider:(done)=>{done(null, response.accessToken)}});
    
    //Chama a função novamente 5 minutos antes do Access Token expirar.
    setTimeout(Conecta_ao_Microsoft_Graph_API, new Date(response.expiresOn) - new Date() - 5 * 60 * 1000);
}

// Faz a primeira chamada para gerar o Microsoft_Graph_API_Client;
Conecta_ao_Microsoft_Graph_API();

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Cria função que transforma datas no formato Excel em datas no formato DD/MMM/AAAA.

function ConverteData(DataExcel) {
    const date = new Date((DataExcel - 25569) * 86400 * 1000);
    return date.toLocaleDateString('pt-BR', { day: '2-digit', month: 'short', year: 'numeric' }).replace(/\bde\b|\./g, '').replace(/\s+/g, '/');
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Cria função que transforma datas no formato JavaScritp em datas no formato Excel (Horário de Brasília).

function ConverteData2(DataJavaScript) {
    return DataJavaScript.toLocaleString('pt-BR', { timeZone: 'America/Sao_Paulo', day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }).replace(',', '');
}


///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////// PROCESSAMENTO DA LANDING PAGE /////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////
// Declara as variáveis mestras.
////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////

var Lead_NomeCompleto;
var Lead_PrimeiroNome;
var Lead_Email;

////////////////////////////////////////////////////////////////////////////////////////
// Serve as imagens estáticas da pasta /img.
////////////////////////////////////////////////////////////////////////////////////////
app.use('/img', express.static('img'));

////////////////////////////////////////////////////////////////////////////////////////
// Endpoint que processa submissão do formulário de cadastro.
////////////////////////////////////////////////////////////////////////////////////////

app.post('/landingpage/cadastro', async (req, res) => {
    
    var { NomeCompleto, Email } = req.body;
    Lead_NomeCompleto = NomeCompleto;
    Lead_PrimeiroNome = Lead_NomeCompleto.split(" ")[0];
    Lead_Email = Email;

    // ////////////////////////////////////////////////////////////////////////////////////////
    // // Cria o evento iCalendar, com alerta de 3 horas antes do início do evento.

    // const cal = new ICalCalendar({ domain: 'ivyroom.com.br', prodId: { company: 'Ivy | Escola de Gestão', product: 'Preparatório em Gestão Generalista', language: 'PT-BR' } });
    // const event = cal.createEvent({
    //     start: new Date(Date.UTC(2024, 10, 29, 3, 0, 0)), // 29/nov/2024, 00:00 BRT
    //     end: new Date(Date.UTC(2024, 10, 30, 2, 59, 0)), // 29/nov/2024, 23:59 BRT
    //     summary: 'Ivy - Turma de Black Friday',
    //     description: 'A turma abre 28/nov às 23:59, no link https://ivygestao.com/',
    //     uid: `${new Date().getTime()}@ivyroom.com.br`,
    //     stamp: new Date()
    // });

    // event.createAlarm({
    //     type: 'display',
    //     trigger: 3 * 60 * 60 * 1000,
    //     description: 'Ivy: Turma de Black Friday - Abre em 3 horas.'
    // });

    ////////////////////////////////////////////////////////////////////////////////////////
    // Envia o e-mail de confirmação de cadastro ao lead, com o evento iCalendar anexado.

    if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

    await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/sendMail').post({
        message: {
            subject: 'Ivy - Preparatório em Gestão Generalista: Finalize seu Cadastro',
            body: {
                contentType: 'HTML',
                content: `
                    <p>Olá ${Lead_PrimeiroNome},</p>
                    <p>Para finalizar seu cadastro na Lista de Espera para a próxima turma do Preparatório em Gestão Generalista entre no <a href="https://www.instagram.com/channel/AbaebGO_wVnsawoW/" target="_blank">Ivy Connecta</a>, nosso canal de comunicação oficial no Instagram.</p>
                    <p>Qualquer dúvida, entre em contato. Sempre à disposição.</p>
                    <p>Atenciosamente,</p>
                    <p><img src="https://plataforma-backend-v3.azurewebsites.net/img/ASSINATURA_E-MAIL.png"/></p>
                `
            },
            toRecipients: [{ emailAddress: { address: Lead_Email } }]
            // attachments: [
            //     {
            //         "@odata.type": "#microsoft.graph.fileAttachment",
            //         name: "Ivy - Turma de Black Friday.ics",
            //         contentBytes: Buffer.from(cal.toString()).toString('base64')
            //     }
            // ]
        }
    })

    .then(emailResponse => {
          
        ////////////////////////////////////////////////////////////////////////////////////////
        // Adiciona o lead na BD - LEADS.

        Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/drive/items/0172BBJBYG24NEFOMGOJCLN5FMDILTSZTC/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{AC8C07F3-9A79-4ABD-8CE8-0C818B0EA1A7}/rows')
        
            .post({"values": [[ConverteData2(new Date()), Lead_Email, Lead_NomeCompleto, "TURMA #1 2025"]]})

            .then(() => {

                res.status(200).send();

            });

    })

    .catch(error => {
        
        res.status(400).send();

    });

});


///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////// PROCESSAMENTO DA PLATAFORMA /////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////
// Declara as variáveis mestras.
////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////

var Usuário_NomeCompleto;
var Usuário_PrimeiroNome;
var Usuário_Login;
var Usuário_Preparatório1_Status;
var Usuário_Preparatório1_PrazoAcesso;
var Usuário_Preparatório1_NúmeroTópicosConcluídos;
var Usuário_Preparatório1_NotaMódulo1;
var Usuário_Preparatório1_NotaMódulo2;
var Usuário_Preparatório1_NotaMódulo3;
var Usuário_Preparatório1_NotaMódulo4;
var Usuário_Preparatório1_NotaMódulo5;
var Usuário_Preparatório1_NotaMódulo6;
var Usuário_Preparatório1_NotaMódulo7;
var Usuário_Preparatório1_NotaAcumulado;
var Usuário_Preparatório1_CertificadoID;


////////////////////////////////////////////////////////////////////////////////////////
// Endpoint que processa submissão do formulário de login.
////////////////////////////////////////////////////////////////////////////////////////

app.post('/login', async (req, res) => {
    
    var { login, senha } = req.body;
    Usuário_Login = login;
    Usuário_Senha = senha;

    ////////////////////////////////////////////////////////////////////////////////////////
    // Obtém os dados da BD - PLATAFORMA.xlsx.

    if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();
    const BD_Plataforma = await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/drive/items/0172BBJB7TUZJNIWDVWFE2MIW7MNKHMWLL/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows').get();
    
    ////////////////////////////////////////////////////////////////////////////////////////
    // Verifica se o Login e Senha cadastrados pelo usuário estão na BD_Plataforma.

    const BD_Plataforma_Número_Linhas = BD_Plataforma.value.length;

    var LoginAutenticado = 0;
    var IndexVerificado = 0;
    var LoginVerificado;
    var SenhaVerificada;

    Verifica_Login_e_Senha();

    function Verifica_Login_e_Senha() {

        if (IndexVerificado < BD_Plataforma_Número_Linhas) {
            
            LoginVerificado = BD_Plataforma.value[IndexVerificado].values[0][2];
            SenhaVerificada = BD_Plataforma.value[IndexVerificado].values[0][3].toString();

            if(Usuário_Login === LoginVerificado) {

                if(Usuário_Senha === SenhaVerificada) {
                    
                    LoginAutenticado = 1;

                    Usuário_PrimeiroNome = BD_Plataforma.value[IndexVerificado].values[0][1];
                    Usuário_Preparatório1_Status = BD_Plataforma.value[IndexVerificado].values[0][4];
                    Usuário_Preparatório1_PrazoAcesso = ConverteData(BD_Plataforma.value[IndexVerificado].values[0][5]);
                    
                }
                
            } else {

                IndexVerificado++;
                Verifica_Login_e_Senha();

            }

        }
        
    }

    res.status(200).json({ 
        
        LoginAutenticado,
        IndexVerificado,
        Usuário_PrimeiroNome,
        Usuário_Preparatório1_Status,
        Usuário_Preparatório1_PrazoAcesso

    });

});

////////////////////////////////////////////////////////////////////////////////////////
// Endpoint que processa carregamento da aba /estudos no Frontend.
////////////////////////////////////////////////////////////////////////////////////////

app.post('/refresh', async (req, res) => {
    
    var { IndexVerificado } = req.body;

    ////////////////////////////////////////////////////////////////////////////////////////
    // Obtém os dados da BD - PLATAFORMA.xlsx.
    if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();
    const BD_Plataforma = await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/drive/items/0172BBJB7TUZJNIWDVWFE2MIW7MNKHMWLL/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows').get();

    Usuário_NomeCompleto = BD_Plataforma.value[IndexVerificado].values[0][0];
    Usuário_PrimeiroNome = BD_Plataforma.value[IndexVerificado].values[0][1];
    Usuário_Preparatório1_Status = BD_Plataforma.value[IndexVerificado].values[0][4];
    Usuário_Preparatório1_PrazoAcesso = ConverteData(BD_Plataforma.value[IndexVerificado].values[0][5]);
    Usuário_Preparatório1_NúmeroTópicosConcluídos = BD_Plataforma.value[IndexVerificado].values[0][6];
    Usuário_Preparatório1_NotaMódulo1 = BD_Plataforma.value[IndexVerificado].values[0][8];
    Usuário_Preparatório1_NotaMódulo2 = BD_Plataforma.value[IndexVerificado].values[0][9];
    Usuário_Preparatório1_NotaMódulo3 = BD_Plataforma.value[IndexVerificado].values[0][10];
    Usuário_Preparatório1_NotaMódulo4 = BD_Plataforma.value[IndexVerificado].values[0][11];
    Usuário_Preparatório1_NotaMódulo5 = BD_Plataforma.value[IndexVerificado].values[0][12];
    Usuário_Preparatório1_NotaMódulo6 = BD_Plataforma.value[IndexVerificado].values[0][13];
    Usuário_Preparatório1_NotaMódulo7 = BD_Plataforma.value[IndexVerificado].values[0][14];
    Usuário_Preparatório1_NotaAcumulado = BD_Plataforma.value[IndexVerificado].values[0][15];
    Usuário_Preparatório1_CertificadoID = BD_Plataforma.value[IndexVerificado].values[0][16];
    Usuário_Preparatório2_Interesse = BD_Plataforma.value[IndexVerificado].values[0][18];
                    
    res.status(200).json({ 
        
        Usuário_NomeCompleto,
        Usuário_PrimeiroNome,
        Usuário_Preparatório1_Status,
        Usuário_Preparatório1_PrazoAcesso,
        Usuário_Preparatório1_NúmeroTópicosConcluídos,
        Usuário_Preparatório1_NotaMódulo1,
        Usuário_Preparatório1_NotaMódulo2,
        Usuário_Preparatório1_NotaMódulo3,
        Usuário_Preparatório1_NotaMódulo4,
        Usuário_Preparatório1_NotaMódulo5,
        Usuário_Preparatório1_NotaMódulo6,
        Usuário_Preparatório1_NotaMódulo7,
        Usuário_Preparatório1_NotaAcumulado,
        Usuário_Preparatório1_CertificadoID,
        Usuário_Preparatório2_Interesse

    });

});


////////////////////////////////////////////////////////////////////////////////////////
// Endpoint que atualiza a BD - PLATAFORMA.
////////////////////////////////////////////////////////////////////////////////////////

app.post('/updates', async (req,res) => {
    
    var { TipoAtualização, IndexVerificado, NúmeroTópicosConcluídos, NúmeroMódulo, NotaTeste, Preparatório2_Interesse } = req.body;

    //Atualiza o Número de Tópicos Concluídos e a Nota no Teste.

    if(TipoAtualização === 'NúmeroTópicosConcluídos-e-NotaTeste'){

        if (NúmeroMódulo === 1){

            if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();
            await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/drive/items/0172BBJB7TUZJNIWDVWFE2MIW7MNKHMWLL/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows/itemAt(index=' + IndexVerificado + ')').update({values: [[null, null, null, null, null, null, NúmeroTópicosConcluídos, null, NotaTeste, null, null, null, null, null, null, null, null, null, null]]});

        } else if (NúmeroMódulo === 2){

            if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();
            await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/drive/items/0172BBJB7TUZJNIWDVWFE2MIW7MNKHMWLL/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows/itemAt(index=' + IndexVerificado + ')').update({values: [[null, null, null, null, null, null, NúmeroTópicosConcluídos, null, null, NotaTeste, null, null, null, null, null, null, null, null, null]]});

        } else if (NúmeroMódulo === 3){

            if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();
            await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/drive/items/0172BBJB7TUZJNIWDVWFE2MIW7MNKHMWLL/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows/itemAt(index=' + IndexVerificado + ')').update({values: [[null, null, null, null, null, null, NúmeroTópicosConcluídos, null, null, null, NotaTeste, null, null, null, null, null, null, null, null]]});

        } else if (NúmeroMódulo === 4){

            if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();
            await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/drive/items/0172BBJB7TUZJNIWDVWFE2MIW7MNKHMWLL/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows/itemAt(index=' + IndexVerificado + ')').update({values: [[null, null, null, null, null, null, NúmeroTópicosConcluídos, null, null, null, null, NotaTeste, null, null, null, null, null, null, null]]});

        } else if (NúmeroMódulo === 5){

            if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();
            await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/drive/items/0172BBJB7TUZJNIWDVWFE2MIW7MNKHMWLL/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows/itemAt(index=' + IndexVerificado + ')').update({values: [[null, null, null, null, null, null, NúmeroTópicosConcluídos, null, null, null, null, null, NotaTeste, null, null, null, null, null, null]]});

        } else if (NúmeroMódulo === 6){

            if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();
            await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/drive/items/0172BBJB7TUZJNIWDVWFE2MIW7MNKHMWLL/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows/itemAt(index=' + IndexVerificado + ')').update({values: [[null, null, null, null, null, null, NúmeroTópicosConcluídos, null, null, null, null, null, null, NotaTeste, null, null, null, null, null]]});

        } else {

            if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();
            await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/drive/items/0172BBJB7TUZJNIWDVWFE2MIW7MNKHMWLL/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows/itemAt(index=' + IndexVerificado + ')').update({values: [[null, null, null, null, null, null, NúmeroTópicosConcluídos, null, null, null, null, null, null, null, NotaTeste, null, null, null, null]]});

        } 

    } else if(TipoAtualização === 'NúmeroTópicosConcluídos') {

        //Atualiza só o Número de Tópicos Concluídos do Prep.
        if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();
        await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/drive/items/0172BBJB7TUZJNIWDVWFE2MIW7MNKHMWLL/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows/itemAt(index=' + IndexVerificado + ')').update({values: [[null, null, null, null, null, null, NúmeroTópicosConcluídos, null, null, null, null, null, null, null, null, null, null, null, null]]});

    } else if(TipoAtualização === 'Preparatório2_Interesse'){

        //Atualiza só o Interesse no Preparatório 2.
        if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();
        await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/drive/items/0172BBJB7TUZJNIWDVWFE2MIW7MNKHMWLL/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows/itemAt(index=' + IndexVerificado + ')').update({values: [[null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, Preparatório2_Interesse]]});

    }

    res.status(200).send();

});


///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////// PROCESSAMENTO DOS LEADS ///////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////
// Declara as variáveis mestras.
////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////

var ProcessamentoLeads_PrimeiroNome;
var ProcessamentoLeads_Email;

////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////
// Envia e-mail em escala para os leads na BD - LEADS.
////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////

app.post('/leads/envioemail', async (req,res) => {
    
    res.status(200).send();
    
    ////////////////////////////////////////////////////////////////////////////////////////
    // Puxa os dados da BD - LEADS.
    
    if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

    const BD_Leads = await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/drive/items/0172BBJBYG24NEFOMGOJCLN5FMDILTSZTC/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{AC8C07F3-9A79-4ABD-8CE8-0C818B0EA1A7}/rows').get();

    const BD_Leads_Número_Linhas = BD_Leads.value.length;

    const BD_Leads_Última_Linha = BD_Leads_Número_Linhas - 1;
    
    // Linha Atual inicial, com e-mail: 483
    for (let LinhaAtual = 483; LinhaAtual <= BD_Leads_Última_Linha; LinhaAtual++) {
        
        ProcessamentoLeads_Email = BD_Leads.value[LinhaAtual].values[0][1];
        ProcessamentoLeads_PrimeiroNome = BD_Leads.value[LinhaAtual].values[0][2].split(" ")[0];

        if (ProcessamentoLeads_Email === "-") {

        } else {

            // ////////////////////////////////////////////////////////////////////////////////////////
            // // Envia o e-mail para o lead na LinhaAtual da BD - LEADS.

            // if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

            // await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/sendMail').post({
            //     message: {
            //         subject: 'Ivy - Turma de Black Friday: Faltam 48 horas',
            //         body: {
            //             contentType: 'HTML',
            //             content: `
            //                 <p>Olá ${ProcessamentoLeads_PrimeiroNome},</p>
            //                 <p>Quem escreve é Lucas Machado, fundador da Ivy. Tudo bem?</p>
            //                 <p>Sinalizo que a <b>Turma Especial de Black Friday</b> do Preparatório em Gestão Generalista <b>abre em pouco mais de 48h</b>, no dia 28/nov (quinta-feira) às 23:59, por meio do link <a href="https://ivygestao.com/" target="_blank">https://ivygestao.com/</a>.</p>
            //                 <p>Ao entrar na turma, você irá adquirir habilidades fundamentais para o seu crescimento de carreira, como dominar o Sistema de Gestão e suas porções, o Ciclo de Melhoria de Resultados (PDCA) e o Ciclo de Estabilização de Processos (SDCA). Tudo de forma extremamente prática e guiada pela aplicação de softwares e Estudos de Caso reais. Fora isto, você receberá <b>três bônus exclusivos</b>:</p>
            //                 <p>&nbsp;&nbsp;&nbsp;&nbsp;1) <b>Caixa de materiais impressos</b> com apostilas, cases e guias de aplicação rápida do conhecimento.</p>
            //                 <p>&nbsp;&nbsp;&nbsp;&nbsp;2) <b>Bibliografias de suporte</b>. Os livros: Princípios (Ray Dalio), O Verdadeiro Poder (Falconi) e Gerenciamento da Rotina (Falconi).</p>
            //                 <p>&nbsp;&nbsp;&nbsp;&nbsp;3) <b>Encontro a portas fechadas</b> com conteúdo exclusivo, trazido da Harvard Business School, disponível aos 10 primeiros alunos.</p>
            //                 <p>Todas as informações sobre o serviço podem ser acessadas por <a href="https://ivygestao.com/" target="_blank">aqui</a>.</p>
            //                 <p>Qualquer dúvida, entre em contato. Sempre à disposição.</p>
            //                 <p>Atenciosamente,</p>
            //                 <p><img src="https://plataforma-backend-v3.azurewebsites.net/img/ASSINATURA_E-MAIL.png"/></p>
            //             `
            //         },
            //         toRecipients: [{ emailAddress: { address: ProcessamentoLeads_Email } }]
            //     }
            // })

            // ////////////////////////////////////////////////////////////////////////////////////////
            // // Envia o e-mail para o lead na LinhaAtual da BD - LEADS.

            // if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

            // await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/sendMail').post({
            //     message: {
            //         subject: 'Ivy - Turma de Black Friday: É hoje!',
            //         body: {
            //             contentType: 'HTML',
            //             content: `
            //                 <p>Bom dia ${ProcessamentoLeads_PrimeiroNome},</p>
            //                 <p>Sinalizamos que a <b>Turma Especial de Black Friday</b> do Preparatório em Gestão Generalista <b>abre hoje, dia 28/nov (quinta-feira) às 23:59</b>, por meio do nosso <b><a href="https://ivygestao.com/" target="_blank">Link da Bio</a></b>!</p>
            //                 <p>Confira todos os detalhes sobre a oferta, incluindo os bônus exclusivos, <a href="https://ivygestao.com/" target="_blank">clicando aqui</a>.</p>
            //                 <p>E se tiver dúvidas, entre em contato. Estamos sempre à disposição.</p>
            //                 <p>Atenciosamente,</p>
            //                 <p><img src="https://plataforma-backend-v3.azurewebsites.net/img/ASSINATURA_E-MAIL.png"/></p>
            //             `
            //         },
            //         toRecipients: [{ emailAddress: { address: ProcessamentoLeads_Email } }]
            //     }
            // })

            // ////////////////////////////////////////////////////////////////////////////////////////
            // // Envia o e-mail para o lead na LinhaAtual da BD - LEADS.

            // if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

            // await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/sendMail').post({
            //     message: {
            //         subject: 'Ivy - Turma de Black Friday: Faltam 3 horas!',
            //         body: {
            //             contentType: 'HTML',
            //             content: `
            //                 <p>Olá ${ProcessamentoLeads_PrimeiroNome},</p>
            //                 <p>Quem escreve é Lucas Machado, fundador da Ivy. Tudo bem?</p>
            //                 <p><b>Faltam 3 horas</b> para a abertura da <b>Turma Especial de Black Friday</b> do Preparatório em Gestão Generalista.</p>
            //                 <p>A turma <b>abre hoje, 28/nov (quinta-feira) às 23:59</b>, em nosso <b><a href="https://ivygestao.com/" target="_blank">Link da Bio</a></b>!</p>
            //                 <p>Aproveito para lhe passar três dicas que podem fazer a diferença para você ter acesso a todos os bônus ofertados (incluindo o encontro fechado para os 10 primeiros alunos):</p>
            //                 <p>&nbsp;&nbsp;&nbsp;&nbsp;1) Às 23:59, aperte Ctrl + F5 para atualizar a página direto do nosso servidor.</p>
            //                 <p>&nbsp;&nbsp;&nbsp;&nbsp;2) Preencha o checkout com calma e atenção. Tenha certeza que seus dados estão corretos.</p>
            //                 <p>&nbsp;&nbsp;&nbsp;&nbsp;3) Se tiver qualquer dúvida, mande um e-mail ou direct. Nossa equipe estará de plantão durante a madrugada.</p>
            //                 <p>Além disso, disponibilizaremos no checkout formas de pagamento flexíveis incluindo PIX, Boleto, Cartão, PIX + Cartão e Dois Cartões, e parcelamento em até 12x.</p>
            //                 <p>Todos os detalhes da oferta, incluindo os bônus exclusivos, podem ser vistos <a href="https://ivygestao.com/" target="_blank">aqui</a>.</p>
            //                 <p>Está bem?</p>
            //                 <p>Atenciosamente,</p>
            //                 <p><img src="https://plataforma-backend-v3.azurewebsites.net/img/ASSINATURA_E-MAIL.png"/></p>
            //             `
            //         },
            //         toRecipients: [{ emailAddress: { address: ProcessamentoLeads_Email } }]
            //     }
            // })

            // ////////////////////////////////////////////////////////////////////////////////////////
            // // Envia o e-mail para o lead na LinhaAtual da BD - LEADS.

            // if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

            // await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/sendMail').post({
            //     message: {
            //         subject: 'Ivy - Turma de Black Friday: Inscrições Abertas!',
            //         body: {
            //             contentType: 'HTML',
            //             content: `
            //                 <p>Boa noite ${ProcessamentoLeads_PrimeiroNome},</p>
            //                 <p>As <b>inscrições estão abertas</b> para a <b>Turma Especial de Black Friday</b> do Preparatório em Gestão Generalista, por meio do nosso <b><a href="https://ivygestao.com/" target="_blank">Link da Bio</a></b>!</p>
            //                 <p>P.S. #1. Sua trajetória rumo a cargos gerenciais começa agora. Parabéns.</p>
            //                 <p>P.S. #2. Ficamos honrados em participar deste processo. Conte com a gente.</p>
            //                 <p>Atenciosamente,</p>
            //                 <p><img src="https://plataforma-backend-v3.azurewebsites.net/img/ASSINATURA_E-MAIL.png"/></p>
            //             `
            //         },
            //         toRecipients: [{ emailAddress: { address: ProcessamentoLeads_Email } }]
            //     }
            // })

            // // ////////////////////////////////////////////////////////////////////////////////////////
            // // Envia o e-mail para o lead na LinhaAtual da BD - LEADS.

            // if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

            // await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/sendMail').post({
            //     message: {
            //         subject: 'Ivy - Reflexão sobre Investimentos',
            //         body: {
            //             contentType: 'HTML',
            //             content: `
            //                 <p><b>REFLEXÃO SOBRE INVESTIMENTOS:</b></p>
            //                 <p>-----</p>
            //                 <p>Antes dos 30 o foco não é juntar dinheiro.</p>
            //                 <p>Ao invés de investir os R$500 que te sobram no mês em ações, por exemplo...</p>
            //                 <p>Invista em adquirir competências que tripliquem seu salário em 2 ou 3 anos.</p>
            //                 <p>Pois R$500 rendendo 10% / ano não mudam nada na sua vida. Mas ganhar 3x mais muda.</p>
            //                 <p>-----</p>
            //                 <p>Se isto fizer sentido para você, a turma de Black Friday do Preparatório em Gestão Generalista está com <a href="https://ivygestao.com/" target="_blank">inscrições abertas</a>. Esta é uma excelente oportunidade para você começar a caminhar nesta direção.</p>
            //                 <p>Saiba mais <a href="https://ivygestao.com/" target="_blank">clicando aqui</a>.</p>
            //                 <p>Atenciosamente,</p>
            //                 <p><img src="https://plataforma-backend-v3.azurewebsites.net/img/ASSINATURA_E-MAIL.png"/></p>
            //             `
            //         },
            //         toRecipients: [{ emailAddress: { address: ProcessamentoLeads_Email } }]
            //     }
            // })

            // ////////////////////////////////////////////////////////////////////////////////////////
            // // Envia o e-mail para o lead na LinhaAtual da BD - LEADS.

            // if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

            // await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/sendMail').post({
            //     message: {
            //         subject: 'Ivy - Reflexão sobre Carreira',
            //         body: {
            //             contentType: 'HTML',
            //             content: `
            //                 <p><b>REFLEXÃO SOBRE CARREIRA:</b></p>
            //                 <p>------</p>
            //                 <p>Todos nós pagamos pela lições que nós temos que aprender na vida profissional ou com <b>tempo</b>, ou com <b>dinheiro</b>.</p>
            //                 <p>E, dentre estas duas opções, nós escolhemos utilizar aquilo que a gente valoriza menos.</p>
            //                 <p>É por isto que a maioria das pessoas fica tentando sozinha, investindo só o próprio tempo, na tentativa e erro.</p>
            //                 <p>Só que o segredo é <b>investir as duas coisas</b> (tempo e dinheiro) em livros, cursos, mentorias, workshops e seminários que te permitam <u><b>aprender com experts</b></u>. Que te permitam aprender com pessoas que já tenham tido muito resultado fazendo aquilo que você quer fazer.</p>
            //                 <p>Assim você poupa anos (ou décadas!) de trabalho improdutivo, simplesmente por você aprender com quem já sabe fazer.</p>
            //                 <p>------</p>
            //                 <p>E quando o assunto é Gestão e Carreira, sem dúvidas nós podemos cumprir este papel para você.</p>
            //                 <p>Por isto lembramos que a Turma de Black Friday do Preparatório em Gestão Generalista está com <a href="https://ivygestao.com/" target="_blank">inscrições abertas</a>.</p>
            //                 <p>Saiba mais <a href="https://ivygestao.com/" target="_blank">clicando aqui</a>.</p>
            //                 <p>Atenciosamente,</p>
            //                 <p><img src="https://plataforma-backend-v3.azurewebsites.net/img/ASSINATURA_E-MAIL.png"/></p>
            //             `
            //         },
            //         toRecipients: [{ emailAddress: { address: ProcessamentoLeads_Email } }]
            //     }
            // })

            // // ////////////////////////////////////////////////////////////////////////////////////////
            // // Envia o e-mail para o lead na LinhaAtual da BD - LEADS.

            // if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

            // await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/sendMail').post({
            //     message: {
            //         subject: 'Ivy - Reflexão sobre o Mercado',
            //         body: {
            //             contentType: 'HTML',
            //             content: `
            //                 <p><b>REFLEXÃO SOBRE O MERCADO:</b></p>
            //                 <p>------</p>
            //                 <p>Você sabe qual é a maior força da natureza?</p>
            //                 <p>A <b>Lei de Mercado</b>. Ou seja, a Lei de <b>Oferta e Demanda</b>.</p>
            //                 <p>Quando observamos esta lei aplicada as nossas vidas profissionais, vemos que o mercado de trabalho está cheio de oportunidades.</p>
            //                 <p>Por exemplo...</p>
            //                 <p>Recentemente nós fizemos um estudo em nossos stories, com pouco mais de 10.000 pessoas.</p>
            //                 <p>Ali, nós colocamos em sequência as 10 perguntas mais básicas que existem sobre Gestão, com alternativas de múltipla escolha.</p>
            //                 <p>Eram perguntas como "Qual é a Equação Fundamental da Gestão?" ou "Qual é o Primeiro Princípio do Método Gerencial?".</p>
            //                 <p>O resultado? 99,6% das pessoas não acertaram nem a metade destas perguntas.</p>
            //                 <p>Isto mostra a <b>absoluta escassez</b> de pessoas que dominam Gestão no mercado de trabalho brasileiro.</p>
            //                 <p>Por outro lado, existem cerca de 800 mil empresas de médio e grande porte em nosso país.</p>
            //                 <p>E a infinita maioria delas não só precisa melhorar urgentemente a própria gestão, como tem a plena consciência disto.</p>
            //                 <p>Em outras palavras, no Brasil existe também <b>enorme demanda</b> por gente que saiba Gestão de verdade (gente que não só saiba teorias soltas e desconexas, mas que domine o software, a ferramenta, a implementação... A prática!).</p>
            //                 <p>E isto é uma excelente notícia sabe para quem?</p>
            //                 <p>Para você!</p>
            //                 <p>Pois onde há escassez de gente bem preparada e enorme demanda por estas pessoas... Há <b>oportunidade</b>.</p>
            //                 <p>Oportunidade para crescer na carreira, se destacar e disputar melhores cargos e salários.</p>
            //                 <p>------</p>
            //                 <p>Com isto, lembramos que a Turma de Black Friday do Preparatório em Gestão Generalista está com <a href="https://ivygestao.com/" target="_blank">inscrições abertas</a> <b>só até quarta-feira (04/dez) às 23:59</b>.</p>
            //                 <p>Saiba mais <a href="https://ivygestao.com/" target="_blank">clicando aqui</a>.</p>
            //                 <p>Atenciosamente,</p>
            //                 <p><img src="https://plataforma-backend-v3.azurewebsites.net/img/ASSINATURA_E-MAIL.png"/></p>
            //             `
            //         },
            //         toRecipients: [{ emailAddress: { address: ProcessamentoLeads_Email } }]
            //     }
                
            // })

            // // ////////////////////////////////////////////////////////////////////////////////////////
            // // Envia o e-mail para o lead na LinhaAtual da BD - LEADS.

            // if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

            // await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/sendMail').post({
            //     message: {
            //         subject: 'Ivy - Black Friday: Inscrições Encerram AMANHÃ!',
            //         body: {
            //             contentType: 'HTML',
            //             content: `
            //                 <p>Boa tarde ${ProcessamentoLeads_PrimeiroNome},</p>
            //                 <p>As <a href="https://ivygestao.com/" target="_blank">inscrições</a> para a Turma de Black Friday do Preparatório em Gestão Generalista <u><b>encerram amanhã</b></u>, quarta-feira (04/dez) às 23:59.</p>
            //                 <p>Lembramos que esta é a <b>última turma de 2024</b>.</p>
            //                 <p>Não perca a oportunidade de aprender Método Gerencial, a habilidade mais importante para o seu crescimento no mundo corporativo, de forma prática, objetiva e guiada pelo uso de softwares e cases reais.</p>
            //                 <p>Se inscreva <a href="https://ivygestao.com/" target="_blank">clicando aqui</a>.</p>
            //                 <p>Atenciosamente,</p>
            //                 <p><img src="https://plataforma-backend-v3.azurewebsites.net/img/ASSINATURA_E-MAIL.png"/></p>
            //             `
            //         },
            //         toRecipients: [{ emailAddress: { address: ProcessamentoLeads_Email } }]
            //     }
                
            // })

            // ////////////////////////////////////////////////////////////////////////////////////////
            // // Envia o e-mail para o lead na LinhaAtual da BD - LEADS.

            // if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

            // await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/sendMail').post({
            //     message: {
            //         subject: 'Ivy - Black Friday: Inscrições Encerram HOJE!',
            //         body: {
            //             contentType: 'HTML',
            //             content: `
            //                 <p>Olá ${ProcessamentoLeads_PrimeiroNome},</p>
            //                 <p>As <a href="https://ivygestao.com/" target="_blank">inscrições</a> para a Turma de Black Friday do Preparatório em Gestão Generalista <u><b>encerram hoje</b></u>, quarta-feira (04/dez) às 23:59.</p>
            //                 <p>Importante: esta é a <b>última turma de 2024</b>.</p>
            //                 <p>Se você deseja aprender Gestão Generalista (a competência central ao seu crescimento de carreira) de forma prática, objetiva e voltada à aplicação de softwares e cases reais, se inscreva <a href="https://ivygestao.com/" target="_blank">clicando aqui</a>.</p>
            //                 <p>Qualquer dúvida ou insegurança, à disposição.</p>
            //                 <p>Atenciosamente,</p>
            //                 <p><img src="https://plataforma-backend-v3.azurewebsites.net/img/ASSINATURA_E-MAIL.png"/></p>
            //             `
            //         },
            //         toRecipients: [{ emailAddress: { address: ProcessamentoLeads_Email } }]
            //     }
                
            // })

            // ////////////////////////////////////////////////////////////////////////////////////////
            // // Envia o e-mail para o lead na LinhaAtual da BD - LEADS.

            if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

            await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/sendMail').post({
                message: {
                    subject: 'Ivy - Black Friday: Última Chamada',
                    body: {
                        contentType: 'HTML',
                        content: `
                            <p>Olá ${ProcessamentoLeads_PrimeiroNome},</p>
                            <p>As <a href="https://ivygestao.com/" target="_blank">inscrições</a> para a Turma de Black Friday do Preparatório em Gestão Generalista <u><b>encerram dentro de 1h</b></u>, hoje, quarta-feira (04/dez) às 23:59.</p>
                            <p>Se inscreva <a href="https://ivygestao.com/" target="_blank">clicando aqui</a>.</p>
                            <p>Atenciosamente,</p>
                            <p><img src="https://plataforma-backend-v3.azurewebsites.net/img/ASSINATURA_E-MAIL.png"/></p>
                        `
                    },
                    toRecipients: [{ emailAddress: { address: 'contato@ivyroom.com.br' } }]
                }
                
            })

        }

    }

});


///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////// COMUNICAÇÃO COM ALUNOS ///////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////
// Declara as variáveis mestras.
////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////

var Aluno_PrimeiroNome;
var Aluno_Email;

////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////
// Envia e-mail em escala para os leads na BD - ALUNOS.
////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////

app.post('/alunos/envioemail', async (req,res) => {
    
    res.status(200).send();
    
    ////////////////////////////////////////////////////////////////////////////////////////
    // Puxa os dados da BD - ALUNOS.
    
    if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

    const BD_Alunos = await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/drive/items/0172BBJB3JXEEKH4PQDFEYODH27M4CPH77/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows').get();

    const BD_Alunos_Número_Linhas = BD_Alunos.value.length;

    const BD_Alunos_Última_Linha = BD_Alunos_Número_Linhas - 1;
    
    // Linha Atual inicial, com e-mail: 204.
    for (let LinhaAtual = 204; LinhaAtual <= BD_Alunos_Última_Linha; LinhaAtual++) {
        
//         ////////////////////////////////////////////////////////////////////////////////////////
//         // Cria o evento iCalendar, com alerta de 3 horas antes do início do evento.

//         const cal = new ICalCalendar({ domain: 'ivyroom.com.br', prodId: { company: 'Ivy | Escola de Gestão', product: 'Encontro Exclusivo: Black Friday', language: 'PT-BR' } });
//         const event = cal.createEvent({
//             start: new Date(Date.UTC(2025, 1, 8, 12, 0, 0)), // 08/fev/2025, 09:00 BRT
//             end: new Date(Date.UTC(2025, 1, 8, 16, 0, 0)), // 08/fev/2025, 12:59 BRT
//             summary: 'Ivy - Encontro Exclusivo: Black Friday',
//             description: `
// Neste encontro, serão discutidas lições gerenciais trazidas pelo Lucas de sua experiência recente na Harvard Business School, divididas em três grandes temas:

//     a) Interface de conhecimento e conexões lógicas entre Ger. Estratégico e Microeconomia.

//     b) Interface de conhecimento e conexões lógicas entre Ger. Tático, Contabilidade e Finanças Corporativas.

//     c) Uso de ferramentas avançadas e programação (VS Code, Git e GitHub, Microsoft Azure) no Ger. Inovações e Ger. Rotina.

// O encontro acontecerá no sábado, dia 08/fev/2025, entre 9h e 13h, via Microsoft Teams, por meio deste link:

// https://teams.microsoft.com/l/meetup-join/19%3ameeting_NjJmMGJjOGMtMDdiMS00NjZkLTlkYzUtNzc3ZjhjYTY2ZGU3%40thread.v2/0?context=%7b%22Tid%22%3a%2249342d16-0605-4267-b540-d1fe7756dbac%22%2c%22Oid%22%3a%22b4a93dcf-5946-4cb2-8368-5db4d242a236%22%7d`,
//             uid: `${new Date().getTime()}@ivyroom.com.br`,
//             stamp: new Date()
//         });

//         event.createAlarm({
//             type: 'display',
//             trigger: 1 * 60 * 60 * 1000,
//             description: 'Ivy - Encontro Exclusivo: Black Friday - Inicia em 1 hora.'
//         });
        
        Aluno_Email = BD_Alunos.value[LinhaAtual].values[0][2];
        Aluno_PrimeiroNome = BD_Alunos.value[LinhaAtual].values[0][1].split(" ")[0];

        if (Aluno_Email === "-") {

        } else {

            // // ////////////////////////////////////////////////////////////////////////////////////////
            // // // Envia o e-mail para o lead na LinhaAtual da BD - ALUNOS.

            // if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

            // await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/sendMail').post({
            //     message: {
            //         subject: 'Ivy - Atualizações Importantes: Encontro Exclusivo, Atendimentos ao Vivo, Materiais Impressos',
            //         body: {
            //             contentType: 'HTML',
            //             content: `
            //                 <p>Boa tarde ${Aluno_PrimeiroNome},</p>
            //                 <p>Este é um e-mail de atualizações importantes referentes a três temas:</p>
            //                 <p>-----</p>
            //                 <p><b>1. Encontro Exclusivo 10 Primeiros Alunos:</b></p>
            //                 <p>Reforçamos que você foi um dos 10 primeiros alunos a entrarem na turma especial de Black Friday e, por isto, está elegível a participar de um encontro exclusivo, com 4h de duração e a portas fechadas, com nosso fundador!<p>
            //                 <p>Neste encontro, serão discutidas lições gerenciais trazidas pelo Lucas de sua experiência recente na Harvard Business School, divididas em três grandes temas:<p>
            //                 <p>&nbsp;&nbsp;&nbsp;&nbsp;a) Interface de conhecimento e conexões lógicas entre Ger. Estratégico e Microeconomia.</p>
            //                 <p>&nbsp;&nbsp;&nbsp;&nbsp;b) Interface de conhecimento e conexões lógicas entre Ger. Tático, Contabilidade e Finanças Corporativas.</p>
            //                 <p>&nbsp;&nbsp;&nbsp;&nbsp;c) Uso de ferramentas avançadas e programação (VS Code, Git, GitHub e Microsoft Azure) no Ger. Inovações e Ger. Rotina.</p>
            //                 <p>O encontro acontecerá no sábado, dia 08/fev/2025, entre 9h e 13h, via Microsoft Teams, por meio <a href="https://teams.microsoft.com/l/meetup-join/19%3ameeting_NjJmMGJjOGMtMDdiMS00NjZkLTlkYzUtNzc3ZjhjYTY2ZGU3%40thread.v2/0?context=%7b%22Tid%22%3a%2249342d16-0605-4267-b540-d1fe7756dbac%22%2c%22Oid%22%3a%22b4a93dcf-5946-4cb2-8368-5db4d242a236%22%7d" target="_blank">deste link</a>.</p>
            //                 <p>Abra o arquivo .ics em anexo e adicione o evento a sua agenda.</p>
            //                 <p>Além disso, reforçamos que os conhecimentos trazidos neste encontro são de <b>extremo valor</b>. Via de regra, estes são temas que abordamos somente junto aos nossos clientes PJ com contratos acima de R$50.000. Por isto, venha preparado. Recomendamos que, na data do encontro, você tenha finalizado pelo menos o estudo dos Módulos 1, 2 e 3 do Preparatório, para que já tenha uma visão sistêmica da Gestão, necessária à absorção adequada dos assuntos e ferramentas trabalhados.</p>
            //                 <p>-----</p>
            //                 <p><b>2. Atendimentos ao Vivo:</b> 
            //                 <p>O primeiro atendimento ao vivo do Preparatório acontecerá na quarta-feira, 18/dez/2024, das 18:30 às 20:00.</p>
            //                 <p>O link de acesso a este atendimento será enviado por e-mail no dia 18/dez, algumas horas antes do encontro.</p>
            //                 <p>Porém, já deixe sua agenda reservada para participar!</p>
            //                 <p>Além das boas-vindas à Turma Especial de Black Friday, o Lucas já iniciará algumas explicações importantes sobre Método Gerencial, com conhecimentos complementares ao conteúdo da plataforma. E estará à disposição para tirar dúvidas iniciais sobre o tema.</p>
            //                 <p>-----</p>
            //                 <p><b>3. Materiais Impressos:</b></p> 
            //                 <p>Seu material impresso está com status <u>Confeccionado</u> e será expedido junto aos correios no dia <u>Entre terça-feira (10/dez/2024) e quinta-feira (12/dez/2024)</u>.</p>                        
            //                 <p><b>Importante!</b> Devido às fortes chuvas no sul do país (onde nossa sede fica localizada), as coletas dos correios estão paralisadas e as expedições terão 3 a 5 dias úteis de atraso. Por isto, a previsão de entrega dos seus materiais foi atualizada para 21/dez/2024.</p>
            //                 <p>Para compensar por quaisquer transtornos, nós prolongamos seu acesso ao serviço por mais 30 dias como uma cortesia :) Seu acesso foi estendido até 05/jan/2026.</p>
            //                 <p>Reiteramos nossa recomendação de que você aguarde a chegada dos materiais para prosseguir com os estudos. As apostilas, guias e cases impressos ajudam muito na absorção do conhecimento!</p>
            //                 <p>-----</p>
            //                 <p>Qualquer dúvida ou insegurança, sempre à disposição.</p>
            //                 <p>Atenciosamente,</p>
            //                 <p><img src="https://plataforma-backend-v3.azurewebsites.net/img/ASSINATURA_E-MAIL.png"/></p>
            //             `
            //         },
            //         toRecipients: [{ emailAddress: { address: Aluno_Email } }],
            //         attachments: [
            //             {
            //                 "@odata.type": "#microsoft.graph.fileAttachment",
            //                 name: "Ivy - Encontro Exclusivo: Black Friday.ics",
            //                 contentBytes: Buffer.from(cal.toString()).toString('base64')
            //             }
            //         ]
            //     }
                
            // })

            // // ////////////////////////////////////////////////////////////////////////////////////////
            // // // Envia o e-mail para o lead na LinhaAtual da BD - ALUNOS.

            // if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

            // await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/sendMail').post({
            //     message: {
            //         subject: 'Ivy - Atualizações Importantes: Encontro Exclusivo, Atendimentos ao Vivo, Materiais Impressos',
            //         body: {
            //             contentType: 'HTML',
            //             content: `
            //                 <p>Boa tarde ${Aluno_PrimeiroNome},</p>
            //                 <p>Este é um e-mail de atualizações importantes referentes a três temas:</p>
            //                 <p>-----</p>
            //                 <p><b>1. Encontro Exclusivo 10 Primeiros Alunos:</b></p>
            //                 <p>Ao longo desta semana, inúmeros alunos solicitaram que nós ampliássemos o número de vagas para o encontro exclusivo com nosso fundador.</p>
            //                 <p>Como várias destas pessoas fizeram a compra do serviço nos primeiros instantes da abertura da turma (ficando de fora da lista de 10 dos primeiros alunos por poucos segundos ou minutos), julgamos que as solicitações tinham mérito. E, por isto, abriremos uma exceção.<p>
            //                 <p><b>Faremos um segundo encontro, com mesmo formato e conteúdo, para todos alunos inscritos na turma.</b></p>
            //                 <p>Neste encontro, serão discutidas lições gerenciais trazidas pelo Lucas de sua experiência recente na Harvard Business School, divididas em três grandes temas:<p>
            //                 <p>&nbsp;&nbsp;&nbsp;&nbsp;a) Interface de conhecimento e conexões lógicas entre Ger. Estratégico e Microeconomia.</p>
            //                 <p>&nbsp;&nbsp;&nbsp;&nbsp;b) Interface de conhecimento e conexões lógicas entre Ger. Tático, Contabilidade e Finanças Corporativas.</p>
            //                 <p>&nbsp;&nbsp;&nbsp;&nbsp;c) Uso de ferramentas avançadas e programação (VS Code, Git, GitHub e Microsoft Azure) no Ger. Inovações e Ger. Rotina.</p>
            //                 <p>O encontro acontecerá no sábado, dia 08/fev/2025, entre 9h e 13h, via Microsoft Teams, por meio <a href="https://teams.microsoft.com/l/meetup-join/19%3ameeting_NjJmMGJjOGMtMDdiMS00NjZkLTlkYzUtNzc3ZjhjYTY2ZGU3%40thread.v2/0?context=%7b%22Tid%22%3a%2249342d16-0605-4267-b540-d1fe7756dbac%22%2c%22Oid%22%3a%22b4a93dcf-5946-4cb2-8368-5db4d242a236%22%7d" target="_blank">deste link</a>.</p>
            //                 <p>Abra o arquivo .ics em anexo e adicione o evento a sua agenda.</p>
            //                 <p>Além disso, reforçamos que os conhecimentos trazidos neste encontro são de <b>extremo valor</b>. Via de regra, estes são temas que abordamos somente junto aos nossos clientes PJ com contratos acima de R$50.000. Por isto, venha preparado. Recomendamos que, na data do encontro, você tenha finalizado pelo menos o estudo dos Módulos 1, 2 e 3 do Preparatório, para que já tenha uma visão sistêmica da Gestão, necessária à absorção adequada dos assuntos e ferramentas trabalhados.</p>
            //                 <p>-----</p>
            //                 <p><b>2. Atendimentos ao Vivo:</b> 
            //                 <p>O primeiro atendimento ao vivo do Preparatório acontecerá na quarta-feira, 18/dez/2024, das 18:30 às 20:00.</p>
            //                 <p>O link de acesso a este atendimento será enviado por e-mail no dia 18/dez, algumas horas antes do encontro.</p>
            //                 <p>Porém, já deixe sua agenda reservada para participar!</p>
            //                 <p>Além das boas-vindas à Turma Especial de Black Friday, o Lucas já iniciará algumas explicações importantes sobre Método Gerencial, com conhecimentos complementares ao conteúdo da plataforma. E estará à disposição para tirar dúvidas iniciais sobre o tema.</p>
            //                 <p>-----</p>
            //                 <p><b>3. Materiais Impressos:</b></p> 
            //                 <p>Seu material impresso está com status <u>Confeccionado</u> e será expedido junto aos correios no dia <u>Entre terça-feira (10/dez/2024) e quinta-feira (12/dez/2024)</u>.</p>                        
            //                 <p><b>Importante!</b> Devido às fortes chuvas no sul do país (onde nossa sede fica localizada), as coletas dos correios estão paralisadas e as expedições terão 3 a 5 dias úteis de atraso. Por isto, a previsão de entrega dos seus materiais foi atualizada para 21/dez/2024.</p>
            //                 <p>Para compensar por quaisquer transtornos, nós prolongamos seu acesso ao serviço por mais 30 dias como uma cortesia :) Seu acesso foi estendido até 05/jan/2026.</p>
            //                 <p>Reiteramos nossa recomendação de que você aguarde a chegada dos materiais para prosseguir com os estudos. As apostilas, guias e cases impressos ajudam muito na absorção do conhecimento!</p>
            //                 <p>-----</p>
            //                 <p>Qualquer dúvida ou insegurança, sempre à disposição.</p>
            //                 <p>Atenciosamente,</p>
            //                 <p><img src="https://plataforma-backend-v3.azurewebsites.net/img/ASSINATURA_E-MAIL.png"/></p>
            //             `
            //         },
            //         toRecipients: [{ emailAddress: { address: 'contato@ivyroom.com.br' } }],
            //         attachments: [
            //             {
            //                 "@odata.type": "#microsoft.graph.fileAttachment",
            //                 name: "Ivy - Encontro Exclusivo: Black Friday.ics",
            //                 contentBytes: Buffer.from(cal.toString()).toString('base64')
            //             }
            //         ]
            //     }
                
            // })

            // ////////////////////////////////////////////////////////////////////////////////////////
            // // Envia o e-mail para o lead na LinhaAtual da BD - ALUNOS.

            if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

            await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/sendMail').post({
                message: {
                    subject: 'Ivy - É hoje! Link de Acesso: Atendimento ao Vivo #1',
                    body: {
                        contentType: 'HTML',
                        content: `
                            <p>Boa tarde ${Aluno_PrimeiroNome},</p>
                            <p>Lembramos que <b>hoje</b>, pontualmente às <b>18h</b>, acontece o <b>primeiro atendimento ao vivo</b> para os participantes da Turma de Black Friday do Preparatório, via Microsoft Teams.</p>
                            <p>Neste encontro, o Lucas já iniciará algumas explicações importantes sobre Método Gerencial, com conhecimentos complementares ao conteúdo da plataforma. E estará à disposição para tirar dúvidas iniciais sobre o tema.</p>
                            <p>Acesse o encontro por meio <a href="https://teams.microsoft.com/l/meetup-join/19%3ameeting_ZGUxZjMwYjktYTU3Ny00Yzc5LTkzYWYtOWU4ZGQ5NDYzM2Y0%40thread.v2/0?context=%7b%22Tid%22%3a%2249342d16-0605-4267-b540-d1fe7756dbac%22%2c%22Oid%22%3a%22b4a93dcf-5946-4cb2-8368-5db4d242a236%22%7d" target="_blank">deste link</a>.</p>
                            <p>Qualquer dúvida ou insegurança, sempre à disposição.</p>
                            <p>Atenciosamente,</p>
                            <p><img src="https://plataforma-backend-v3.azurewebsites.net/img/ASSINATURA_E-MAIL.png"/></p>
                        `
                    },
                    toRecipients: [{ emailAddress: { address: Aluno_Email } }]
                }
                
            })

        }

    }

});