///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////// IMPORTAÇÃO DE BIBLIOTECAS, CRIAÇÃO DE FUNÇÕES E VARIÁVEIS /////////////////////////////////////////////////////
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
// Cria a aplicação necessária para receber HTTP Requests (Express).
// Configura a aplicação para receber as requests de diferentes origens, inclusive do Localhost (Cors).

const express = require('express');
const cors = require('cors');
const app = express();
app.use(cors());
app.use(express.json());

// //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// // Importa a biblioteca para trabalhar com multipart/form-data, necessária para fazer o upload das Fotos de Referência dos alunos.

// const multer = require('multer');

// //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// // Importa as bibliotecas necessárias para comunicação com o Azure Face API.

// const { AzureKeyCredential } = require("@azure/core-auth");
// const FaceClient = require("@azure-rest/ai-vision-face").default;
// const { readFileSync } = require('fs');

//////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Importa o módulo necessário para criar servidores HTTP (HTTP).
// Importa a biblioteca para comunicação bidirecional entre o backend e o frontend (WebSocket).
// Cria o servidor capaz de processar HTTP Requests (Express) e de criar a conexão bidirecional (WebSocket).
// Cria os clientes (frontends) que acessam a conexão bidirecional e estabelece a conexão.

const http = require('http');
const WebSocket = require('ws');
const server = http.createServer(app);
const wss = new WebSocket.Server({ server });
const port = process.env.PORT || 3000;
server.listen(port);

let client = null;

wss.on('connection', (ws) => {
    
    console.log("WebSocket Connected");
    
    client = ws;

    ws.on('close', () => {
        
        client = null;
    
        console.log("WebSocket Disconnected");

    });

});

//////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Importa o módulo necessário para o agendamento de rotinas.

const cron = require('node-cron');

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
// Cria função que transforma datas no formato JavaScript em datas no formato Excel (Horário de Brasília).

function ConverteData2(DataJavaScript) {
    return DataJavaScript.toLocaleString('en-US', { timeZone: 'America/Sao_Paulo', day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false }).replace(',', '');
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Cria função que transforma datas no formato JavaScript em datas no formato DD/MM/AAAA hh:mm.

function ConverteData3(DataJavaScript) {
    return DataJavaScript.toLocaleString('pt-BR', { timeZone: 'America/Sao_Paulo', day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit', hour12: false }).replace(',', '');
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Cria função que transforma datas no formato DD/MM/AAAA hh:mm (Excel) em datas no formato JavaScript.

function ConverteData4(DataExcel) {
    return new Date((DataExcel - 25569) * 86400000); 
}

////////////////////////////////////////////////////////////////////////////////////////
// Importa a biblioteca para criptografar variáveis para o formato SHA256.

const crypto = require('crypto');

////////////////////////////////////////////////////////////////////////////////////////
// Importa a biblioteca para criar Idempotency-Keys.

const { v4: uuidv4 } = require('uuid');

////////////////////////////////////////////////////////////////////////////////////////
// Cria as variáveis de interface com o Meta Graph API.
////////////////////////////////////////////////////////////////////////////////////////

const Meta_Graph_API_Latest_Version = process.env.META_GRAPH_API_LATEST_VERSION;
const Meta_Graph_API_Access_Token = process.env.META_GRAPH_API_ACCESS_TOKEN;
const Meta_Graph_API_Instagram_Business_Account_ID = process.env.META_GRAPH_API_INSTAGRAM_BUSINESS_ACCOUNT_ID;
const Meta_Graph_API_Facebook_Page_ID = process.env.META_GRAPH_API_FACEBOOK_PAGE_ID;
const Meta_Graph_API_Ad_Account_ID = process.env.META_GRAPH_API_AD_ACCOUNT_ID;
const Meta_Graph_API_Custom_Audience_ID_Seguidores = process.env.META_GRAPH_API_CUSTOM_AUDIENCE_ID_SEGUIDORES;

// ////////////////////////////////////////////////////////////////////////////////////////
// // Cria as variáveis de interface com o Azure Face API.
// ////////////////////////////////////////////////////////////////////////////////////////

// const Azure_Face_API_Endpoint = process.env.AZURE_FACE_API_ENDPOINT;
// const Azure_Face_API_Key = process.env.AZURE_FACE_API_KEY;

////////////////////////////////////////////////////////////////////////////////////////
// Cria as variáveis de interface com o API da Pagar.Me.
////////////////////////////////////////////////////////////////////////////////////////

const PagarMe_API_Latest_Version = process.env.PAGARME_API_LATEST_VERSION;
const PagarMe_SecretKey_Base64_Encoded = process.env.PAGARME_SECRETKEY_BASE64_ENCODED;

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////// PROCESSAMENTO DA LANDING PAGE /////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


////////////////////////////////////////////////////////////////////////////////////////
// Serve as imagens estáticas da pasta /img.
////////////////////////////////////////////////////////////////////////////////////////

app.use('/img', express.static('img'));

////////////////////////////////////////////////////////////////////////////////////////
// Endpoint que processa submissão do formulário de cadastro.
////////////////////////////////////////////////////////////////////////////////////////

var Lead_NomeCompleto;
var Lead_PrimeiroNome;
var Lead_Email;

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
        
            .post({"values": [[ConverteData2(new Date()), Lead_Email, Lead_NomeCompleto, "TURMA #2 2025"]]})

            .then(() => {

                res.status(200).send();

            });

    })

    .catch(error => {
        
        res.status(400).send();

    });

});


////////////////////////////////////////////////////////////////////////////////////////
// Endpoint que processa submissão do formulário de Solicitação de Orçamento.
////////////////////////////////////////////////////////////////////////////////////////

app.post('/landingpage/solicitacaoorcamento', async (req, res) => {
    
    var { 
        
        Solicitante_NomeCompleto,
        Solicitante_Email,
        Solicitante_Telefone,
        Solicitante_Cargo,
        Solicitante_NomeEmpresa,
        Solicitante_CNPJ,
        Solicitante_Observações
    
    } = req.body;
    
    ////////////////////////////////////////////////////////////////////////////////////////
    // Envia o e-mail de confirmação de cadastro ao lead, com o evento iCalendar anexado.

    if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

    await Microsoft_Graph_API_Client.api('/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/sendMail').post({
        message: {
            subject: 'Machado - Nova Solicitação de Orçamento',
            body: {
                contentType: 'HTML',
                content: `
                    <p><b>Dados do Solicitante:</b></p>
                    <p>${Solicitante_NomeCompleto}</p>
                    <p>${Solicitante_Email}</p>
                    <p>${Solicitante_Telefone}</p>
                    <p>${Solicitante_Cargo}</p>
                    <p><b>Dados da Empresa:</b></p>
                    <p>${Solicitante_NomeEmpresa}</p>
                    <p>${Solicitante_CNPJ}</p>
                    <p>${Solicitante_Observações}</p>
                    <p><img width="500" height="auto" src="https://plataforma-backend-v3.azurewebsites.net/img/ASSINATURA_E-MAIL.jpg"/></p>
                `
            },
            toRecipients: [{ emailAddress: { address: 'contato@machadogestao.com' } }]
        }
    })

    .then(() => {

        res.status(200).send();

    });

});

////////////////////////////////////////////////////////////////////////////////////////
// Endpoint que processa submissão do formulário de Validação de Certificados.
////////////////////////////////////////////////////////////////////////////////////////

app.post('/landingpage/validacaocertificados', async (req, res) => {
    
    var { Solicitante_CertificadoID, Solicitante_Email } = req.body;
    
    ////////////////////////////////////////////////////////////////////////////////////////
    // Verifica se o Certificado ID está na BD - PLATAFORMA.

    if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

    const BD_Plataforma = await Microsoft_Graph_API_Client.api('/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECRBV4WKTQCI2ZAKCY56VL6IF7ZM/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows').get();
    
    const BD_Plataforma_Número_Linhas = BD_Plataforma.value.length;

    var CertificadoAutenticado = 0;
    var IndexVerificado = 0;
    var CertificadoIDVerificado;
    var NomeCompletoVerificado;

    Verifica_CertificadoID();

    function Verifica_CertificadoID() {

        if (IndexVerificado < BD_Plataforma_Número_Linhas) {
            
            CertificadoIDVerificado = BD_Plataforma.value[IndexVerificado].values[0][16];

            if(Solicitante_CertificadoID === CertificadoIDVerificado) {

                    CertificadoAutenticado = 1;
                    NomeCompletoVerificado = BD_Plataforma.value[IndexVerificado].values[0][0];

            } else {

                IndexVerificado++;
                Verifica_CertificadoID();

            }

        }
    
    }

    if (CertificadoAutenticado === 0){

        await Microsoft_Graph_API_Client.api('/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/sendMail').post({
            message: {
                subject: 'Machado: Resultado - Validação de Certificado',
                body: {
                    contentType: 'HTML',
                    content: `
                        <p>O Certificado ID# pesquisado não foi encontrado em nossas bases de dados internas.</p>
                        <p>Atenciosamente,</p>
                        <p><img width="500" height="auto" src="https://plataforma-backend-v3.azurewebsites.net/img/ASSINATURA_E-MAIL.jpg"/></p>
                    `
                },
                toRecipients: [{ emailAddress: { address: Solicitante_Email } }]
            }
        })

        .then(() => {

            res.status(200).send();

        });

    } else {

        await Microsoft_Graph_API_Client.api('/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/sendMail').post({
            message: {
                subject: 'Machado: Resultado - Validação de Certificado',
                body: {
                    contentType: 'HTML',
                    content: `
                        <p>O Certificado ID# pesquisado é valido!</p>
                        <p>Profissional certificado: ${NomeCompletoVerificado}</p>
                        <p>Atenciosamente,</p>
                        <p><img width="500" height="auto" src="https://plataforma-backend-v3.azurewebsites.net/img/ASSINATURA_E-MAIL.jpg"/></p>
                    `
                },
                toRecipients: [{ emailAddress: { address: Solicitante_Email } }]
            }
        })

        .then(() => {

            res.status(200).send();

        });

    }
    
});

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////// PROCESSAMENTO DO CHECKOUT /////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

////////////////////////////////////////////////////////////////////////////////////////
// Endpoint: Processar Pagamentos.
////////////////////////////////////////////////////////////////////////////////////////

app.post('/checkout/processarpagamento', async (req, res) => {

    ////////////////////////////////////////////////////////////////////////////////////////
    // Obtém as variáveis enviadas pelo frontend.

    let { 
        
        Nome_Produto,
        Código_do_Produto,

        NomeCompleto,
        Email_do_Cliente,
        Campo_de_Preenchimento_CPF,
        Campo_de_Preenchimento_CPF_Dígitos,
        Campo_de_Preenchimento_DDD,
        Campo_de_Preenchimento_Celular,
        Campo_de_Preenchimento_Celular_Dígitos,

        Endereço_Rua,
        Endereço_Número,
        Endereço_Complemento,
        Endereço_Bairro,
        Endereço_Cidade,
        Endereço_Estado,
        Endereço_CEP,
        Endereço_CEP_Dígitos,

        Tipo_de_Pagamento_Escolhido,
        
        Número_do_Cartão,
        Número_do_Cartão_Dígitos,
        Nome_do_Titular_do_Cartão_CaracteresOriginais,
        Nome_do_Titular_do_Cartão_CaracteresAjustados,
        Campo_de_Preenchimento_Mês_Cartão,
        Campo_de_Preenchimento_Ano_Cartão,
        Campo_de_Preenchimento_CVV_Cartão,
        Número_de_Parcelas_Cartão_do_UM_CARTAO,
        Valor_Total_da_Compra_com_Juros_UM_CARTAO,
        Valor_Total_da_Compra_com_Juros_UM_CARTAO_Dígitos,

        Valor_Total_da_Compra_no_PIX_À_VISTA,
        Valor_Total_da_Compra_no_PIX_À_VISTA_Dígitos,

        Valor_Total_da_Compra_no_PIX_CARTÃO,
        Valor_no_PIX_do_PIX_CARTÃO,
        Valor_no_PIX_do_PIX_CARTÃO_Dígitos,
        Valor_com_Juros_no_Cartão_do_PIX_CARTÃO,
        Valor_com_Juros_no_Cartão_do_PIX_CARTÃO_Dígitos,
        Número_do_Cartão_do_PIX_CARTÃO,
        Número_do_Cartão_do_PIX_CARTÃO_Dígitos,
        Nome_do_Titular_do_Cartão_do_PIX_CARTÃO_CaracteresOriginais,
        Nome_do_Titular_do_Cartão_do_PIX_CARTÃO_CaracteresAjustados,
        Campo_de_Preenchimento_Mês_Cartão_do_PIX_CARTÃO,
        Campo_de_Preenchimento_Ano_Cartão_do_PIX_CARTÃO,
        Campo_de_Preenchimento_CVV_Cartão_do_PIX_CARTÃO,
        Número_de_Parcelas_Cartão_do_PIX_CARTÃO
    
    } = req.body;

    let PrimeiroNome = NomeCompleto.split(" ")[0];

    ////////////////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////////////////
    // MODALIDADE DE PAGAMENTO: UM_CARTAO
    ////////////////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////////////////

    ////////////////////////////////////////////////////////////////////////////////////////
    // Insere o pedido na BD - PEDIDOS.

    if (Tipo_de_Pagamento_Escolhido === "UM_CARTAO") {

        if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

        await Microsoft_Graph_API_Client.api('/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECR5HB32SW46MRH37KVHCWW4MDNU/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows')
        
        .post({"values": [[ ConverteData2(new Date()), NomeCompleto, Email_do_Cliente, Campo_de_Preenchimento_DDD, Campo_de_Preenchimento_Celular, Campo_de_Preenchimento_CPF, Endereço_Rua, Endereço_Número, Endereço_Complemento, Endereço_Bairro, Endereço_Cidade, Endereço_Estado, Endereço_CEP, Nome_Produto, Tipo_de_Pagamento_Escolhido, Valor_Total_da_Compra_com_Juros_UM_CARTAO, Valor_Total_da_Compra_com_Juros_UM_CARTAO,  Número_do_Cartão, Nome_do_Titular_do_Cartão_CaracteresOriginais, Campo_de_Preenchimento_Mês_Cartão, Campo_de_Preenchimento_Ano_Cartão, Campo_de_Preenchimento_CVV_Cartão, Número_de_Parcelas_Cartão_do_UM_CARTAO, "-", "-", "-", "-", "-", "-", "-", "-", "-", "-", "-" ]]})  

        .then(async (response) => {

            res.status(200).json({});

            let Número_Linha_Adicionada_à_BD_Cobranças = response.index;

            ////////////////////////////////////////////////////////////////////////////////////////
            // Envia o e-mail de "Novo Pedido Gerado no Checkout".

            if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

            await Microsoft_Graph_API_Client.api('/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/sendMail').post({
                message: {
                    subject: 'Novo Pedido Gerado no Checkout',
                    body: {
                        contentType: 'HTML',
                        content: `
                            <p><b>Aluno:</b> ${NomeCompleto}</p>
                            <p>Atenciosamente,</p>
                            <p><img width="600" height="auto" src="https://plataforma-backend-v3.azurewebsites.net/img/ASSINATURA_E-MAIL.jpg"/></p>
                        `
                    },
                    toRecipients: [{ emailAddress: { address: 'contato@machadogestao.com' } }]
                }
            });

            ////////////////////////////////////////////////////////////////////////////////////////
            // Processa cobrança junto à Pagar.Me.

            fetch(`https://api.pagar.me/core/${PagarMe_API_Latest_Version}/orders`, {
                method: 'POST',
                headers: { 
                    'Authorization': 'Basic ' + PagarMe_SecretKey_Base64_Encoded,
                    'Accept': 'application/json',
                    'Content-Type': 'application/json' 
                },
                body: JSON.stringify({
                    items: [{
                        amount: Valor_Total_da_Compra_com_Juros_UM_CARTAO_Dígitos, 
                        description: Nome_Produto, 
                        quantity: 1,
                        code: Código_do_Produto
                    }],
                    customer: {
                        name: NomeCompleto,
                        type: 'individual', 
                        email: Email_do_Cliente,
                        document: Campo_de_Preenchimento_CPF_Dígitos,
                        document_type: 'CPF',
                        phones: {
                            mobile_phone: {
                                country_code: 55,
                                area_code: Campo_de_Preenchimento_DDD,
                                number: Campo_de_Preenchimento_Celular_Dígitos
                            }
                        }
                    },
                    shipping: {
                        amount: 0,
                        description: Nome_Produto,
                        recipient_name: NomeCompleto,
                        address: {
                            line_1: Endereço_Número + ', ' + Endereço_Rua + ', ' + Endereço_Bairro,
                            line_2: Endereço_Complemento,
                            zip_code: Endereço_CEP_Dígitos,
                            city: Endereço_Cidade,
                            state: Endereço_Estado,
                            country: 'BR'
                        }
                    },
                    payments: [{
                        payment_method: 'credit_card',
                        credit_card: {
                            recurrence: false,
                            installments: Número_de_Parcelas_Cartão_do_UM_CARTAO,
                            statement_descriptor: Código_do_Produto,
                            card: {
                                number: Número_do_Cartão_Dígitos,
                                holder_name: Nome_do_Titular_do_Cartão_CaracteresAjustados,
                                exp_month: Campo_de_Preenchimento_Mês_Cartão,
                                exp_year: Campo_de_Preenchimento_Ano_Cartão,
                                cvv: Campo_de_Preenchimento_CVV_Cartão
                            }
                        }
                    }]
                })
            })

            .then(response => response.json()).then(async json => {

                let Retorno_Processamento_Cobrança_PagarMe = JSON.stringify(json);

                let Status_Cobrança_Cartão = json.charges?.[0]?.status ?? '-';

                ////////////////////////////////////////////////////////////////////////////////////////
                // Insere o Retorno_Processamento_Cobrança_PagarMe e o Status_Cobrança_Cartão na BD - PEDIDOS.

                if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

                await Microsoft_Graph_API_Client.api('/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECR5HB32SW46MRH37KVHCWW4MDNU/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows/itemAt(index=' + Número_Linha_Adicionada_à_BD_Cobranças + ')').update({values: [[null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, Retorno_Processamento_Cobrança_PagarMe, Status_Cobrança_Cartão, null, null, null, null, null, null, null, null, null ]]});

            });

        });

    }


    ////////////////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////////////////
    // MODALIDADE DE PAGAMENTO: PIX_À_VISTA
    ////////////////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////////////////

    ////////////////////////////////////////////////////////////////////////////////////////
    // Insere o pedido na BD - PEDIDOS.

    if (Tipo_de_Pagamento_Escolhido === "PIX_À_VISTA") {

        if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

        await Microsoft_Graph_API_Client.api('/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECR5HB32SW46MRH37KVHCWW4MDNU/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows')
        
        .post({"values": [[ ConverteData2(new Date()), NomeCompleto, Email_do_Cliente, Campo_de_Preenchimento_DDD, Campo_de_Preenchimento_Celular, Campo_de_Preenchimento_CPF, Endereço_Rua, Endereço_Número, Endereço_Complemento, Endereço_Bairro, Endereço_Cidade, Endereço_Estado, Endereço_CEP, Nome_Produto, Tipo_de_Pagamento_Escolhido, Valor_Total_da_Compra_no_PIX_À_VISTA, "-", "-", "-", "-", "-", "-", "-", "-", "-", Valor_Total_da_Compra_no_PIX_À_VISTA, "-", "-", "-", "-", "-", "-", "-", "-" ]]})  

        .then(async (response) => {

            res.status(200).json({});

            let Número_Linha_Adicionada_à_BD_Cobranças = response.index;

            ////////////////////////////////////////////////////////////////////////////////////////
            // Envia o e-mail de "Novo Pedido Gerado no Checkout".

            if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

            await Microsoft_Graph_API_Client.api('/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/sendMail').post({
                message: {
                    subject: 'Novo Pedido Gerado no Checkout',
                    body: {
                        contentType: 'HTML',
                        content: `
                            <p><b>Aluno:</b> ${NomeCompleto}</p>
                            <p>Atenciosamente,</p>
                            <p><img width="600" height="auto" src="https://plataforma-backend-v3.azurewebsites.net/img/ASSINATURA_E-MAIL.jpg"/></p>
                        `
                    },
                    toRecipients: [{ emailAddress: { address: 'contato@machadogestao.com' } }]
                }
            });

            ////////////////////////////////////////////////////////////////////////////////////////
            // Processa cobrança junto à Pagar.Me.

            fetch(`https://api.pagar.me/core/${PagarMe_API_Latest_Version}/orders`, {
                method: 'POST',
                headers: { 
                    'Authorization': 'Basic ' + PagarMe_SecretKey_Base64_Encoded,
                    'Accept': 'application/json',
                    'Content-Type': 'application/json' 
                },
                body: JSON.stringify({
                    items: [{
                        amount: Valor_Total_da_Compra_no_PIX_À_VISTA_Dígitos, 
                        description: Nome_Produto, 
                        quantity: 1,
                        code: Código_do_Produto
                    }],
                    customer: {
                        name: NomeCompleto,
                        type: 'individual', 
                        email: Email_do_Cliente,
                        document: Campo_de_Preenchimento_CPF_Dígitos,
                        document_type: 'CPF',
                        phones: {
                            mobile_phone: {
                                country_code: 55,
                                area_code: Campo_de_Preenchimento_DDD,
                                number: Campo_de_Preenchimento_Celular_Dígitos
                            }
                        }
                    },
                    shipping: {
                        amount: 0,
                        description: Nome_Produto,
                        recipient_name: NomeCompleto,
                        address: {
                            line_1: Endereço_Número + ', ' + Endereço_Rua + ', ' + Endereço_Bairro,
                            line_2: Endereço_Complemento,
                            zip_code: Endereço_CEP_Dígitos,
                            city: Endereço_Cidade,
                            state: Endereço_Estado,
                            country: 'BR'
                        }
                    },
                    payments: [{
                        payment_method: 'pix',
                        pix: {
                            expires_in: 1800
                        }
                    }]
                })
            })

            .then(response => response.json()).then(async json => {

                let Retorno_Processamento_Cobrança_PagarMe = JSON.stringify(json);

                let Status_Cobrança_Pix = json.charges?.[0]?.status ?? '-';

                let Pix_Url_QR_Code = json.charges?.[0]?.last_transaction?.qr_code_url ?? '-';

                let Pix_QR_Code_Prazo_Vencimento = ConverteData3(new Date(Date.now() + 1800000));

                ////////////////////////////////////////////////////////////////////////////////////////
                // Insere o Retorno_Processamento_Cobrança_PagarMe e o Status_Cobrança_Cartão na BD - PEDIDOS.

                if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

                await Microsoft_Graph_API_Client.api('/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECR5HB32SW46MRH37KVHCWW4MDNU/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows/itemAt(index=' + Número_Linha_Adicionada_à_BD_Cobranças + ')').update({values: [[null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, Retorno_Processamento_Cobrança_PagarMe, Status_Cobrança_Pix, null, null, null, null, null, null ]]})

                .then(async (response) => {

                    ////////////////////////////////////////////////////////////////////////////////////////
                    // Envia as instruções de pagamento via PIX ao comprador.

                    if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

                    await Microsoft_Graph_API_Client.api('/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/sendMail').post({
                        message: {
                            subject: 'Ivy: PIX para Pagamento',
                            body: {
                                contentType: 'HTML',
                                content: `
                                    <html>
                                    <head>
                                    <style type="text/css">
                                
                                    #Container_PIX {
                                        height: 340px;
                                        width: 250px;
                                        padding-top: 30px;
                                        padding-bottom: 30px;
                                        padding-left: 30px;
                                        padding-right: 30px;
                                        border: 2px solid rgb(149, 201, 63);
                                        border-radius: 15px;
                                    }
                                        
                                    #Container_Valor_Total_da_Compra_no_PIX {
                                        display: inline-flex;
                                        margin-bottom: 5px;
                                        font-size: 16px;
                                    }
                                    
                                    #Título_Valor_Total_da_Compra_no_PIX {
                                        margin-right: 5px;
                                    }
                                    
                                    #Container_Validade_qr_code_url {
                                        display: inline-flex;
                                        font-size: 16px;
                                    }
                                    
                                    #Título_Validade_qr_code_url {
                                        margin-right: 5px;
                                    }

                                    #qr_code_url{
                                        margin-top: 10px;
                                        height: 250px;
                                        width: 250px; 
                                    }
                                    
                                    #Container_Logo_PagarMe {
                                        display: inline-flex;
                                        margin-left: 38px;
                                        margin-right: auto;
                                    }
                                    
                                    #PoweredBy {
                                        font-size: 10px;
                                        width: 60px;
                                        margin-right: 3px;
                                        margin-top: 20px;
                                    }

                                    #Logo_PagarMe{
                                        width: 110px;
                                        height: 40px;
                                    }
                                        
                                    </style>
                                    </head>
                                    <body>
                                        <p>Prezado(a) ${PrimeiroNome},<br></p>
                                        <p>Para concluir a compra do <b> ${Nome_Produto}</b>, escaneie o QR Code e realize o PIX.<br></p>
                                        <p>As orientações de acesso ao serviço serão enviadas para você por e-mail, assim que o pagamento for processado.<br></p>

                                        <div id="Container_PIX">
                                            <div id="Container_Valor_Total_da_Compra_no_PIX">
                                                    <div id="Título_Valor_Total_da_Compra_no_PIX"><b>Valor:</b></div>
                                                    <div id="Valor_Total_da_Compra_no_PIX">${Valor_Total_da_Compra_no_PIX_À_VISTA}</div>
                                            </div>
                                            <div id="Container_Validade_qr_code_url">
                                                <div id="Título_Validade_qr_code_url"><b>Validade:</b></div>
                                                <div id="Validade_qr_code_url">${Pix_QR_Code_Prazo_Vencimento}</div>
                                            </div>
                                            <img id="qr_code_url" src="${Pix_Url_QR_Code}" alt="qr_code_url">
                                            <div id="Container_Logo_PagarMe">
                                                    <p id="PoweredBy">Powered by</p>
                                                    <img id="Logo_PagarMe" src="https://plataforma-backend-v3.azurewebsites.net/img/LOGO_PAGAR.ME.png"/>
                                            </div>
                                        </div>
                                        
                                    <p><br>Por favor entre em contato se surgirem dúvidas ou se precisar de auxílio.<br></p>

                                    <p>Estamos sempre à disposição.<br></p>

                                    <p>Atenciosamente,<br></p>

                                    <img width="600" height="auto" src="https://plataforma-backend-v3.azurewebsites.net/img/ASSINATURA_E-MAIL.jpg"/>

                                    </body>
                                    </html>
                                `
                            },
                            toRecipients: [{ emailAddress: { address: Email_do_Cliente } }]
                        }
                    });

                });

            });

        });

    }

    ////////////////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////////////////
    // MODALIDADE DE PAGAMENTO: PIX_CARTAO
    ////////////////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////////////////

    ////////////////////////////////////////////////////////////////////////////////////////
    // Insere o pedido na BD - PEDIDOS.

    if (Tipo_de_Pagamento_Escolhido === "PIX_CARTAO") {

        if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

        await Microsoft_Graph_API_Client.api('/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECR5HB32SW46MRH37KVHCWW4MDNU/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows')
        
        .post({"values": [[ ConverteData2(new Date()), NomeCompleto, Email_do_Cliente, Campo_de_Preenchimento_DDD, Campo_de_Preenchimento_Celular, Campo_de_Preenchimento_CPF, Endereço_Rua, Endereço_Número, Endereço_Complemento, Endereço_Bairro, Endereço_Cidade, Endereço_Estado, Endereço_CEP, Nome_Produto, Tipo_de_Pagamento_Escolhido, Valor_Total_da_Compra_no_PIX_CARTÃO, Valor_com_Juros_no_Cartão_do_PIX_CARTÃO, Número_do_Cartão_do_PIX_CARTÃO, Nome_do_Titular_do_Cartão_do_PIX_CARTÃO_CaracteresOriginais, Campo_de_Preenchimento_Mês_Cartão_do_PIX_CARTÃO, Campo_de_Preenchimento_Ano_Cartão_do_PIX_CARTÃO, Campo_de_Preenchimento_CVV_Cartão_do_PIX_CARTÃO, Número_de_Parcelas_Cartão_do_PIX_CARTÃO, "-", "-", Valor_no_PIX_do_PIX_CARTÃO, "-", "-", "-", "-", "-", "-", "-", "-" ]]})  

        .then(async (response) => {

            res.status(200).json({});

            let Número_Linha_Adicionada_à_BD_Cobranças = response.index;

            ////////////////////////////////////////////////////////////////////////////////////////
            // Envia o e-mail de "Novo Pedido Gerado no Checkout".

            if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

            await Microsoft_Graph_API_Client.api('/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/sendMail').post({
                message: {
                    subject: 'Novo Pedido Gerado no Checkout',
                    body: {
                        contentType: 'HTML',
                        content: `
                            <p><b>Aluno:</b> ${NomeCompleto}</p>
                            <p>Atenciosamente,</p>
                            <p><img width="600" height="auto" src="https://plataforma-backend-v3.azurewebsites.net/img/ASSINATURA_E-MAIL.jpg"/></p>
                        `
                    },
                    toRecipients: [{ emailAddress: { address: 'contato@machadogestao.com' } }]
                }
            });

            ////////////////////////////////////////////////////////////////////////////////////////
            // PROCESSAMENTO DA COBRANÇA NO CARTÃO DE CRÉDITO.
            ////////////////////////////////////////////////////////////////////////////////////////

            ////////////////////////////////////////////////////////////////////////////////////////
            // Processa cobrança no Cartão de Crédito junto à Pagar.Me.

            fetch(`https://api.pagar.me/core/${PagarMe_API_Latest_Version}/orders`, {
                method: 'POST',
                headers: { 
                    'Authorization': 'Basic ' + PagarMe_SecretKey_Base64_Encoded,
                    'Accept': 'application/json',
                    'Content-Type': 'application/json' 
                },
                body: JSON.stringify({
                    items: [{
                        amount: Valor_com_Juros_no_Cartão_do_PIX_CARTÃO_Dígitos, 
                        description: Nome_Produto, 
                        quantity: 1,
                        code: Código_do_Produto
                    }],
                    customer: {
                        name: NomeCompleto,
                        type: 'individual', 
                        email: Email_do_Cliente,
                        document: Campo_de_Preenchimento_CPF_Dígitos,
                        document_type: 'CPF',
                        phones: {
                            mobile_phone: {
                                country_code: 55,
                                area_code: Campo_de_Preenchimento_DDD,
                                number: Campo_de_Preenchimento_Celular_Dígitos
                            }
                        }
                    },
                    shipping: {
                        amount: 0,
                        description: Nome_Produto,
                        recipient_name: NomeCompleto,
                        address: {
                            line_1: Endereço_Número + ', ' + Endereço_Rua + ', ' + Endereço_Bairro,
                            line_2: Endereço_Complemento,
                            zip_code: Endereço_CEP_Dígitos,
                            city: Endereço_Cidade,
                            state: Endereço_Estado,
                            country: 'BR'
                        }
                    },
                    payments: [{
                        payment_method: 'credit_card',
                        credit_card: {
                            recurrence: false,
                            installments: Número_de_Parcelas_Cartão_do_PIX_CARTÃO,
                            statement_descriptor: Código_do_Produto,
                            card: {
                                number: Número_do_Cartão_do_PIX_CARTÃO_Dígitos,
                                holder_name: Nome_do_Titular_do_Cartão_do_PIX_CARTÃO_CaracteresAjustados,
                                exp_month: Campo_de_Preenchimento_Mês_Cartão_do_PIX_CARTÃO,
                                exp_year: Campo_de_Preenchimento_Ano_Cartão_do_PIX_CARTÃO,
                                cvv: Campo_de_Preenchimento_CVV_Cartão_do_PIX_CARTÃO
                            }
                        }
                    }]
                })
            })

            .then(response => response.json()).then(async json => {

                let Retorno_Processamento_Cobrança_Cartão_PIX_CARTAO_PagarMe = JSON.stringify(json);

                let Status_Cobrança_Cartão_PIX_CARTAO = json.charges?.[0]?.status ?? '-';

                ////////////////////////////////////////////////////////////////////////////////////////
                // Insere o Retorno_Processamento_Cobrança_PagarMe e o Status_Cobrança_Cartão_PIX_CARTAO na BD - PEDIDOS.

                if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

                await Microsoft_Graph_API_Client.api('/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECR5HB32SW46MRH37KVHCWW4MDNU/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows/itemAt(index=' + Número_Linha_Adicionada_à_BD_Cobranças + ')').update({values: [[null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, Retorno_Processamento_Cobrança_Cartão_PIX_CARTAO_PagarMe, Status_Cobrança_Cartão_PIX_CARTAO, null, null, null, null, null, null, null, null, null ]]})

                .then(async (response) => {

                    ////////////////////////////////////////////////////////////////////////////////////////
                    // PROCESSAMENTO DA COBRANÇA NO PIX.
                    ////////////////////////////////////////////////////////////////////////////////////////

                    ////////////////////////////////////////////////////////////////////////////////////////
                    // Processa cobrança do PIX junto à Pagar.Me.

                    fetch(`https://api.pagar.me/core/${PagarMe_API_Latest_Version}/orders`, {
                        method: 'POST',
                        headers: { 
                            'Authorization': 'Basic ' + PagarMe_SecretKey_Base64_Encoded,
                            'Accept': 'application/json',
                            'Content-Type': 'application/json' 
                        },
                        body: JSON.stringify({
                            items: [{
                                amount: Valor_no_PIX_do_PIX_CARTÃO_Dígitos, 
                                description: Nome_Produto, 
                                quantity: 1,
                                code: Código_do_Produto
                            }],
                            customer: {
                                name: NomeCompleto,
                                type: 'individual', 
                                email: Email_do_Cliente,
                                document: Campo_de_Preenchimento_CPF_Dígitos,
                                document_type: 'CPF',
                                phones: {
                                    mobile_phone: {
                                        country_code: 55,
                                        area_code: Campo_de_Preenchimento_DDD,
                                        number: Campo_de_Preenchimento_Celular_Dígitos
                                    }
                                }
                            },
                            shipping: {
                                amount: 0,
                                description: Nome_Produto,
                                recipient_name: NomeCompleto,
                                address: {
                                    line_1: Endereço_Número + ', ' + Endereço_Rua + ', ' + Endereço_Bairro,
                                    line_2: Endereço_Complemento,
                                    zip_code: Endereço_CEP_Dígitos,
                                    city: Endereço_Cidade,
                                    state: Endereço_Estado,
                                    country: 'BR'
                                }
                            },
                            payments: [{
                                payment_method: 'pix',
                                pix: {
                                    expires_in: 1800
                                }
                            }]
                        })
                    })

                    .then(response => response.json()).then(async json => {

                        let Retorno_Processamento_Cobrança_Pix_PIX_CARTAO_PagarMe = JSON.stringify(json);

                        let Status_Cobrança_Pix = json.charges?.[0]?.status ?? '-';

                        let Pix_Url_QR_Code = json.charges?.[0]?.last_transaction?.qr_code_url ?? '-';

                        let Pix_QR_Code_Prazo_Vencimento = ConverteData3(new Date(Date.now() + 1800000));

                        ////////////////////////////////////////////////////////////////////////////////////////
                        // Insere o Retorno_Processamento_Cobrança_PagarMe e o Status_Cobrança_Cartão na BD - PEDIDOS.

                        if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

                        await Microsoft_Graph_API_Client.api('/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECR5HB32SW46MRH37KVHCWW4MDNU/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows/itemAt(index=' + Número_Linha_Adicionada_à_BD_Cobranças + ')').update({values: [[null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, Retorno_Processamento_Cobrança_Pix_PIX_CARTAO_PagarMe, Status_Cobrança_Pix, null, null, null, null, null, null ]]})

                        .then(async (response) => {

                            ////////////////////////////////////////////////////////////////////////////////////////
                            // Envia as instruções de pagamento via PIX ao comprador.

                            if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

                            await Microsoft_Graph_API_Client.api('/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/sendMail').post({
                                message: {
                                    subject: 'Ivy: PIX para Pagamento',
                                    body: {
                                        contentType: 'HTML',
                                        content: `
                                            <html>
                                            <head>
                                            <style type="text/css">
                                        
                                            #Container_PIX {
                                                height: 340px;
                                                width: 250px;
                                                padding-top: 30px;
                                                padding-bottom: 30px;
                                                padding-left: 30px;
                                                padding-right: 30px;
                                                border: 2px solid rgb(149, 201, 63);
                                                border-radius: 15px;
                                            }
                                                
                                            #Container_Valor_Total_da_Compra_no_PIX {
                                                display: inline-flex;
                                                margin-bottom: 5px;
                                                font-size: 16px;
                                            }
                                            
                                            #Título_Valor_Total_da_Compra_no_PIX {
                                                margin-right: 5px;
                                            }
                                            
                                            #Container_Validade_qr_code_url {
                                                display: inline-flex;
                                                font-size: 16px;
                                            }
                                            
                                            #Título_Validade_qr_code_url {
                                                margin-right: 5px;
                                            }

                                            #qr_code_url{
                                                margin-top: 10px;
                                                height: 250px;
                                                width: 250px; 
                                            }
                                            
                                            #Container_Logo_PagarMe {
                                                display: inline-flex;
                                                margin-left: 38px;
                                                margin-right: auto;
                                            }
                                            
                                            #PoweredBy {
                                                font-size: 10px;
                                                width: 60px;
                                                margin-right: 3px;
                                                margin-top: 20px;
                                            }

                                            #Logo_PagarMe{
                                                width: 110px;
                                                height: 40px;
                                            }
                                                
                                            </style>
                                            </head>
                                            <body>
                                                <p>Prezado(a) ${PrimeiroNome},<br></p>
                                                <p>Para concluir a compra do <b> ${Nome_Produto}</b>, escaneie o QR Code e realize o PIX.<br></p>
                                                <p>As orientações de acesso ao serviço serão enviadas para você por e-mail, assim que o pagamento for processado.<br></p>

                                                <div id="Container_PIX">
                                                    <div id="Container_Valor_Total_da_Compra_no_PIX">
                                                            <div id="Título_Valor_Total_da_Compra_no_PIX"><b>Valor:</b></div>
                                                            <div id="Valor_Total_da_Compra_no_PIX">${Valor_no_PIX_do_PIX_CARTÃO}</div>
                                                    </div>
                                                    <div id="Container_Validade_qr_code_url">
                                                        <div id="Título_Validade_qr_code_url"><b>Validade:</b></div>
                                                        <div id="Validade_qr_code_url">${Pix_QR_Code_Prazo_Vencimento}</div>
                                                    </div>
                                                    <img id="qr_code_url" src="${Pix_Url_QR_Code}" alt="qr_code_url">
                                                    <div id="Container_Logo_PagarMe">
                                                            <p id="PoweredBy">Powered by</p>
                                                            <img id="Logo_PagarMe" src="https://plataforma-backend-v3.azurewebsites.net/img/LOGO_PAGAR.ME.png"/>
                                                    </div>
                                                </div>
                                                
                                            <p><br>Por favor entre em contato se surgirem dúvidas ou se precisar de auxílio.<br></p>

                                            <p>Estamos sempre à disposição.<br></p>

                                            <p>Atenciosamente,<br></p>

                                            <img width="600" height="auto" src="https://plataforma-backend-v3.azurewebsites.net/img/ASSINATURA_E-MAIL.jpg"/>

                                            </body>
                                            </html>
                                        `
                                    },
                                    toRecipients: [{ emailAddress: { address: Email_do_Cliente } }]
                                }
                            });

                        });

                    });

                });

            });

        });

    }

});

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////// PROCESSAMENTO DA PLATAFORMA_v1 ///////////////////////////////////////////////////////////////////
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
var Usuário_Senha;
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
////////////////////////////////////////////////////////////////////////////////////////
// Endpoints que realizam os processos envolvidos no login.
////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////

////////////////////////////////////////////////////////////////////////////////////////
// Endpoint que processa submissão do formulário de login (sem Face ID).
////////////////////////////////////////////////////////////////////////////////////////

app.post('/login', async (req, res) => {
    
    var { login, senha } = req.body;
    Usuário_Login = login;
    Usuário_Senha = senha;

    ////////////////////////////////////////////////////////////////////////////////////////
    // Obtém os dados da BD - PLATAFORMA.xlsx no OneDrive do contato@machadogestao.com.

    if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();
    const BD_Plataforma = await Microsoft_Graph_API_Client.api('/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECRBV4WKTQCI2ZAKCY56VL6IF7ZM/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows').get();
    
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
// Endpoint que processa carregamento da aba /estudos no Frontend (sem FaceID)
////////////////////////////////////////////////////////////////////////////////////////

app.post('/refresh', async (req, res) => {
    
    var { IndexVerificado } = req.body;

    ////////////////////////////////////////////////////////////////////////////////////////
    // Obtém os dados da BD - PLATAFORMA.xlsx no OneDrive do contato@machadogestao.com.

    if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();
    const BD_Plataforma = await Microsoft_Graph_API_Client.api('/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECRBV4WKTQCI2ZAKCY56VL6IF7ZM/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows').get();

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
// Endpoint que atualiza a BD - PLATAFORMA (sem FaceID).
////////////////////////////////////////////////////////////////////////////////////////

app.post('/updates', async (req,res) => {
    
    var { TipoAtualização, IndexVerificado, NúmeroTópicosConcluídos, NúmeroMódulo, NotaTeste, Preparatório2_Interesse } = req.body;

    //Atualiza o Número de Tópicos Concluídos e a Nota no Teste.

    if(TipoAtualização === 'NúmeroTópicosConcluídos-e-NotaTeste'){

        if (NúmeroMódulo === 1){

            if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();
            await Microsoft_Graph_API_Client.api('/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECRBV4WKTQCI2ZAKCY56VL6IF7ZM/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows/itemAt(index=' + IndexVerificado + ')').update({values: [[null, null, null, null, null, null, NúmeroTópicosConcluídos, null, NotaTeste, null, null, null, null, null, null, null, null, null, null]]});

        } else if (NúmeroMódulo === 2){

            if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();
            await Microsoft_Graph_API_Client.api('/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECRBV4WKTQCI2ZAKCY56VL6IF7ZM/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows/itemAt(index=' + IndexVerificado + ')').update({values: [[null, null, null, null, null, null, NúmeroTópicosConcluídos, null, null, NotaTeste, null, null, null, null, null, null, null, null, null]]});

        } else if (NúmeroMódulo === 3){

            if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();
            await Microsoft_Graph_API_Client.api('/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECRBV4WKTQCI2ZAKCY56VL6IF7ZM/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows/itemAt(index=' + IndexVerificado + ')').update({values: [[null, null, null, null, null, null, NúmeroTópicosConcluídos, null, null, null, NotaTeste, null, null, null, null, null, null, null, null]]});

        } else if (NúmeroMódulo === 4){

            if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();
            await Microsoft_Graph_API_Client.api('/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECRBV4WKTQCI2ZAKCY56VL6IF7ZM/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows/itemAt(index=' + IndexVerificado + ')').update({values: [[null, null, null, null, null, null, NúmeroTópicosConcluídos, null, null, null, null, NotaTeste, null, null, null, null, null, null, null]]});

        } else if (NúmeroMódulo === 5){

            if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();
            await Microsoft_Graph_API_Client.api('/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECRBV4WKTQCI2ZAKCY56VL6IF7ZM/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows/itemAt(index=' + IndexVerificado + ')').update({values: [[null, null, null, null, null, null, NúmeroTópicosConcluídos, null, null, null, null, null, NotaTeste, null, null, null, null, null, null]]});

        } else if (NúmeroMódulo === 6){

            if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();
            await Microsoft_Graph_API_Client.api('/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECRBV4WKTQCI2ZAKCY56VL6IF7ZM/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows/itemAt(index=' + IndexVerificado + ')').update({values: [[null, null, null, null, null, null, NúmeroTópicosConcluídos, null, null, null, null, null, null, NotaTeste, null, null, null, null, null]]});

        } else {

            if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();
            await Microsoft_Graph_API_Client.api('/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECRBV4WKTQCI2ZAKCY56VL6IF7ZM/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows/itemAt(index=' + IndexVerificado + ')').update({values: [[null, null, null, null, null, null, NúmeroTópicosConcluídos, null, null, null, null, null, null, null, NotaTeste, null, null, null, null]]});

        } 

    } else if(TipoAtualização === 'NúmeroTópicosConcluídos') {

        //Atualiza só o Número de Tópicos Concluídos do Prep.
        if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();
        await Microsoft_Graph_API_Client.api('/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECRBV4WKTQCI2ZAKCY56VL6IF7ZM/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows/itemAt(index=' + IndexVerificado + ')').update({values: [[null, null, null, null, null, null, NúmeroTópicosConcluídos, null, null, null, null, null, null, null, null, null, null, null, null]]});

    } else if(TipoAtualização === 'Preparatório2_Interesse'){

        //Atualiza só o Interesse no Preparatório 2.
        if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();
        await Microsoft_Graph_API_Client.api('/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECRBV4WKTQCI2ZAKCY56VL6IF7ZM/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows/itemAt(index=' + IndexVerificado + ')').update({values: [[null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, Preparatório2_Interesse]]});

    }

    res.status(200).send();

});

// ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// //////////////////////////////////////////////////////////// PROCESSAMENTO DA PLATAFORMA_v2 ///////////////////////////////////////////////////////////////////
// ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// ////////////////////////////////////////////////////////////////////////////////////////
// ////////////////////////////////////////////////////////////////////////////////////////
// // Declara as variáveis mestras.
// ////////////////////////////////////////////////////////////////////////////////////////
// ////////////////////////////////////////////////////////////////////////////////////////

// var Usuário_Foto_Cadastrada;
// var Usuário_Status_FaceID;

// ////////////////////////////////////////////////////////////////////////////////////////
// ////////////////////////////////////////////////////////////////////////////////////////
// // Endpoints que realizam os processos envolvidos no login.
// ////////////////////////////////////////////////////////////////////////////////////////
// ////////////////////////////////////////////////////////////////////////////////////////

// ////////////////////////////////////////////////////////////////////////////////////////
// // Endpoint que processa submissão do formulário de login (com Face ID).
// ////////////////////////////////////////////////////////////////////////////////////////

// app.post('/login', async (req, res) => {
    
//     var { login, senha } = req.body;
    
//     Usuário_Login = login;
//     Usuário_Senha = senha;

//     ////////////////////////////////////////////////////////////////////////////////////////
//     // Obtém os dados da BD - PLATAFORMA.xlsx.

//     if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();
//     const BD_Plataforma = await Microsoft_Graph_API_Client.api('/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECSBYCZNYGEWFFDLEOZ36WI2PDWO/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows').get();
    
//     ////////////////////////////////////////////////////////////////////////////////////////
//     // Verifica se o Login e Senha cadastrados pelo usuário estão na BD_Plataforma.

//     const BD_Plataforma_Número_Linhas = BD_Plataforma.value.length;

//     var LoginAutenticado = 0;
//     var IndexVerificado = 0;
//     var LoginVerificado;
//     var SenhaVerificada;

//     Verifica_Login_e_Senha();

//     function Verifica_Login_e_Senha() {

//         if (IndexVerificado < BD_Plataforma_Número_Linhas) {
            
//             LoginVerificado = BD_Plataforma.value[IndexVerificado].values[0][2];
//             SenhaVerificada = BD_Plataforma.value[IndexVerificado].values[0][3].toString();

//             if(Usuário_Login === LoginVerificado) {

//                 if(Usuário_Senha === SenhaVerificada) {
                    
//                     LoginAutenticado = 1;

//                     Usuário_PrimeiroNome = BD_Plataforma.value[IndexVerificado].values[0][1];
//                     Usuário_Status_FaceID = BD_Plataforma.value[IndexVerificado].values[0][4];
//                     Usuário_Foto_Cadastrada = BD_Plataforma.value[IndexVerificado].values[0][5];
//                     Usuário_Preparatório1_Status = BD_Plataforma.value[IndexVerificado].values[0][6];
//                     Usuário_Preparatório1_PrazoAcesso = ConverteData(BD_Plataforma.value[IndexVerificado].values[0][7]);
                    
//                 }
                
//             } else {

//                 IndexVerificado++;
//                 Verifica_Login_e_Senha();

//             }

//         }
        
//     }

//     res.status(200).json({ 
        
//         LoginAutenticado,
//         IndexVerificado,
//         Usuário_PrimeiroNome,
//         Usuário_Status_FaceID,
//         Usuário_Foto_Cadastrada,
//         Usuário_Preparatório1_Status,
//         Usuário_Preparatório1_PrazoAcesso

//     });

// });

// ////////////////////////////////////////////////////////////////////////////////////////
// // Endpoints que processam o FaceID (Liveness Session) junto ao Azure Face API.
// ////////////////////////////////////////////////////////////////////////////////////////

// app.post('/CadastroFoto_e_FaceID', multer().single('file'), async (req, res) => {
    
//     let IndexVerificado = req.body.IndexVerificado;
//     let FotoReferência = req.file.buffer;
    
//     // Armazena a FotoReferência em PG - FOTOS DE REFERÊNCIA.
    
//     if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();
    
//     async function UploadFotoReferência(limite_tentativas = 5, intervalo = 2000) {
    
//         for (let tentativa = 1; tentativa <= limite_tentativas; tentativa++) {
//           try {
//             console.log(tentativa);
//             await Microsoft_Graph_API_Client.api(`/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/root:/SISTEMA DE GESTÃO/2. ENTREGA/1. CONTROLAR PLATAFORMA/PG - FOTOS DE REFERÊNCIA/${IndexVerificado}.jpg:/content`).put(FotoReferência);
//             return; 
//           } catch (err) {
//             if (tentativa === limite_tentativas) throw err;
//             await new Promise(res => setTimeout(res, intervalo));
//           }
//         }
    
//     }
    
//     UploadFotoReferência();

//     // Atualiza a coluna FOTO CADASTRADA para 'Sim' na BD - PLATAFORMA.

//     await Microsoft_Graph_API_Client.api('/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECSBYCZNYGEWFFDLEOZ36WI2PDWO/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows/itemAt(index=' + IndexVerificado + ')').update({values: [[null, null, null, null, null, 'Sim', null, null, null, null, null, null, null, null, null, null, null, null, null, null, null]]});
    
//     // Roda o Face ID (Liveness Session).
    
//     let Azure_Face_API_Credential = new AzureKeyCredential(Azure_Face_API_Key);
//     let Azure_Face_API_Client = FaceClient(Azure_Face_API_Endpoint, Azure_Face_API_Credential);

//     let Azure_Face_API_LivenessSession = await Azure_Face_API_Client.path("/detectLivenessWithVerify-sessions")
//     .post({
//         contentType: "multipart/form-data",
//         body: [
//             {
//                 name: "VerifyImage",
//                 body: FotoReferência
//             },
//             {
//                 name: "livenessOperationMode",
//                 body: "Passive"
//             },
//             {
//                 name: "deviceCorrelationId",
//                 body: uuidv4()
//             }
//         ]
//     });

//     let Azure_Face_API_LivenessSession_authToken = Azure_Face_API_LivenessSession.body.authToken;
//     let Azure_Face_API_LivenessSession_sessionID = Azure_Face_API_LivenessSession.body.sessionId;

//     res.status(200).json({ Azure_Face_API_LivenessSession_authToken, Azure_Face_API_LivenessSession_sessionID });

// });

// app.post('/FaceID', async (req, res) => {

//     var { IndexVerificado } = req.body;
    
//     let FotoReferência = readFileSync(`C:/Users/lucas/OneDrive - Machado/2. ENTREGA/1. CONTROLAR PLATAFORMA/PG - FOTOS DE REFERÊNCIA/${IndexVerificado}.jpg`);

//     let Azure_Face_API_Credential = new AzureKeyCredential(Azure_Face_API_Key);
//     let Azure_Face_API_Client = FaceClient(Azure_Face_API_Endpoint, Azure_Face_API_Credential);

//     let Azure_Face_API_LivenessSession = await Azure_Face_API_Client.path("/detectLivenessWithVerify-sessions")
//     .post({
//         contentType: "multipart/form-data",
//         body: [
//             {
//                 name: "VerifyImage",
//                 body: FotoReferência
//             },
//             {
//                 name: "livenessOperationMode",
//                 body: "Passive"
//             },
//             {
//                 name: "deviceCorrelationId",
//                 body: uuidv4()
//             }
//         ]
//     });

//     let Azure_Face_API_LivenessSession_authToken = Azure_Face_API_LivenessSession.body.authToken;
//     let Azure_Face_API_LivenessSession_sessionID = Azure_Face_API_LivenessSession.body.sessionId;

//     res.status(200).json({ Azure_Face_API_LivenessSession_authToken, Azure_Face_API_LivenessSession_sessionID });

// });

// app.get('/FaceID_resultado/:Azure_Face_API_LivenessSession_sessionID', async (req, res) => {

//     let Azure_Face_API_LivenessSession_sessionID = req.params.Azure_Face_API_LivenessSession_sessionID;

//     let Azure_Face_API_Credential = new AzureKeyCredential(Azure_Face_API_Key);
//     let Azure_Face_API_Client = FaceClient(Azure_Face_API_Endpoint, Azure_Face_API_Credential);

//     let Azure_Face_API_LivenessSession  = await Azure_Face_API_Client.path('/detectLivenessWithVerify-sessions/{sessionId}', Azure_Face_API_LivenessSession_sessionID).get();
    
//     let Azure_Face_API_LivenessSession_LivenessDecision = Azure_Face_API_LivenessSession.body.results.attempts[0].result.livenessDecision;
//     let Azure_Face_API_LivenessSession_MatchConfidence = Azure_Face_API_LivenessSession.body.results.attempts[0].result.verifyResult.matchConfidence;
//     let Azure_Face_API_LivenessSession_MatchDecision = Azure_Face_API_LivenessSession.body.results.attempts[0].result.verifyResult.isIdentical;

//     console.log(Azure_Face_API_LivenessSession_LivenessDecision);
//     console.log(Azure_Face_API_LivenessSession_MatchConfidence);
//     console.log(Azure_Face_API_LivenessSession_MatchDecision);

//     res.status(200).json({ Azure_Face_API_LivenessSession_LivenessDecision, Azure_Face_API_LivenessSession_MatchDecision });

// });


// ////////////////////////////////////////////////////////////////////////////////////////
// // Endpoint que processa carregamento da aba /estudos no Frontend (com FaceID)
// ////////////////////////////////////////////////////////////////////////////////////////

// app.post('/refresh', async (req, res) => {
    
//     var { IndexVerificado } = req.body;

//     ////////////////////////////////////////////////////////////////////////////////////////
//     // Obtém os dados da BD - PLATAFORMA.xlsx.
//     if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();
//     const BD_Plataforma = await Microsoft_Graph_API_Client.api('/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECSBYCZNYGEWFFDLEOZ36WI2PDWO/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows').get();

//     Usuário_NomeCompleto = BD_Plataforma.value[IndexVerificado].values[0][0];
//     Usuário_PrimeiroNome = BD_Plataforma.value[IndexVerificado].values[0][1];
//     Usuário_Preparatório1_Status = BD_Plataforma.value[IndexVerificado].values[0][6];
//     Usuário_Preparatório1_PrazoAcesso = ConverteData(BD_Plataforma.value[IndexVerificado].values[0][7]);
//     Usuário_Preparatório1_NúmeroTópicosConcluídos = BD_Plataforma.value[IndexVerificado].values[0][8];
//     Usuário_Preparatório1_NotaMódulo1 = BD_Plataforma.value[IndexVerificado].values[0][10];
//     Usuário_Preparatório1_NotaMódulo2 = BD_Plataforma.value[IndexVerificado].values[0][11];
//     Usuário_Preparatório1_NotaMódulo3 = BD_Plataforma.value[IndexVerificado].values[0][12];
//     Usuário_Preparatório1_NotaMódulo4 = BD_Plataforma.value[IndexVerificado].values[0][13];
//     Usuário_Preparatório1_NotaMódulo5 = BD_Plataforma.value[IndexVerificado].values[0][14];
//     Usuário_Preparatório1_NotaMódulo6 = BD_Plataforma.value[IndexVerificado].values[0][15];
//     Usuário_Preparatório1_NotaMódulo7 = BD_Plataforma.value[IndexVerificado].values[0][16];
//     Usuário_Preparatório1_NotaAcumulado = BD_Plataforma.value[IndexVerificado].values[0][17];
//     Usuário_Preparatório1_CertificadoID = BD_Plataforma.value[IndexVerificado].values[0][18];
//     Usuário_Preparatório2_Interesse = BD_Plataforma.value[IndexVerificado].values[0][20];
                    
//     res.status(200).json({ 
        
//         Usuário_NomeCompleto,
//         Usuário_PrimeiroNome,
//         Usuário_Preparatório1_Status,
//         Usuário_Preparatório1_PrazoAcesso,
//         Usuário_Preparatório1_NúmeroTópicosConcluídos,
//         Usuário_Preparatório1_NotaMódulo1,
//         Usuário_Preparatório1_NotaMódulo2,
//         Usuário_Preparatório1_NotaMódulo3,
//         Usuário_Preparatório1_NotaMódulo4,
//         Usuário_Preparatório1_NotaMódulo5,
//         Usuário_Preparatório1_NotaMódulo6,
//         Usuário_Preparatório1_NotaMódulo7,
//         Usuário_Preparatório1_NotaAcumulado,
//         Usuário_Preparatório1_CertificadoID,
//         Usuário_Preparatório2_Interesse

//     });

// });

// ////////////////////////////////////////////////////////////////////////////////////////
// // Endpoint que atualiza a BD - PLATAFORMA (com FaceID)
// ////////////////////////////////////////////////////////////////////////////////////////

// app.post('/updates', async (req,res) => {
    
//     var { TipoAtualização, IndexVerificado, NúmeroTópicosConcluídos, NúmeroMódulo, NotaTeste, Preparatório2_Interesse } = req.body;

//     //Atualiza o Número de Tópicos Concluídos e a Nota no Teste.

//     if(TipoAtualização === 'NúmeroTópicosConcluídos-e-NotaTeste'){

//         if (NúmeroMódulo === 1){

//             if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();
//             await Microsoft_Graph_API_Client.api('/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECSBYCZNYGEWFFDLEOZ36WI2PDWO/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows/itemAt(index=' + IndexVerificado + ')').update({values: [[null, null, null, null, null, null, null, NúmeroTópicosConcluídos, null, NotaTeste, null, null, null, null, null, null, null, null, null, null]]});

//         } else if (NúmeroMódulo === 2){

//             if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();
//             await Microsoft_Graph_API_Client.api('/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECSBYCZNYGEWFFDLEOZ36WI2PDWO/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows/itemAt(index=' + IndexVerificado + ')').update({values: [[null, null, null, null, null, null, null, NúmeroTópicosConcluídos, null, null, NotaTeste, null, null, null, null, null, null, null, null, null]]});

//         } else if (NúmeroMódulo === 3){

//             if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();
//             await Microsoft_Graph_API_Client.api('/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECSBYCZNYGEWFFDLEOZ36WI2PDWO/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows/itemAt(index=' + IndexVerificado + ')').update({values: [[null, null, null, null, null, null, null, NúmeroTópicosConcluídos, null, null, null, NotaTeste, null, null, null, null, null, null, null, null]]});

//         } else if (NúmeroMódulo === 4){

//             if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();
//             await Microsoft_Graph_API_Client.api('/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECSBYCZNYGEWFFDLEOZ36WI2PDWO/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows/itemAt(index=' + IndexVerificado + ')').update({values: [[null, null, null, null, null, null, null, NúmeroTópicosConcluídos, null, null, null, null, NotaTeste, null, null, null, null, null, null, null]]});

//         } else if (NúmeroMódulo === 5){

//             if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();
//             await Microsoft_Graph_API_Client.api('/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECSBYCZNYGEWFFDLEOZ36WI2PDWO/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows/itemAt(index=' + IndexVerificado + ')').update({values: [[null, null, null, null, null, null, null, NúmeroTópicosConcluídos, null, null, null, null, null, NotaTeste, null, null, null, null, null, null]]});

//         } else if (NúmeroMódulo === 6){

//             if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();
//             await Microsoft_Graph_API_Client.api('/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECSBYCZNYGEWFFDLEOZ36WI2PDWO/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows/itemAt(index=' + IndexVerificado + ')').update({values: [[null, null, null, null, null, null, null, NúmeroTópicosConcluídos, null, null, null, null, null, null, NotaTeste, null, null, null, null, null]]});

//         } else {

//             if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();
//             await Microsoft_Graph_API_Client.api('/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECSBYCZNYGEWFFDLEOZ36WI2PDWO/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows/itemAt(index=' + IndexVerificado + ')').update({values: [[null, null, null, null, null, null, null, NúmeroTópicosConcluídos, null, null, null, null, null, null, null, NotaTeste, null, null, null, null]]});

//         } 

//     } else if(TipoAtualização === 'NúmeroTópicosConcluídos') {

//         //Atualiza só o Número de Tópicos Concluídos do Prep.
//         if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();
//         await Microsoft_Graph_API_Client.api('/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECSBYCZNYGEWFFDLEOZ36WI2PDWO/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows/itemAt(index=' + IndexVerificado + ')').update({values: [[null, null, null, null, null, null, null, NúmeroTópicosConcluídos, null, null, null, null, null, null, null, null, null, null, null, null]]});

//     } else if(TipoAtualização === 'Preparatório2_Interesse'){

//         //Atualiza só o Interesse no Preparatório 2.
//         if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();
//         await Microsoft_Graph_API_Client.api('/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECSBYCZNYGEWFFDLEOZ36WI2PDWO/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows/itemAt(index=' + IndexVerificado + ')').update({values: [[null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, Preparatório2_Interesse]]});

//     }

//     res.status(200).send();

// });

////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////
// Endpoint que serve como Authorization URL para o PlayReady do EZDRM.
////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////

// app.get('/ezdrm-playready-authorization-url', (req, res) => {

//   console.log(req);
  
//   const token = req.query.token || "";
//   const customData = req.query.CustomData || "";

//   const response =
//     "p1=5&p2=&p3=&p4=1&p5=0&p6=1&p7=0&p8=0" +
//     "&token=" + encodeURIComponent(token) +
//     "&CustomData=" + encodeURIComponent(customData);

//   res.set("Content-Type", "text/plain");
//   res.status(200).send(response);

// });

app.post('/ezdrm-playready-authorization-url', (req, res) => {
  res.status(200);
  res.set("Content-Type", "text/plain");
  res.send("p1=5&p2=&p3=&p4=1&p5=0&p6=1&p7=0&p8=0");
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
// Envia e-mail de RL em escala para os leads na BD - LEADS.
////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////

app.post('/leads/email_RL', async (req,res) => {
    
    res.status(200).send();
    
    ////////////////////////////////////////////////////////////////////////////////////////
    // Puxa os dados da BD - LEADS.
    
    if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

    const BD_Leads = await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/drive/items/0172BBJBYG24NEFOMGOJCLN5FMDILTSZTC/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{AC8C07F3-9A79-4ABD-8CE8-0C818B0EA1A7}/rows').get();

    const BD_Leads_Última_Linha = BD_Leads.value.length - 1;

    async function Envia_Email_Leads() {
    
        for (let LinhaAtual = 0; LinhaAtual <= BD_Leads_Última_Linha; LinhaAtual++) {
            
            ProcessamentoLeads_Email = BD_Leads.value[LinhaAtual].values[0][1];
            ProcessamentoLeads_PrimeiroNome = BD_Leads.value[LinhaAtual].values[0][2].split(" ")[0];

            ////////////////////////////////////////////////////////////////////////////////////////
            // Envia o e-mail para o lead na LinhaAtual da BD - LEADS.

            if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

            await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/sendMail').post({
                message: {
                    subject: 'Ivy - Conteúdo 🎯: A Cultura do Sugar até Espanar',
                    body: {
                        contentType: 'HTML',
                        content: `
                            <p>Olá ${ProcessamentoLeads_PrimeiroNome},</p>
                            <p>Quem escreve é Lucas Machado, fundador da Ivy | Escola de Gestão. Tudo bem?</p>
                            <p>Passo para compartilhar um conteúdo de extremo valor para profissionais com interesse em Gestão. Espero que traga boas reflexões.</p>
                            <p>--------------------------------------</p>
                            <p><b>A CULTURA DO SUGAR ATÉ ESPANAR</b></p>
                            <p>Não seja ingênuo.</p>
                            <p>No Brasil, infelizmente, tem muita empresa que adota a cultura do <b>sugar até espanar</b>.</p>
                            <p>Ou seja.</p>
                            <p>São empresas que buscam contratar gente que trabalha duro e “pede” pouco, e que sugam estas pessoas ao máximo (sem as contrapartidas coerentes, é claro) até que elas espanem e peçam demissão.</p>
                            <p>Daí a pessoa é substituída. E o ciclo reinicia.</p>
                            <p>Por isto, tenha segurança disto: esta é uma cultura <b>medíocre</b>. Isto é a antítese da boa Gestão. Cedo ou tarde estas empresas quebram. E se você for vítima deste ciclo, minha orientação é: não hesite em sair.</p>
                            <p>--------------------------------------</p>
                            <p>Caso queira se aprofundar no tema, acompanhe os stories e nosso Canal de Transmissão no Instagram amanhã (sexta, 07/mar).
                            <p>P.S. Nas próximas semanas traremos mais conteúdos nesta linha.</p>
                            <p>P.S.2. Idem para informações sobre a próxima turma do Preparatório em Gestão Generalista.</p>
                            <p>Sempre à disposição</p>
                            <p>Atenciosamente,</p>
                            <p><img src="https://plataforma-backend-v3.azurewebsites.net/img/ASSINATURA_E-MAIL.png"/></p>
                        `
                    },
                    toRecipients: [{ emailAddress: { address: ProcessamentoLeads_Email } }]
                }
                
            })

            await new Promise(resolve => setTimeout(resolve, 2000));

            console.log(`E-mail ${LinhaAtual + 1} enviado: ${ProcessamentoLeads_PrimeiroNome}`);

        }
    
    }

    Envia_Email_Leads();

});

////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////
// Envia e-mail de CV em escala para os leads na BD - LEADS.
////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////

app.post('/leads/email_CV', async (req,res) => {
    
    res.status(200).send();

    ////////////////////////////////////////////////////////////////////////////////////////
    // Puxa os dados da BD - LEADS.
    
    if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

    const BD_Leads = await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/drive/items/0172BBJBYG24NEFOMGOJCLN5FMDILTSZTC/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{AC8C07F3-9A79-4ABD-8CE8-0C818B0EA1A7}/rows').get();

    const BD_Leads_Última_Linha = BD_Leads.value.length - 1;

    async function Envia_Email_Leads() {
    
        for (let LinhaAtual = 0; LinhaAtual <= BD_Leads_Última_Linha; LinhaAtual++) {
            
            Lead_Email = BD_Leads.value[LinhaAtual].values[0][1];
            Lead_PrimeiroNome = BD_Leads.value[LinhaAtual].values[0][2].split(" ")[0];

            ////////////////////////////////////////////////////////////////////////////////////////
            // Envia o e-mail para o lead na LinhaAtual da BD - ALUNOS.

            if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

            await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/sendMail').post({

                message: {
                    subject: 'Ivy - 🚨ÚLTIMA CHAMADA🚨',
                    body: {
                        contentType: 'HTML',
                        content: `
                            <p>Última chamada!</p>
                            <p>As inscrições para a próxima turma do Prep. Gestão Generalista encerram em menos de 60min (<b>hoje, quarta, 16/abril às 22:00</b> via <a href="https://ivygestao.com/">Link da Bio</a>).</p>
                            <p>Se você quer construir carreira gerencial, este é um dos momentos mais importantes em toda a sua trajetória.</p>
                            <p>Não deixe a oportunidade passar.</p>
                            <p>Dúvidas em resposta a este e-mail.</p>
                            <p>Atenciosamente,</p>
                            <p><img src="https://plataforma-backend-v3.azurewebsites.net/img/ASSINATURA_E-MAIL.png"/></p>
                        `
                    },
                    toRecipients: [{ emailAddress: { address: Lead_Email } }]
                }
            
            });

            await new Promise(resolve => setTimeout(resolve, 1000));

            console.log(`E-mail ${LinhaAtual + 1} enviado: ${Lead_PrimeiroNome}`);

        }
    
    }

    Envia_Email_Leads();

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
// Envia e-mails individuais (em escala) para os alunos na BD - ALUNOS.
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
    
    for (let LinhaAtual = 0; LinhaAtual <= 2; LinhaAtual++) {
                
        Aluno_Email = BD_Alunos.value[LinhaAtual].values[0][2];
        Aluno_PrimeiroNome = BD_Alunos.value[LinhaAtual].values[0][1].split(" ")[0];

        if (Aluno_Email === "-") {

        } else {

            // ////////////////////////////////////////////////////////////////////////////////////////
            // // Envia o e-mail para o aluno atual na BD - ALUNOS.

            if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

            await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/sendMail').post({
                message: {
                    subject: 'Teste',
                    body: {
                        contentType: 'HTML',
                        content: `
                            <p>Bom dia ${Aluno_PrimeiroNome},</p>
                            <p>Este é um teste.</p>
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

app.post('/alunos/envioemail02', async (req,res) => {

    let { Data_Início_Atendimentos, Link_Microsoft_Teams } = req.body;
    
    //res.status(200).json({ message: "1. Request recebida." });

    console.log(`1. Request recebida.`);

    let [Dia_Início_Atendimentos,Mês_Início_Atendimentos,Ano_Início_Atendimentos] = Data_Início_Atendimentos.split("/").map(num => parseInt(num, 10));

    let Dia_da_Semana_Data_Início_Atendimentos = new Intl.DateTimeFormat('pt-BR', { weekday: 'long' }).format(new Date(Ano_Início_Atendimentos, Mês_Início_Atendimentos - 1, Dia_Início_Atendimentos));

    ////////////////////////////////////////////////////////////////////////////////////////
    // Puxa os dados da BD - ATENDIMENTOS.
    
    if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

    const BD_Atendimentos = await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/drive/items/0172BBJB6LBE4HY3JHYZFJHV2OJWRVOW2W/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows').get();

    //if (BD_Atendimentos !== null && client) client.send(JSON.stringify({ message: `2. BD - ATENDIMENTOS obtida.`, origin: "ConviteAtendimentos" }));
    
    if (BD_Atendimentos !== null) console.log(`2. BD - ATENDIMENTOS obtida.`);

    const BD_Atendimentos_Última_Linha = BD_Atendimentos.value.length - 1;

    let Número_Invite_Enviado = 0;

    ////////////////////////////////////////////////////////////////////////////////////////
    // Aguarda 1s para iniciar o envio dos e-mails, para que o WebSocket possa enviar os dados de volta ao frontend.
    // Então envia um invite a cada 2s.
    
    async function Envia_Invites_Atendimentos() {

        for (let LinhaAtual = 256; LinhaAtual <= BD_Atendimentos_Última_Linha; LinhaAtual++) {

            ///////////////////////////////////////////////////////////////////////////////////////////////////
            // Puxa as variáveis do aluno da BD - ATENDIMENTOS.
    
            Aluno_PrimeiroNome = BD_Atendimentos.value[LinhaAtual].values[0][1].split(" ")[0];
            Aluno_Email = BD_Atendimentos.value[LinhaAtual].values[0][2];
            Aluno_Status_Envio_Convite_Atendimentos = BD_Atendimentos.value[LinhaAtual].values[0][3];
    
            if (Aluno_Status_Envio_Convite_Atendimentos === "SIM") {
    
                Número_Invite_Enviado++;
    
                //if (client) client.send(JSON.stringify({ message: `3. Invite #${Número_Invite_Enviado} enviado para: ${Aluno_PrimeiroNome}`, origin: "ConviteAtendimentos" }));
    
                console.log(`3. Invite #${Número_Invite_Enviado} enviado para: ${Aluno_PrimeiroNome}`);
                
                //if (LinhaAtual === BD_Atendimentos_Última_Linha && client) client.send(JSON.stringify({ message: `--- fim ---`, origin: "ConviteAtendimentos" }));
                
                if (LinhaAtual === BD_Atendimentos_Última_Linha) console.log(`--- fim ---`);

                ///////////////////////////////////////////////////////////////////////////////////////////////////
                // Cria o evento iCalendar para os Atendimentos, com alerta de 1 hora antes do início do encontro.
    
                const cal = new ICalCalendar({ domain: 'ivyroom.com.br', prodId: { company: 'Ivy | Escola de Gestão', product: 'Ivy - Atendimentos', language: 'PT-BR' } });
                const event = cal.createEvent({
                    start: new Date(Date.UTC(Ano_Início_Atendimentos, Mês_Início_Atendimentos - 1, Dia_Início_Atendimentos, 21, 30, 0)), // 18:30 BRT
                    end: new Date(Date.UTC(Ano_Início_Atendimentos, Mês_Início_Atendimentos - 1, Dia_Início_Atendimentos, 23, 0, 0)), // 20:00 BRT
                    summary: 'Atendimento ao Vivo',
                    description: ` Link do Encontro (Microsoft Teams): ${Link_Microsoft_Teams}`,
                    uid: `${new Date().getTime()}@ivyroom.com.br`,
                    stamp: new Date()
                });
    
                event.createAlarm({
                    type: 'display',
                    trigger: 1 * 60 * 60 * 1000,
                    description: 'Atendimento ao Vivo (Ivy) - Inicia em 1 hora.'
                });
    
                ////////////////////////////////////////////////////////////////////////////////////////
                // Envia o e-mail para o aluno na LinhaAtual da BD - ATENDIMENTOS.
    
                if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();
    
                await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/sendMail').post({
    
                    message: {
                        subject: 'Ivy - Atualizações: Bônus e Atendimentos ao Vivo',
                        body: {
                            contentType: 'HTML',
                            content: `
                                <p>Bom dia ${Aluno_PrimeiroNome},</p>
                                <p>Passamos para atualizá-los sobre dois temas importantes:<br><br></p>
                                <p><b>1. BÔNUS</b></p>
                                <p>Todos os bônus já foram preparados e embalados para expedição. Para alunos com endereço cadastrado em:</p>
                                <p><b>• Curitiba/PR e Região Metropolitana:</b> bônus será enviado via Uber Delivery ou Loggi na segunda-feira (28/abril). Fique atento! Entraremos em contato via WhatsApp para alinharmos os detalhes antes do envio.</p>
                                <p><b>• Demais localidades:</b> bônus já expedido via Correios. Previsão de entrega entre hoje (24/abril) e segunda-feira (28/abril). Monitoramento feito automaticamente por nós via API. Entramos em contato se necessário. Basta aguardar.<br><br></p>
                                <p><b>2. ATENDIMENTOS AO VIVO</b></p>
                                <p>O primeiro atendimento ao vivo com o Lucas, nosso fundador, acontecerá <b>${Dia_da_Semana_Data_Início_Atendimentos} (${Data_Início_Atendimentos}) às 18:30</b>, via Microsoft Teams, por meio <a href=${Link_Microsoft_Teams} target="_blank">deste link</a>.</p>
                                <p>Abra o arquivo .ics em anexo e adicione o evento a sua agenda.</p>
                                <p>Reforçamos que você é o protagonista destes encontros. Por isto, se prepare previamente e tenha em mãos suas dúvidas, anotações e materiais de suporte ao Prep.<br><br></p> 
                                <p>Qualquer dúvida, à disposição.</p>
                                <p>Atenciosamente,</p>
                                <p><img src="https://plataforma-backend-v3.azurewebsites.net/img/ASSINATURA_E-MAIL.png"/></p>
                            `
                        },
                        toRecipients: [{ emailAddress: { address: Aluno_Email } }],
                        attachments: [
                            {
                                "@odata.type": "#microsoft.graph.fileAttachment",
                                name: "Ivy - Atendimento ao Vivo.ics",
                                contentBytes: Buffer.from(cal.toString()).toString('base64')
                            }
                        ]
                    }
                
                });

                await new Promise(resolve => setTimeout(resolve, 2000));
    
            } else {

                await new Promise(resolve => setTimeout(resolve, 0));

                //if (LinhaAtual === BD_Atendimentos_Última_Linha && client) client.send(JSON.stringify({ message: `--- fim ---`, origin: "ConviteAtendimentos" }));

                if (LinhaAtual === BD_Atendimentos_Última_Linha) console.log(`--- fim ---`);

            }
    
        }

    }

    setTimeout(Envia_Invites_Atendimentos, 1000);

});

app.post('/alunos/envioemail03', async (req,res) => {

    console.log(`1. Request recebida.`);

    ////////////////////////////////////////////////////////////////////////////////////////
    // Puxa os dados da BD - ATENDIMENTOS.
    
    if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

    //////////////////////////////////////////////////////////////////////////////////////
    // Envia o e-mail para o aluno na LinhaAtual da BD - ATENDIMENTOS.

    const BD_Atendimentos = await Microsoft_Graph_API_Client.api('/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECTOY3PC2EDNIJG2C4B7OWMAJL7J/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows').get();

    if (BD_Atendimentos !== null) console.log(`2. BD - ATENDIMENTOS obtida.`);

    const BD_Atendimentos_Última_Linha = BD_Atendimentos.value.length - 1;

    let Número_Email_Enviado = 0;

    ////////////////////////////////////////////////////////////////////////////////////////
    // Aguarda 1s para iniciar o envio dos e-mails, para que o WebSocket possa enviar os dados de volta ao frontend.
    // Então envia um invite a cada 2s.
    
    async function Envia_Invites_Atendimentos() {

        for (let LinhaAtual = 139; LinhaAtual <= BD_Atendimentos_Última_Linha; LinhaAtual++) {

            ///////////////////////////////////////////////////////////////////////////////////////////////////
            // Puxa as variáveis do aluno da BD - ATENDIMENTOS.
    
            Aluno_PrimeiroNome = BD_Atendimentos.value[LinhaAtual].values[0][1].split(" ")[0];
            Aluno_Email = BD_Atendimentos.value[LinhaAtual].values[0][2];
    
            if (Aluno_Email !== "-") {
    
                Número_Email_Enviado++;
    
                console.log(`3. E-mail #${Número_Email_Enviado} enviado para: ${Aluno_PrimeiroNome}`);
                
                if (LinhaAtual === BD_Atendimentos_Última_Linha) console.log(`--- fim ---`);

                ////////////////////////////////////////////////////////////////////////////////////////
                // Envia o e-mail para o aluno na LinhaAtual da BD - ATENDIMENTOS.
    
                if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();
    
                await Microsoft_Graph_API_Client.api('/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/sendMail').post({
    
                    message: {
                        subject: 'Machado (antiga Ivy) - Novo Link de Acesso à Plataforma',
                        body: {
                            contentType: 'HTML',
                            content: `
                                <p>Boa tarde ${Aluno_PrimeiroNome},</p>
                                <p>Quem escreve é Lucas Machado, fundador da Machado (antiga Ivy). Como vai?</p>
                                <p>Sinalizo que, a partir de hoje (22/ago/2025 às 15h30) o acesso ao Preparatório em Gestão Generalista deve realizado por meio do link:</p>
                                <p><b><a href="https://machadogestao.com/plataforma/login">https://machadogestao.com/plataforma/login</a></b></p>
                                <p>O link anterior (https://ivygestao.com/plataforma/login) está sendo desativado de forma definitiva.</p>
                                <p>Qualquer dúvida ou insegurança, sempre à disposição.</p>
                                <p>Atenciosamente,</p>
                                <p><img src="https://plataforma-backend-v3.azurewebsites.net/img/ASSINATURA_E-MAIL.jpg" width="600" /></p>
                            `
                        },
                        toRecipients: [{ emailAddress: { address: Aluno_Email } }]
                    }
                
                });

                await new Promise(resolve => setTimeout(resolve, 2000));
    
            } else {

                await new Promise(resolve => setTimeout(resolve, 0));

                if (LinhaAtual === BD_Atendimentos_Última_Linha) console.log(`--- fim ---`);

            }
    
        }

    }

    setTimeout(Envia_Invites_Atendimentos, 1000);

});

////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////
// Envia e-mails individuais para clientes na BD - PLATAFORMA.
////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////

app.post('/alunos/envioemail04', async (req,res) => {

    console.log(`1. Request recebida.`);

    ////////////////////////////////////////////////////////////////////////////////////////
    // Puxa os dados da BD - PLATAFORMA.
    
    if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

    //////////////////////////////////////////////////////////////////////////////////////
    // Envia o e-mail para o aluno na LinhaAtual da BD - ATENDIMENTOS.

    const BD_Plataforma = await Microsoft_Graph_API_Client.api('/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECRBV4WKTQCI2ZAKCY56VL6IF7ZM/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows').get();

    if (BD_Plataforma !== null) console.log(`2. BD_Plataforma obtida.`);

    const BD_Plataforma_Última_Linha = BD_Plataforma.value.length - 1;

    let Número_Email_Enviado = 0;

    ////////////////////////////////////////////////////////////////////////////////////////
    // Então envia um email a cada 2s.
    
    async function Envia_Email_Clientes() {

        for (let LinhaAtual = 287; LinhaAtual <= BD_Plataforma_Última_Linha; LinhaAtual++) {

            ///////////////////////////////////////////////////////////////////////////////////////////////////
            // Puxa as variáveis do aluno da BD - PLATAFORMA.
    
            let Cliente_PrimeiroNome = BD_Plataforma.value[LinhaAtual].values[0][1];
            let Cliente_Email = BD_Plataforma.value[LinhaAtual].values[0][2];
            let Cliente_Senha = BD_Plataforma.value[LinhaAtual].values[0][3];
    
            Número_Email_Enviado++;

            console.log(`3. E-mail #${Número_Email_Enviado} enviado para: ${Cliente_PrimeiroNome}`);
            
            if (LinhaAtual === BD_Plataforma_Última_Linha) console.log(`--- fim ---`);

            ////////////////////////////////////////////////////////////////////////////////////////
            // Envia o e-mail para o aluno na LinhaAtual da BD - ATENDIMENTOS.

            if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

            await Microsoft_Graph_API_Client.api('/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/sendMail').post({

                message: {
                    subject: 'Instruções Iniciais e Acesso à Plataforma - Machado | Método Gerencial para Empresas',
                    body: {
                        contentType: 'HTML',
                        content: `
                            <p>Bom dia ${Cliente_PrimeiroNome},</p>
                            <p>Quem escreve é Lucas Machado, fundador da Machado | Método Gerencial para Empresas. Como vai?</p>
                            <p>Recentemente a MGF contratou nosso serviço, a Solução em Método Gerencial, para auxiliarmos no amadurecimento do Sistema de Gestão da empresa. E você é um dos profissionais selecionados para participar!</p>
                            <p>Nosso trabalho conjunto terá duas grandes etapas:</p>
                            <p><b>• Preparatório em Gestão Generalista:</b> etapa de formação em gerenciamento científico (Método Gerencial), que acontece por meio de nossa plataforma de ensino em formato assíncrono (com vídeos, ferramentas para download e estudos de caso reais). Esta etapa dura cerca de 60 dias e o acompanhamento do progresso de vocês (percentual de evolução nos estudos e resultados nos testes ao final de cada um dos 7 módulos) será gerenciado via Status Reports enviados ao grupo de WhatsApp que criaremos ao longo do dia de hoje.</p>
                            <p><b>• Encontros Presenciais:</b> etapa posterior à formação. Aqui iremos tirar dúvidas, revisar conceitos e de fato fazer o ataque consultivo à MGF. Nesta fase, eu irei até vocês para 3 dias inteiros de encontros presenciais (das 9h às 18h). A data destes encontros será definida junto à diretoria, posteriormente.</p>
                            <p>Amanhã (quarta-feira, 17/set às 9h) faremos um <a href="https://teams.microsoft.com/l/meetup-join/19%3ameeting_Y2VkYTQ3NzktZmQ0NS00M2U0LWFhM2ItZGEyMmI2ZGYxYzgx%40thread.v2/0?context=%7b%22Tid%22%3a%2249342d16-0605-4267-b540-d1fe7756dbac%22%2c%22Oid%22%3a%22a8f570ff-a292-4b2f-a1e4-629ccd7a26be%22%7d">encontro via Microsoft Teams</a> para passarmos por todos estes pontos em detalhes. Por favor entre na reunião no horário. Começaremos pontualmente.</p>
                            <p>Dito isto, por favor utilize as credenciais abaixo para acessar o Preparatório:</p>
                            <span><b>Link:</b> <a href="https://machadogestao.com/plataforma/login">https://machadogestao.com/plataforma/login</a><br></span>
                            <span><b>Login:</b> ${Cliente_Email}<br></span>
                            <span><b>Senha:</b> ${Cliente_Senha}<br></span>
                            <p><b>Observações Importantes:</b></p>
                            <p>• Caso tenha qualquer dificuldade de acesso à plataforma, envie em resposta a este e-mail ou via inbox no WhatsApp para +55 41 99679 9092. Iremos auxiliá-lo(a) prontamente.</p>
                            <p>• Já preparamos e estamos expedindo amanhã (17/set, via Sedex) <b>materiais impressos</b> que darão suporte aos estudos. Cada um de vocês receberá, no endereço da empresa, uma caixa personalizada com apostilas, guias de aplicação rápida do conhecimento e um compilado com todos os estudos de caso. <b>Sugerimos aguardar a chegada dos materiais antes de iniciarem os estudos.</b></p>
                            <p>• Estamos passando por um processo de rebranding de nossa empresa (nossa marca original é Ivy | Escola de Gestão). Por isto, parte do Preparatório e dos materiais impressos ainda remetem a esta marca. Tudo certo neste ponto. Não há impacto nos conteúdos.</p>
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

////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////
// Envia convites para os atendimentos ao vivo.
////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////

app.post('/alunos/convite-atendimentos', async (req,res) => {

    let { Data_Início_Atendimentos, Link_Microsoft_Teams } = req.body;
    
    //res.status(200).json({ message: "1. Request recebida." });

    console.log(`1. Request recebida.`);

    let [Dia_Início_Atendimentos,Mês_Início_Atendimentos,Ano_Início_Atendimentos] = Data_Início_Atendimentos.split("/").map(num => parseInt(num, 10));

    let Dia_da_Semana_Data_Início_Atendimentos = new Intl.DateTimeFormat('pt-BR', { weekday: 'long' }).format(new Date(Ano_Início_Atendimentos, Mês_Início_Atendimentos - 1, Dia_Início_Atendimentos));

    ////////////////////////////////////////////////////////////////////////////////////////
    // Puxa os dados da BD - ATENDIMENTOS.
    
    if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

    const BD_Atendimentos = await Microsoft_Graph_API_Client.api('/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECTOY3PC2EDNIJG2C4B7OWMAJL7J/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows').get();

    //if (BD_Atendimentos !== null && client) client.send(JSON.stringify({ message: `2. BD - ATENDIMENTOS obtida.`, origin: "ConviteAtendimentos" }));
    
    if (BD_Atendimentos !== null) console.log(`2. BD - ATENDIMENTOS obtida.`);
    
    const BD_Atendimentos_Última_Linha = BD_Atendimentos.value.length - 1;

    let Número_Invite_Enviado = 0;

    ////////////////////////////////////////////////////////////////////////////////////////
    // Aguarda 1s para iniciar o envio dos e-mails, para que o WebSocket possa enviar os dados de volta ao frontend.
    // Então envia um invite a cada 2s.
    
    async function Envia_Invites_Atendimentos() {

        for (let LinhaAtual = 204; LinhaAtual <= BD_Atendimentos_Última_Linha; LinhaAtual++) {

            ///////////////////////////////////////////////////////////////////////////////////////////////////
            // Puxa as variáveis do aluno da BD - ATENDIMENTOS.
    
            Aluno_PrimeiroNome = BD_Atendimentos.value[LinhaAtual].values[0][1].split(" ")[0];
            Aluno_Email = BD_Atendimentos.value[LinhaAtual].values[0][2];
            Aluno_Status_Envio_Convite_Atendimentos = BD_Atendimentos.value[LinhaAtual].values[0][3];
    
            if (Aluno_Status_Envio_Convite_Atendimentos === "SIM") {
    
                Número_Invite_Enviado++;
    
                //if (client) client.send(JSON.stringify({ message: `3. Invite #${Número_Invite_Enviado} enviado para: ${Aluno_PrimeiroNome}`, origin: "ConviteAtendimentos" }));
    
                console.log(`3. Invite #${Número_Invite_Enviado} enviado para: ${Aluno_PrimeiroNome}`);
                
                //if (LinhaAtual === BD_Atendimentos_Última_Linha && client) client.send(JSON.stringify({ message: `--- fim ---`, origin: "ConviteAtendimentos" }));
                
                if (LinhaAtual === BD_Atendimentos_Última_Linha) console.log(`--- fim ---`);

                ///////////////////////////////////////////////////////////////////////////////////////////////////
                // Cria o evento iCalendar para os Atendimentos, com alerta de 1 hora antes do início do encontro.
    
                const cal = new ICalCalendar({ domain: 'machadogestao.com', prodId: { company: 'Machado | Método Gerencial para Empresas', product: 'Machado - Atendimento ao Vivo', language: 'PT-BR' } });
                const event = cal.createEvent({
                    start: new Date(Date.UTC(Ano_Início_Atendimentos, Mês_Início_Atendimentos - 1, Dia_Início_Atendimentos, 21, 30, 0)), // 18:30 BRT
                    end: new Date(Date.UTC(Ano_Início_Atendimentos, Mês_Início_Atendimentos - 1, Dia_Início_Atendimentos, 23, 0, 0)), // 20:00 BRT
                    summary: 'Atendimento ao Vivo',
                    description: ` Link do Encontro (Microsoft Teams): ${Link_Microsoft_Teams}`,
                    uid: `${new Date().getTime()}@machadogestao.com`,
                    stamp: new Date()
                });
    
                event.createAlarm({
                    type: 'display',
                    trigger: 1 * 60 * 60 * 1000,
                    description: 'Atendimento ao Vivo (Machado) - Inicia em 1 hora.'
                });
    
                ////////////////////////////////////////////////////////////////////////////////////////
                // Envia o e-mail para o aluno na LinhaAtual da BD - ATENDIMENTOS.
    
                if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();
    
                await Microsoft_Graph_API_Client.api('/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/sendMail').post({
    
                    message: {
                        subject: 'Machado - Convite: Atendimento ao Vivo',
                        body: {
                            contentType: 'HTML',
                            content: `
                                <p>Olá ${Aluno_PrimeiroNome},</p>
                                <p>Informamos que o próximo atendimento ao vivo com o Lucas Machado, acontecerá <b>${Dia_da_Semana_Data_Início_Atendimentos} (${Data_Início_Atendimentos}) às 18:30</b>, via Microsoft Teams, por meio <a href=${Link_Microsoft_Teams} target="_blank">deste link</a>.</p>
                                <p><b>Por favor abra o arquivo .ics em anexo e adicione o evento a sua agenda.</b></p>
                                <p>Reforçamos que você é o protagonista destes encontros. Por isto, se prepare previamente e tenha em mãos suas dúvidas, anotações e materiais de suporte ao Prep.</p> 
                                <p>Qualquer dúvida, à disposição.</p>
                                <p>Atenciosamente,</p>
                                <p><img width="600" height="auto" src="https://plataforma-backend-v3.azurewebsites.net/img/ASSINATURA_E-MAIL.jpg"/></p>
                            `
                        },
                        toRecipients: [{ emailAddress: { address: Aluno_Email } }],
                        attachments: [
                            {
                                "@odata.type": "#microsoft.graph.fileAttachment",
                                name: "Machado - Atendimento ao Vivo.ics",
                                contentBytes: Buffer.from(cal.toString()).toString('base64')
                            }
                        ]
                    }
                
                });

                await new Promise(resolve => setTimeout(resolve, 2000));
    
            } else {

                await new Promise(resolve => setTimeout(resolve, 0));

                //if (LinhaAtual === BD_Atendimentos_Última_Linha && client) client.send(JSON.stringify({ message: `--- fim ---`, origin: "ConviteAtendimentos" }));

                if (LinhaAtual === BD_Atendimentos_Última_Linha) console.log(`--- fim ---`);

            }
    
        }

    }

    setTimeout(Envia_Invites_Atendimentos, 1000);

});

////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////
// Envia lembretes para os atendimentos ao vivo.
////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////

app.post('/alunos/lembrete-atendimentos', async (req,res) => {

    let { Data_Início_Atendimentos, Link_Microsoft_Teams } = req.body;
    
    //res.status(200).json({ message: "1. Request recebida." });

    console.log(`1. Request recebida.`);

    let [Dia_Início_Atendimentos,Mês_Início_Atendimentos,Ano_Início_Atendimentos] = Data_Início_Atendimentos.split("/").map(num => parseInt(num, 10));

    let Dia_da_Semana_Data_Início_Atendimentos = new Intl.DateTimeFormat('pt-BR', { weekday: 'long' }).format(new Date(Ano_Início_Atendimentos, Mês_Início_Atendimentos - 1, Dia_Início_Atendimentos));

    ////////////////////////////////////////////////////////////////////////////////////////
    // Puxa os dados da BD - ATENDIMENTOS.
    
    if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

    const BD_Atendimentos = await Microsoft_Graph_API_Client.api('/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECTOY3PC2EDNIJG2C4B7OWMAJL7J/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows').get();

    //if (BD_Atendimentos !== null && client) client.send(JSON.stringify({ message: `2. BD - ATENDIMENTOS obtida.`, origin: "ConviteAtendimentos" }));
    
    if (BD_Atendimentos !== null) console.log(`2. BD - ATENDIMENTOS obtida.`);
    
    const BD_Atendimentos_Última_Linha = BD_Atendimentos.value.length - 1;

    let Número_Lembrete_Enviado = 0;

    ////////////////////////////////////////////////////////////////////////////////////////
    // Aguarda 1s para iniciar o envio dos e-mails, para que o WebSocket possa enviar os dados de volta ao frontend.
    // Então envia um invite a cada 2s.
    
    async function Envia_Lembretes_Atendimentos() {

        for (let LinhaAtual = 204; LinhaAtual <= BD_Atendimentos_Última_Linha; LinhaAtual++) {

            ///////////////////////////////////////////////////////////////////////////////////////////////////
            // Puxa as variáveis do aluno da BD - ATENDIMENTOS.
    
            Aluno_PrimeiroNome = BD_Atendimentos.value[LinhaAtual].values[0][1].split(" ")[0];
            Aluno_Email = BD_Atendimentos.value[LinhaAtual].values[0][2];
            Aluno_Status_Envio_Convite_Atendimentos = BD_Atendimentos.value[LinhaAtual].values[0][3];
    
            if (Aluno_Status_Envio_Convite_Atendimentos === "SIM") {
    
                Número_Lembrete_Enviado++;
    
                //if (client) client.send(JSON.stringify({ message: `3. Invite #${Número_Lembrete_Enviado} enviado para: ${Aluno_PrimeiroNome}`, origin: "ConviteAtendimentos" }));
    
                console.log(`3. Invite #${Número_Lembrete_Enviado} enviado para: ${Aluno_PrimeiroNome}`);
                
                //if (LinhaAtual === BD_Atendimentos_Última_Linha && client) client.send(JSON.stringify({ message: `--- fim ---`, origin: "ConviteAtendimentos" }));
                
                if (LinhaAtual === BD_Atendimentos_Última_Linha) console.log(`--- fim ---`);

                ///////////////////////////////////////////////////////////////////////////////////////////////////
                // Cria o evento iCalendar para os Atendimentos, com alerta de 1 hora antes do início do encontro.
    
                const cal = new ICalCalendar({ domain: 'machadogestao.com', prodId: { company: 'Machado | Método Gerencial para Empresas', product: 'Machado - Atendimento ao Vivo', language: 'PT-BR' } });
                const event = cal.createEvent({
                    start: new Date(Date.UTC(Ano_Início_Atendimentos, Mês_Início_Atendimentos - 1, Dia_Início_Atendimentos, 21, 30, 0)), // 18:30 BRT
                    end: new Date(Date.UTC(Ano_Início_Atendimentos, Mês_Início_Atendimentos - 1, Dia_Início_Atendimentos, 23, 0, 0)), // 20:00 BRT
                    summary: 'Atendimento ao Vivo',
                    description: ` Link do Encontro (Microsoft Teams): ${Link_Microsoft_Teams}`,
                    uid: `${new Date().getTime()}@machadogestao.com`,
                    stamp: new Date()
                });
    
                event.createAlarm({
                    type: 'display',
                    trigger: 1 * 60 * 60 * 1000,
                    description: 'Atendimento ao Vivo (Machado) - Inicia em 1 hora.'
                });
    
                ////////////////////////////////////////////////////////////////////////////////////////
                // Envia o e-mail para o aluno na LinhaAtual da BD - ATENDIMENTOS.
    
                if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();
    
                await Microsoft_Graph_API_Client.api('/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/sendMail').post({
    
                    message: {
                        subject: 'Machado - Lembrete: Atendimento ao Vivo',
                        body: {
                            contentType: 'HTML',
                            content: `
                                <p>Olá ${Aluno_PrimeiroNome},</p>
                                <p>Lembramos que o próximo atendimento ao vivo com o Lucas Machado acontecerá hoje, <b>${Dia_da_Semana_Data_Início_Atendimentos} (${Data_Início_Atendimentos}) às 18:30</b>, via Microsoft Teams, por meio <a href=${Link_Microsoft_Teams} target="_blank">deste link</a>.</p>
                                <p><b>Por favor abra o arquivo .ics em anexo e adicione o evento a sua agenda.</b></p>
                                <p>Reforçamos que você é o protagonista destes encontros. Por isto, se prepare previamente e tenha em mãos suas dúvidas, anotações e materiais de suporte ao Prep.</p> 
                                <p>Qualquer dúvida, à disposição.</p>
                                <p>Atenciosamente,</p>
                                <p><img width="600" height="auto" src="https://plataforma-backend-v3.azurewebsites.net/img/ASSINATURA_E-MAIL.jpg"/></p>
                            `
                        },
                        toRecipients: [{ emailAddress: { address: Aluno_Email } }],
                        attachments: [
                            {
                                "@odata.type": "#microsoft.graph.fileAttachment",
                                name: "Machado - Atendimento ao Vivo.ics",
                                contentBytes: Buffer.from(cal.toString()).toString('base64')
                            }
                        ]
                    }
                
                });

                await new Promise(resolve => setTimeout(resolve, 2000));
    
            } else {

                await new Promise(resolve => setTimeout(resolve, 0));

                //if (LinhaAtual === BD_Atendimentos_Última_Linha && client) client.send(JSON.stringify({ message: `--- fim ---`, origin: "ConviteAtendimentos" }));

                if (LinhaAtual === BD_Atendimentos_Última_Linha) console.log(`--- fim ---`);

            }
    
        }

    }

    setTimeout(Envia_Lembretes_Atendimentos, 1000);

});

// ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// //////////////////////////////////////////////////////////////// META GRAPH API INTERFACE /////////////////////////////////////////////////////////////////////
// ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////
// Endpoint: Registrar novo post na BD - RESULTADOS
////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////

app.post('/meta/registrar-novo-post', async (req, res) => {

    let { Reel_Código } = req.body;

    ////////////////////////////////////////////////////////////////////////////////////////
    // Obtém o IG Media ID.
    ////////////////////////////////////////////////////////////////////////////////////////
    
    fetch(`https://graph.facebook.com/${Meta_Graph_API_Latest_Version}/${Meta_Graph_API_Instagram_Business_Account_ID}/media?fields=id,timestamp&limit=1&access_token=${Meta_Graph_API_Access_Token}`, { method: 'GET'})

    .then(response => response.json()).then(async data => {

        let Reel_IG_Media_ID = data.data[0].id;
        let Reel_Data_Hora_Postagem = data.data[0].timestamp;

        if (Reel_IG_Media_ID !== null) console.log(`1. IG Media ID obtido: ${Reel_IG_Media_ID}`);
        if (Reel_Data_Hora_Postagem !== null) console.log(`2. Data e Hora da postagem obtidos: ${Reel_Data_Hora_Postagem}`);

        ////////////////////////////////////////////////////////////////////////////////////////
        // Obtém o Número de Seguidores atualizado.
        ////////////////////////////////////////////////////////////////////////////////////////

        fetch(`https://graph.facebook.com/${Meta_Graph_API_Latest_Version}/${Meta_Graph_API_Instagram_Business_Account_ID}?fields=followers_count&access_token=${Meta_Graph_API_Access_Token}`, { method: 'GET'})

        .then(response => response.json()).then(async data => {

            let Número_Seguidores = data.followers_count;

            if (Número_Seguidores !== null) console.log(`3. Número de Seguidores obtido: ${Número_Seguidores}`);
        
            ///////////////////////////////////////////////////////////////////////////////////////
            // Adiciona as informações à BD - RESULTADOS (RELACIONAMENTO).
            ///////////////////////////////////////////////////////////////////////////////////////

            if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

            await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/drive/items/0172BBJB5JTOTCSWCLGBB2HKLEFJVR7AUC/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{122865F8-2E2D-4B60-A34C-E02E001E835E}/rows')
            
            .post({"values": [[ Reel_Código, `'${Reel_IG_Media_ID}`, ConverteData2(new Date()), Número_Seguidores, null, null, null, null, null, null, null ]]})
            
            .then(async response => {

                console.log(`4. BD - RESULTADOS atualizada.`);
                
                /////////////////////////////////////////////////////////////////////////////////////////////////////
                // Cria o evento na agenda (calendário) para criação da campanha de DB (72h depois).
                /////////////////////////////////////////////////////////////////////////////////////////////////////

                let Horário_Início_Criação_Campanha_DB = new Date(new Date().setMinutes(0, 0, 0) + 3 * 24 * 60 * 60 * 1000);
                let Horário_Término_Criação_Campanha_DB = new Date(Horário_Início_Criação_Campanha_DB.getTime() + 60 * 60 * 1000);

                if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

                await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/calendar/events').post({
                    
                    subject: "CAMPANHA DB - " + Reel_Código,
                    
                    start: {
                        "dateTime": Horário_Início_Criação_Campanha_DB,
                        "timeZone": "UTC"
                    },
                    
                    end: {
                        "dateTime": Horário_Término_Criação_Campanha_DB,
                        "timeZone": "UTC"
                    }
                    
                })

                .then(async () => {

                    console.log(`5. Criação da campanha de DB agendada.`);
                    
                });

            });

        });

    });

});

////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////
// Endpoint: Registra os Desempenhos Orgânicos de 5% e de 72h.
// ---> Acionado pela function01.js uma vez a cada 10min <--- 
////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////

app.post('/meta/RegistraDesempenhosOrganicos', async (req,res) => {

    res.status(200).send();

    let Data_e_Hora_Atual = new Date(new Date().getTime() - 3 * 60 * 60 * 1000);

    console.log(`Endpoint /meta/RegistraDesempenhosOrganicos acionado agora (${Data_e_Hora_Atual}) pela function01.js.`);

    ////////////////////////////////////////////////////////////////////////////////////////
    // Puxa os dados da BD - RESULTADOS do OneDrive do contato@machadogestao.com.
    
    if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

    let BD_Resultados_RL = await Microsoft_Graph_API_Client.api('/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECT6SJFAPWNDHZAZ4NX5CRUWSUQG/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{122865F8-2E2D-4B60-A34C-E02E001E835E}/rows').get();

    ///////////////////////////////////////////////////////////////////////////////////////////////////////////
    // Verifica se há criativos com:
    // --> O número de CONTAS ALCANÇADAS (5%) e INTERAÇÕES (5%) em branco.
    // --> O número de CONTAS ALCANÇADAS (72h) e INTERAÇÕES (72h) em branco.

    let BD_Resultados_RL_Número_Linhas = BD_Resultados_RL.value.length;
    let BD_Resultados_RL_Última_Linha = BD_Resultados_RL_Número_Linhas - 1;
    
    for (let LinhaVerificada = 0; LinhaVerificada <= BD_Resultados_RL_Última_Linha; LinhaVerificada++) {

        let Reel_IG_Media_ID = BD_Resultados_RL.value[LinhaVerificada].values[0][1];
        let Reel_Data_e_Hora_Postagem = ConverteData4(BD_Resultados_RL.value[LinhaVerificada].values[0][2]);
        let Reel_Número_de_Seguidores_Momento_Postagem = BD_Resultados_RL.value[LinhaVerificada].values[0][3];
        let Reel_Contas_Alcançadas_5Porcento_Registrado = BD_Resultados_RL.value[LinhaVerificada].values[0][4];
        let Reel_Contas_Alcançadas_72Horas_Registrado = BD_Resultados_RL.value[LinhaVerificada].values[0][7];

        ///////////////////////////////////////////////////////////////////////////////////////////////////////////
        // Caso número de CONTAS ALCANÇADAS (5%) e INTERAÇÕES (5%) esteja em branco.
        // E tenham passado menos de 72h desde a postagem.
        // --> Puxa os dados do Meta Graph API.
        
        if (Reel_Contas_Alcançadas_5Porcento_Registrado === "" && Data_e_Hora_Atual - Reel_Data_e_Hora_Postagem < 72 * 60 * 60 * 1000){

            fetch(`https://graph.facebook.com/${Meta_Graph_API_Latest_Version}/${Reel_IG_Media_ID}/insights?metric=reach,likes,saved,shares&access_token=${Meta_Graph_API_Access_Token}`, { method: 'GET'})

            .then(response => response.json()).then(async data => {

                let Reel_Organic_Reach_Atual = data.data.find(metric => metric.name === 'reach').values[0].value;

                ///////////////////////////////////////////////////////////////////////////////////////////////////////////
                // Caso o Reel_Organic_Reach_5_Porcento seja >= 5% do Criativo_Número_de_Seguidores_Momento_Postagem:
                // --> Registra as informações do criativo na BD - RESULTADOS.

                if (Reel_Organic_Reach_Atual >= Math.ceil(Reel_Número_de_Seguidores_Momento_Postagem * 0.05)) {

                    let Reel_Organic_Likes_Atual = data.data.find(metric => metric.name === 'likes').values[0].value;
                    let Reel_Organic_Saved_Atual = data.data.find(metric => metric.name === 'saved').values[0].value;
                    let Reel_Organic_Shares_Atual = data.data.find(metric => metric.name === 'shares').values[0].value;

                    let Reel_Organic_Interactions_Atual = Reel_Organic_Likes_Atual + Reel_Organic_Saved_Atual + Reel_Organic_Shares_Atual;

                    if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

                    await Microsoft_Graph_API_Client.api('/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECT6SJFAPWNDHZAZ4NX5CRUWSUQG/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{122865F8-2E2D-4B60-A34C-E02E001E835E}/rows/itemAt(index=' + LinhaVerificada + ')').update({values: [[null, null, null, null, Reel_Organic_Reach_Atual, Reel_Organic_Interactions_Atual, null, null, null, null, null ]]})

                } 

            });

        }

        ///////////////////////////////////////////////////////////////////////////////////////////////////////////
        // Caso número de CONTAS ALCANÇADAS (72h) e INTERAÇÕES (72h) em branco.
        // E tenham passado mais ou igual a 72h desde a postagem.
        // --> Puxa os dados do Meta Graph API.

        if (Reel_Contas_Alcançadas_72Horas_Registrado === "" && Data_e_Hora_Atual - Reel_Data_e_Hora_Postagem >= 72 * 60 * 60 * 1000){

            fetch(`https://graph.facebook.com/${Meta_Graph_API_Latest_Version}/${Reel_IG_Media_ID}/insights?metric=reach,likes,saved,shares&access_token=${Meta_Graph_API_Access_Token}`, { method: 'GET'})

            .then(response => response.json()).then(async data => {

                let Reel_Organic_Reach_Atual = data.data.find(metric => metric.name === 'reach').values[0].value;
                let Reel_Organic_Likes_Atual = data.data.find(metric => metric.name === 'likes').values[0].value;
                let Reel_Organic_Saved_Atual = data.data.find(metric => metric.name === 'saved').values[0].value;
                let Reel_Organic_Shares_Atual = data.data.find(metric => metric.name === 'shares').values[0].value;

                let Reel_Organic_Interactions_Atual = Reel_Organic_Likes_Atual + Reel_Organic_Saved_Atual + Reel_Organic_Shares_Atual;

                ///////////////////////////////////////////////////////////////////////////////////////////////////////////
                // Caso número de CONTAS ALCANÇADAS (5%) e INTERAÇÕES (5%) também esteja em branco.
                // --> Registra as informações do criativo na BD - RESULTADOS:
                //       - Tanto nas colunas CONTAS ALCANÇADAS (5%) e INTERAÇÕES (5%).
                //       - Quanto nas colunas CONTAS ALCANÇADAS (72h) e INTERAÇÕES (72h).

                if ( Reel_Contas_Alcançadas_5Porcento_Registrado === "" ) {

                    if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

                    await Microsoft_Graph_API_Client.api('/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECT6SJFAPWNDHZAZ4NX5CRUWSUQG/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{122865F8-2E2D-4B60-A34C-E02E001E835E}/rows/itemAt(index=' + LinhaVerificada + ')').update({values: [[null, null, null, null, Reel_Organic_Reach_Atual, Reel_Organic_Interactions_Atual, null, Reel_Organic_Reach_Atual, null, Reel_Organic_Interactions_Atual, null ]]});
                    
                } 
                
                ///////////////////////////////////////////////////////////////////////////////////////////////////////////
                // Caso número de CONTAS ALCANÇADAS (5%) e INTERAÇÕES (5%) não esteja em branco.
                // --> Registra as informações do criativo na BD - RESULTADOS:
                //       - Somente nas colunas CONTAS ALCANÇADAS (72h) e INTERAÇÕES (72h).
                
                else {

                    if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();
                    
                    await Microsoft_Graph_API_Client.api('/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECT6SJFAPWNDHZAZ4NX5CRUWSUQG/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{122865F8-2E2D-4B60-A34C-E02E001E835E}/rows/itemAt(index=' + LinhaVerificada + ')').update({values: [[null, null, null, null, null, null, null, Reel_Organic_Reach_Atual, null, Reel_Organic_Interactions_Atual, null ]]});
                    
                }

            });

        }

    };

});

////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////
// Endpoint: Registra Desempenho - Campanhas DB.
// ---> Acionado pela function02.js às 00:01 AM, todos os dias <--- 
////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////

app.post('/meta/RegistraDesempenhosCampanhasDB', async (req,res) => {

    let Data_Hoje_Formatada_Meta_Graph_API = (new Date()).toISOString().split('T')[0];

    ///////////////////////////////////////////////////////////////////////////////////////
    // Obtém a BD - STATUS CAMPANHAS do OneDrive do contato@machadogestao.com.
    ///////////////////////////////////////////////////////////////////////////////////////

    if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();
    
    let BD_Status_Campanhas_DB = await Microsoft_Graph_API_Client.api('/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECUCH2N5WBZ3MJHJRSW3UAV6PDRX/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{93C2A633-D78C-42B0-9A68-937848657884}/rows').get();

    const BD_Status_Campanhas_DB_Última_Linha = BD_Status_Campanhas_DB.value.length - 1;

    ///////////////////////////////////////////////////////////////////////////////////////
    // Registra um de desempenho de Campanha de DB a cada 2s.
    ///////////////////////////////////////////////////////////////////////////////////////

    for (let LinhaAtual = 0; LinhaAtual <= BD_Status_Campanhas_DB_Última_Linha; LinhaAtual++) {

        let Campanha_DB_Status = BD_Status_Campanhas_DB.value[LinhaAtual].values[0][4];

        if (Campanha_DB_Status === "ATIVA") {

            let Campanha_DB_Reel_Código = BD_Status_Campanhas_DB.value[LinhaAtual].values[0][0];
            let Campanha_DB_Ad_ID = BD_Status_Campanhas_DB.value[LinhaAtual].values[0][1];
            let Campanha_DB_Descrição = BD_Status_Campanhas_DB.value[LinhaAtual].values[0][2];
            let Campanha_DB_Qualidade_Clique = BD_Status_Campanhas_DB.value[LinhaAtual].values[0][3];
            
            ///////////////////////////////////////////////////////////////////////////////////////////////////
            // Obtém as variáveis de desempenho do Ad junto ao Meta Graph API.

            fetch(`https://graph.facebook.com/${Meta_Graph_API_Latest_Version}/${Campanha_DB_Ad_ID}/insights?fields=campaign_id,spend,reach,impressions,actions&time_range={"since":"2022-08-31","until":"${Data_Hoje_Formatada_Meta_Graph_API}"}&filtering=[{field: "action_type",operator:"IN", value: ['link_click']}]`, {
                method: 'GET',
                headers: {
                    'Content-Type': 'application/json',
                    'Authorization': `Bearer ${Meta_Graph_API_Access_Token}`
                }
            })

            .then(response => response.json()).then(async data => {

                let Campanha_DB_Campaign_ID = data.data[0].campaign_id;
                let Campanha_DB_Ad_Spend = data.data[0].spend;
                let Campanha_DB_Ad_Reach = data.data[0].reach;
                let Campanha_DB_Ad_Impressions = data.data[0].impressions;
                let Campanha_DB_Ad_Link_Clicks = data.data[0].actions[0].value;

                ///////////////////////////////////////////////////////////////////////////////////////////////////
                // Obtém o orçamento da campanha junto ao Meta Graph API.

                fetch(`https://graph.facebook.com/${Meta_Graph_API_Latest_Version}/${Campanha_DB_Campaign_ID}?fields=daily_budget`, {
                    method: 'GET',
                    headers: {
                        'Content-Type': 'application/json',
                        'Authorization': `Bearer ${Meta_Graph_API_Access_Token}`
                    }
                })

                .then(response => response.json()).then(async data => {

                    let Campanha_DB_Campaign_Daily_Budget = data.daily_budget;
                    
                    ////////////////////////////////////////////////////////////////////////////////////////
                    // Adiciona as informações à BD - RESULTADOS CAMPANHAS do OneDrive do contato@machadogestao.com.

                    if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

                    await Microsoft_Graph_API_Client.api('/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECVHHGFYL55S4NBKGCBC43AZB3SY/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{93C2A633-D78C-42B0-9A68-937848657884}/rows')
                    
                    .post({"values": [[ "-", ConverteData2(new Date(new Date().setDate(new Date().getDate() - 1))), Campanha_DB_Reel_Código, `'${Campanha_DB_Ad_ID}`, Campanha_DB_Descrição, Campanha_DB_Qualidade_Clique, Campanha_DB_Ad_Spend, Campanha_DB_Ad_Reach, Campanha_DB_Ad_Impressions, Campanha_DB_Ad_Link_Clicks, "-", "-", `=${Campanha_DB_Campaign_Daily_Budget}/100`, "-", "-", "-", "-", "-", "-", "-" ]]})

                });

            });

            await new Promise(resolve => setTimeout(resolve, 2000));

        } else {

            await new Promise(resolve => setTimeout(resolve, 0));

        }

    }

});