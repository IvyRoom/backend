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

////////////////////////////////////////////////////////////////////////////////////////
// Cria as variáveis de interface com o Meta Conversions API.
////////////////////////////////////////////////////////////////////////////////////////

const Meta_Dataset_ID = process.env.META_DATASET_ID;
const Meta_Conversions_API_Access_Token = process.env.META_CONVERSIONS_API_ACCESS_TOKEN;

////////////////////////////////////////////////////////////////////////////////////////
// Cria as variáveis de interface com o API da Pagar.Me.
////////////////////////////////////////////////////////////////////////////////////////

const PagarMe_API_Latest_Version = process.env.PAGARME_API_LATEST_VERSION;
const PagarMe_SecretKey_Base64_Encoded = process.env.PAGARME_SECRETKEY_BASE64_ENCODED;

////////////////////////////////////////////////////////////////////////////////////////
// Cria as variáveis de interface com o API da PagaLeve.
////////////////////////////////////////////////////////////////////////////////////////

const PagaLeve_API_Key = process.env.PAGALEVE_API_KEY;
const PagaLeve_API_Secret = process.env.PAGALEVE_API_SECRET;

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
// Endpoint: Registra junto à Meta a Visualização da Página Principal.
////////////////////////////////////////////////////////////////////////////////////////

app.post('/landingpage/meta/viewcontent', async (req, res) => {
    
    let { 
    
        Meta_Server_Event_Parameter_Event_Name, 
        Meta_Server_Event_Parameter_Event_Time,
        Meta_Server_Event_Parameter_Event_Source_Url,
        Meta_Server_Event_Parameter_Opt_Out, 
        Meta_Server_Event_Parameter_Event_ID, 
        Meta_Server_Event_Parameter_Action_Source, 
        Meta_Server_Event_Parameter_Data_Processing_Options,
        Meta_Customer_Information_Parameter_Country_NotHashed,
        Meta_Customer_Information_Parameter_External_ID_NotHashed,
        Meta_Customer_Information_Parameter_Client_User_Agent,
        Meta_Customer_Information_Parameter_fbc,
        Meta_Customer_Information_Parameter_fbp
    
    } = req.body;

    let Meta_Customer_Information_Parameter_Client_IP_Address = req.headers['x-forwarded-for'] || req.socket.remoteAddress;
    let Meta_Customer_Information_Parameter_Facebook_Page_ID = Meta_Graph_API_Facebook_Page_ID;

    let Meta_Customer_Information_Parameter_Country_Hashed = crypto.createHash('sha256').update(Meta_Customer_Information_Parameter_Country_NotHashed).digest('hex');
    let Meta_Customer_Information_Parameter_External_ID_Hashed = crypto.createHash('sha256').update(Meta_Customer_Information_Parameter_External_ID_NotHashed).digest('hex');

    res.status(200).json();

    ///////////////////////////////////////////////////////////////////////////////////////
    // Envia os dados do visitante ao Meta Conversions API.
    ///////////////////////////////////////////////////////////////////////////////////////
    
    fetch(`https://graph.facebook.com/${Meta_Graph_API_Latest_Version}/${Meta_Dataset_ID}/events`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
            data: [
                {
                    event_name: Meta_Server_Event_Parameter_Event_Name,
                    event_time: Meta_Server_Event_Parameter_Event_Time,
                    event_source_url: Meta_Server_Event_Parameter_Event_Source_Url,
                    opt_out: Meta_Server_Event_Parameter_Opt_Out,
                    event_id: Meta_Server_Event_Parameter_Event_ID,
                    action_source: Meta_Server_Event_Parameter_Action_Source,
                    data_processing_options: Meta_Server_Event_Parameter_Data_Processing_Options,
                    user_data: {
                        country: Meta_Customer_Information_Parameter_Country_Hashed,
                        external_id: Meta_Customer_Information_Parameter_External_ID_Hashed,
                        client_ip_address: Meta_Customer_Information_Parameter_Client_IP_Address,
                        client_user_agent: Meta_Customer_Information_Parameter_Client_User_Agent,
                        fbc: Meta_Customer_Information_Parameter_fbc,
                        fbp: Meta_Customer_Information_Parameter_fbp,
                        page_id: Meta_Customer_Information_Parameter_Facebook_Page_ID
                    }
                }
            ],
            access_token: Meta_Conversions_API_Access_Token
        })
    })

    .then(response => {
        console.log(response.status);
    })
    
    .catch(error => {
        console.error(error);
    });

});

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
        
            .post({"values": [[ConverteData2(new Date()), Lead_Email, Lead_NomeCompleto, "TURMA #1 2025"]]})

            .then(() => {

                res.status(200).send();

            });

    })

    .catch(error => {
        
        res.status(400).send();

    });

});

////////////////////////////////////////////////////////////////////////////////////////
// Endpoint: Registra junto à Meta a Geração do Lead.
////////////////////////////////////////////////////////////////////////////////////////

app.post('/landingpage/meta/lead', async (req, res) => {
    
    let { 
    
        Meta_Server_Event_Parameter_Event_Name, 
        Meta_Server_Event_Parameter_Event_Time,
        Meta_Server_Event_Parameter_Event_Source_Url,
        Meta_Server_Event_Parameter_Opt_Out, 
        Meta_Server_Event_Parameter_Event_ID, 
        Meta_Server_Event_Parameter_Action_Source, 
        Meta_Server_Event_Parameter_Data_Processing_Options,

        Meta_Customer_Information_Parameter_Email_NotHashed,
        Meta_Customer_Information_Parameter_First_Name_NotHashed,
        Meta_Customer_Information_Parameter_Last_Name_NotHashed,
        Meta_Customer_Information_Parameter_Country_NotHashed,
        Meta_Customer_Information_Parameter_External_ID_NotHashed,
        Meta_Customer_Information_Parameter_Client_User_Agent,
        Meta_Customer_Information_Parameter_fbc,
        Meta_Customer_Information_Parameter_fbp
    
    } = req.body;

    let Meta_Customer_Information_Parameter_Client_IP_Address = req.headers['x-forwarded-for'] || req.socket.remoteAddress;
    let Meta_Customer_Information_Parameter_Facebook_Page_ID = Meta_Graph_API_Facebook_Page_ID;

    let Meta_Customer_Information_Parameter_Email_Hashed = crypto.createHash('sha256').update(Meta_Customer_Information_Parameter_Email_NotHashed).digest('hex');
    let Meta_Customer_Information_Parameter_First_Name_Hashed = crypto.createHash('sha256').update(Meta_Customer_Information_Parameter_First_Name_NotHashed).digest('hex');
    let Meta_Customer_Information_Parameter_Last_Name_Hashed = crypto.createHash('sha256').update(Meta_Customer_Information_Parameter_Last_Name_NotHashed).digest('hex');
    let Meta_Customer_Information_Parameter_Country_Hashed = crypto.createHash('sha256').update(Meta_Customer_Information_Parameter_Country_NotHashed).digest('hex');
    let Meta_Customer_Information_Parameter_External_ID_Hashed = crypto.createHash('sha256').update(Meta_Customer_Information_Parameter_External_ID_NotHashed).digest('hex');
    
    res.status(200).json();

    ///////////////////////////////////////////////////////////////////////////////////////
    // Envia os dados do Lead ao Meta Conversions API.
    ///////////////////////////////////////////////////////////////////////////////////////
    
    fetch(`https://graph.facebook.com/${Meta_Graph_API_Latest_Version}/${Meta_Dataset_ID}/events`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
            data: [
                {
                    event_name: Meta_Server_Event_Parameter_Event_Name,
                    event_time: Meta_Server_Event_Parameter_Event_Time,
                    event_source_url: Meta_Server_Event_Parameter_Event_Source_Url,
                    opt_out: Meta_Server_Event_Parameter_Opt_Out,
                    event_id: Meta_Server_Event_Parameter_Event_ID,
                    action_source: Meta_Server_Event_Parameter_Action_Source,
                    data_processing_options: Meta_Server_Event_Parameter_Data_Processing_Options,
                    user_data: {
                        em: Meta_Customer_Information_Parameter_Email_Hashed,
                        fn: Meta_Customer_Information_Parameter_First_Name_Hashed,
                        ln: Meta_Customer_Information_Parameter_Last_Name_Hashed,
                        country: Meta_Customer_Information_Parameter_Country_Hashed,
                        external_id: Meta_Customer_Information_Parameter_External_ID_Hashed,
                        client_ip_address: Meta_Customer_Information_Parameter_Client_IP_Address,
                        client_user_agent: Meta_Customer_Information_Parameter_Client_User_Agent,
                        fbc: Meta_Customer_Information_Parameter_fbc,
                        fbp: Meta_Customer_Information_Parameter_fbp,
                        page_id: Meta_Customer_Information_Parameter_Facebook_Page_ID
                    }
                }
            ],
            access_token: Meta_Conversions_API_Access_Token
        })
    })

    .then(response => {
        console.log(response.status);
    })
    
    .catch(error => {
        console.error(error);
    });

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

        Valor_Nominal_da_Compra_no_PIX_PARCELADO,
        Valor_Nominal_da_Compra_no_PIX_PARCELADO_Dígitos,
        Url_Aprovação_PIX_PARCELADO,
        Url_Cancelamento_PIX_PARCELADO,
        Url_Webhook_PIX_PARCELADO,

        Valor_Total_da_Compra_no_PIX_À_VISTA,
        Valor_Total_da_Compra_no_PIX_À_VISTA_Dígitos,

        Valor_Total_da_Compra_no_BOLETO,
        Valor_Total_da_Compra_no_BOLETO_Dígitos,

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

        await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/drive/items/0172BBJB2FRGLQQA7KHNCYUNGTVRU3HTG7/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows')
        
        .post({"values": [[ ConverteData2(new Date()), NomeCompleto, Email_do_Cliente, Campo_de_Preenchimento_DDD, Campo_de_Preenchimento_Celular, Campo_de_Preenchimento_CPF, Endereço_Rua, Endereço_Número, Endereço_Complemento, Endereço_Bairro, Endereço_Cidade, Endereço_Estado, Endereço_CEP, Nome_Produto, Tipo_de_Pagamento_Escolhido, Valor_Total_da_Compra_com_Juros_UM_CARTAO, Valor_Total_da_Compra_com_Juros_UM_CARTAO,  Número_do_Cartão, Nome_do_Titular_do_Cartão_CaracteresOriginais, Campo_de_Preenchimento_Mês_Cartão, Campo_de_Preenchimento_Ano_Cartão, Campo_de_Preenchimento_CVV_Cartão, Número_de_Parcelas_Cartão_do_UM_CARTAO, "-", "-", "-", "-", "-", "-", "-", "-", "-", "-", "-" ]]})  

        .then(async (response) => {

            res.status(200).json({});

            let Número_Linha_Adicionada_à_BD_Cobranças = response.index;

            ////////////////////////////////////////////////////////////////////////////////////////
            // Envia o e-mail de "Novo Pedido Gerado no Checkout".

            if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

            await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/sendMail').post({
                message: {
                    subject: 'Novo Pedido Gerado no Checkout',
                    body: {
                        contentType: 'HTML',
                        content: `
                            <p><b>Aluno:</b> ${NomeCompleto}</p>
                            <p>Atenciosamente,</p>
                            <p><img src="https://plataforma-backend-v3.azurewebsites.net/img/ASSINATURA_E-MAIL.png"/></p>
                        `
                    },
                    toRecipients: [{ emailAddress: { address: 'contato@ivyroom.com.br' } }]
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
                        country_code: 55,
                        area_code: Campo_de_Preenchimento_DDD,
                        number: Campo_de_Preenchimento_Celular_Dígitos
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

                await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/drive/items/0172BBJB2FRGLQQA7KHNCYUNGTVRU3HTG7/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows/itemAt(index=' + Número_Linha_Adicionada_à_BD_Cobranças + ')').update({values: [[null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, Retorno_Processamento_Cobrança_PagarMe, Status_Cobrança_Cartão, null, null, null, null, null, null, null, null, null ]]});

            });

        });

    }

    ////////////////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////////////////
    // MODALIDADE DE PAGAMENTO: PIX_PARCELADO
    ////////////////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////////////////

    if (Tipo_de_Pagamento_Escolhido === "PIX_PARCELADO") {
        
        ////////////////////////////////////////////////////////////////////////////////////////
        // Direciona à análise de crédito junto à PagaLeve.
        ////////////////////////////////////////////////////////////////////////////////////////

        ////////////////////////////////////////////////////////////////////////////////////////
        // Obtém o Access Token junto à PagaLeve (Endpoint: Criar uma Sessão Segura).

        fetch('https://api.pagaleve.com.br/v1/authentication', {
            method: 'POST',
            headers: {accept: 'application/json', 'content-type': 'application/json'},
            body: JSON.stringify({
                password: PagaLeve_API_Secret,
                username: PagaLeve_API_Key
            })
        })

        .then(response => response.json()).then(async json => {

            let PagaLeve_Session_Token = 'Bearer ' + json.token;

            let PagaLeve_Session_IdempotencyKey = uuidv4();
            
            let [PrimeiroNome, ...rest] = NomeCompleto.split(" "), Sobrenome = rest.join(" ");

            let DDD_Telefone = Campo_de_Preenchimento_DDD + Campo_de_Preenchimento_Celular_Dígitos;

            ////////////////////////////////////////////////////////////////////////////////////////
            // Cria o "Checkout PagaLeve" (Endpoint: Criar Checkout).

            fetch('https://api.pagaleve.com.br/v1/checkouts', {
                method: 'POST',
                headers: {
                    accept: 'application/json',
                    'Idempotency-Key': PagaLeve_Session_IdempotencyKey,
                    'content-type': 'application/json',
                    authorization: PagaLeve_Session_Token
                },
                body: JSON.stringify({
                    webhook_url: Url_Webhook_PIX_PARCELADO,
                    order: {
                        amount: parseInt(Valor_Nominal_da_Compra_no_PIX_PARCELADO_Dígitos),
                        reference: PagaLeve_Session_IdempotencyKey
                    },
                    shopper: {
                        cpf: Campo_de_Preenchimento_CPF_Dígitos,
                        email: Email_do_Cliente,
                        first_name: PrimeiroNome,
                        last_name: Sobrenome,
                        phone: DDD_Telefone
                    },
                    metadata: {
                        full_name: NomeCompleto,
                        email: Email_do_Cliente,
                        ddd: Campo_de_Preenchimento_DDD,
                        phone: Campo_de_Preenchimento_Celular,
                        cpf: Campo_de_Preenchimento_CPF,
                        shipping_address: {
                            street: Endereço_Rua,
                            number: Endereço_Número,
                            complement: Endereço_Complemento,
                            neighborhood: Endereço_Bairro,
                            city: Endereço_Cidade,
                            state: Endereço_Estado,
                            zip_code: Endereço_CEP
                        },
                        product: Nome_Produto,
                        amount: Valor_Nominal_da_Compra_no_PIX_PARCELADO
                    },
                    approve_url: Url_Aprovação_PIX_PARCELADO,
                    cancel_url: Url_Cancelamento_PIX_PARCELADO,
                    is_pix_upfront: false
                })
            })
            
            .then(response => response.json()).then(async json => {

                let PagaLeve_Checkout_URL = json.checkout_url;
                
                res.status(200).json({ PagaLeve_Checkout_URL });

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

        await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/drive/items/0172BBJB2FRGLQQA7KHNCYUNGTVRU3HTG7/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows')
        
        .post({"values": [[ ConverteData2(new Date()), NomeCompleto, Email_do_Cliente, Campo_de_Preenchimento_DDD, Campo_de_Preenchimento_Celular, Campo_de_Preenchimento_CPF, Endereço_Rua, Endereço_Número, Endereço_Complemento, Endereço_Bairro, Endereço_Cidade, Endereço_Estado, Endereço_CEP, Nome_Produto, Tipo_de_Pagamento_Escolhido, Valor_Total_da_Compra_no_PIX_À_VISTA, "-", "-", "-", "-", "-", "-", "-", "-", "-", Valor_Total_da_Compra_no_PIX_À_VISTA, "-", "-", "-", "-", "-", "-", "-", "-" ]]})  

        .then(async (response) => {

            res.status(200).json({});

            let Número_Linha_Adicionada_à_BD_Cobranças = response.index;

            ////////////////////////////////////////////////////////////////////////////////////////
            // Envia o e-mail de "Novo Pedido Gerado no Checkout".

            if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

            await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/sendMail').post({
                message: {
                    subject: 'Novo Pedido Gerado no Checkout',
                    body: {
                        contentType: 'HTML',
                        content: `
                            <p><b>Aluno:</b> ${NomeCompleto}</p>
                            <p>Atenciosamente,</p>
                            <p><img src="https://plataforma-backend-v3.azurewebsites.net/img/ASSINATURA_E-MAIL.png"/></p>
                        `
                    },
                    toRecipients: [{ emailAddress: { address: 'contato@ivyroom.com.br' } }]
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
                        country_code: 55,
                        area_code: Campo_de_Preenchimento_DDD,
                        number: Campo_de_Preenchimento_Celular_Dígitos
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

                await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/drive/items/0172BBJB2FRGLQQA7KHNCYUNGTVRU3HTG7/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows/itemAt(index=' + Número_Linha_Adicionada_à_BD_Cobranças + ')').update({values: [[null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, Retorno_Processamento_Cobrança_PagarMe, Status_Cobrança_Pix, null, null, null, null, null, null ]]})

                .then(async (response) => {

                    ////////////////////////////////////////////////////////////////////////////////////////
                    // Envia as instruções de pagamento via PIX ao comprador.

                    if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

                    await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/sendMail').post({
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

                                    <img src="https://plataforma-backend-v3.azurewebsites.net/img/ASSINATURA_E-MAIL.png"/>

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
    // MODALIDADE DE PAGAMENTO: BOLETO
    ////////////////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////////////////

    ////////////////////////////////////////////////////////////////////////////////////////
    // Insere o pedido na BD - PEDIDOS.

    if (Tipo_de_Pagamento_Escolhido === "BOLETO") {

        if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

        await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/drive/items/0172BBJB2FRGLQQA7KHNCYUNGTVRU3HTG7/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows')
        
        .post({"values": [[ ConverteData2(new Date()), NomeCompleto, Email_do_Cliente, Campo_de_Preenchimento_DDD, Campo_de_Preenchimento_Celular, Campo_de_Preenchimento_CPF, Endereço_Rua, Endereço_Número, Endereço_Complemento, Endereço_Bairro, Endereço_Cidade, Endereço_Estado, Endereço_CEP, Nome_Produto, Tipo_de_Pagamento_Escolhido, Valor_Total_da_Compra_no_BOLETO, "-", "-", "-", "-", "-", "-", "-", "-", "-", "-", "-", "-", Valor_Total_da_Compra_no_BOLETO, "-", "-", "-", "-", "-" ]]})  

        .then(async (response) => {

            res.status(200).json({});

            let Número_Linha_Adicionada_à_BD_Cobranças = response.index;

            let Boleto_Prazo_Vencimento_Processamento_PagarMe = new Date(Date.now() + 86400000);

            ////////////////////////////////////////////////////////////////////////////////////////
            // Envia o e-mail de "Novo Pedido Gerado no Checkout".

            if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

            await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/sendMail').post({
                message: {
                    subject: 'Novo Pedido Gerado no Checkout',
                    body: {
                        contentType: 'HTML',
                        content: `
                            <p><b>Aluno:</b> ${NomeCompleto}</p>
                            <p>Atenciosamente,</p>
                            <p><img src="https://plataforma-backend-v3.azurewebsites.net/img/ASSINATURA_E-MAIL.png"/></p>
                        `
                    },
                    toRecipients: [{ emailAddress: { address: 'contato@ivyroom.com.br' } }]
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
                        amount: Valor_Total_da_Compra_no_BOLETO_Dígitos, 
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
                        country_code: 55,
                        area_code: Campo_de_Preenchimento_DDD,
                        number: Campo_de_Preenchimento_Celular_Dígitos
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
                        payment_method: 'boleto',
                        boleto: {
                            instructions: 'Não aceitar o pagamento após o vencimento. A emissão deste boleto foi solicitada e/ou intermediada pela empresa IVY ROOM LTDA - CNPJ: 39.794.363/0001-81.',
                            due_at: Boleto_Prazo_Vencimento_Processamento_PagarMe
                        }
                    }]
                })
            })

            .then(response => response.json()).then(async json => {

                let Retorno_Processamento_Cobrança_PagarMe = JSON.stringify(json);

                let Status_Cobrança_Boleto = json.charges?.[0]?.status ?? '-';

                let Boleto_Url_Download = json.charges?.[0]?.last_transaction?.url ?? '-';

                let Boleto_Prazo_Vencimento_Email = ConverteData3(new Date(Date.now() + 86400000));

                ////////////////////////////////////////////////////////////////////////////////////////
                // Insere o Retorno_Processamento_Cobrança_PagarMe e o Status_Cobrança_Boleto na BD - PEDIDOS.

                if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

                await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/drive/items/0172BBJB2FRGLQQA7KHNCYUNGTVRU3HTG7/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows/itemAt(index=' + Número_Linha_Adicionada_à_BD_Cobranças + ')').update({values: [[null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, Retorno_Processamento_Cobrança_PagarMe, Status_Cobrança_Boleto, null, null, null ]]})

                .then(async (response) => {

                    ////////////////////////////////////////////////////////////////////////////////////////
                    // Envia as instruções de pagamento via Boleto ao comprador.

                    if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

                    await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/sendMail').post({
                        message: {
                            subject: 'Ivy: Boleto para Pagamento',
                            body: {
                                contentType: 'HTML',
                                content: `
                                    <html>
                                    <head>
                                        <style type="text/css">
                                    
                                        #Container_Boleto {
                                            height: 110px;
                                            width: 250px;
                                            padding-top: 30px;
                                            padding-bottom: 30px;
                                            padding-left: 30px;
                                            padding-right: 30px;
                                            border: 2px solid rgb(149, 201, 63);
                                            border-radius: 15px;
                                        }
                                        
                                    #Container_Valor_Total_da_Compra_no_Boleto {
                                        display: inline-flex;
                                        margin-bottom: 5px;
                                        font-size: 16px;
                                    }
                                    
                                    #Título_Valor_Total_da_Compra_no_Boleto {
                                        margin-right: 5px;
                                    }
                                    
                                    #Container_Validade_Boleto {
                                        display: inline-flex;
                                        margin-bottom: 5px;
                                        font-size: 16px;
                                    }
                                    
                                    #Título_Validade_Boleto {
                                        margin-right: 5px;
                                    }

                                    #Container_Link_Boleto {
                                        display: inline-flex;
                                        margin-bottom: 10px;
                                        font-size: 16px;
                                    }
                                    
                                    #Título_Link_Boleto {
                                        margin-right: 5px;
                                    }
                                    
                                    #Container_Logo_PagarMe {
                                        display: inline-flex;
                                        margin-left: 35px;
                                        margin-right: auto;
                                    }
                                    
                                    #PoweredBy {
                                        font-size: 10px;
                                        margin-right: 3px;
                                        margin-top: 18px;
                                        width: 60px;
                                    }

                                    #Logo_PagarMe{
                                        width: 110px;
                                        height: 40px;
                                    }
                                        
                                    </style>
                                    </head>
                                    <body>
                                        <p>Prezado(a) ${PrimeiroNome},<br></p>
                                        <p>Para concluir a compra do <b> ${Nome_Produto}</b>, faça o pagamento do boleto utilizando o link abaixo.<br></p>
                                        <p>As orientações de acesso ao serviço serão enviadas para você por e-mail, assim que o pagamento for processado.<br></p>

                                        <div id="Container_Boleto">
                                            <div id="Container_Valor_Total_da_Compra_no_Boleto">
                                                    <div id="Título_Valor_Total_da_Compra_no_Boleto"><b>Valor:</b></div>
                                                    <div id="Valor_Total_da_Compra_no_Boleto"> ${Valor_Total_da_Compra_no_BOLETO}</div>
                                            </div>
                                            <div id="Container_Validade_Boleto">
                                                <div id="Título_Validade_Boleto"><b>Validade:</b></div>
                                                <div id="Validade_Boleto"> ${Boleto_Prazo_Vencimento_Email}</div>
                                            </div>
                                            <div id="Container_Link_Boleto">
                                                <div id="Título_Link_Boleto"><b>Link de Acesso:</b></div>
                                                <a id="url_Boleto" href="${Boleto_Url_Download}">Boleto</a>
                                            </div>
                                            <div id="Container_Logo_PagarMe">
                                                    <p id="PoweredBy">Powered by</p>
                                                    <img id="Logo_PagarMe" src="https://plataforma-backend-v3.azurewebsites.net/img/LOGO_PAGAR.ME.png"/>
                                            </div>
                                        </div>
                                        
                                    <p><br>Por favor entre em contato se surgirem dúvidas ou se precisar de auxílio.<br></p>

                                    <p>Estamos sempre à disposição.<br></p>

                                    <p>Atenciosamente,<br></p>

                                    <p><img src="https://plataforma-backend-v3.azurewebsites.net/img/ASSINATURA_E-MAIL.png"/></p>

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

        await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/drive/items/0172BBJB2FRGLQQA7KHNCYUNGTVRU3HTG7/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows')
        
        .post({"values": [[ ConverteData2(new Date()), NomeCompleto, Email_do_Cliente, Campo_de_Preenchimento_DDD, Campo_de_Preenchimento_Celular, Campo_de_Preenchimento_CPF, Endereço_Rua, Endereço_Número, Endereço_Complemento, Endereço_Bairro, Endereço_Cidade, Endereço_Estado, Endereço_CEP, Nome_Produto, Tipo_de_Pagamento_Escolhido, Valor_Total_da_Compra_no_PIX_CARTÃO, Valor_com_Juros_no_Cartão_do_PIX_CARTÃO, Número_do_Cartão_do_PIX_CARTÃO, Nome_do_Titular_do_Cartão_do_PIX_CARTÃO_CaracteresOriginais, Campo_de_Preenchimento_Mês_Cartão_do_PIX_CARTÃO, Campo_de_Preenchimento_Ano_Cartão_do_PIX_CARTÃO, Campo_de_Preenchimento_CVV_Cartão_do_PIX_CARTÃO, Número_de_Parcelas_Cartão_do_PIX_CARTÃO, "-", "-", Valor_no_PIX_do_PIX_CARTÃO, "-", "-", "-", "-", "-", "-", "-", "-" ]]})  

        .then(async (response) => {

            res.status(200).json({});

            let Número_Linha_Adicionada_à_BD_Cobranças = response.index;

            ////////////////////////////////////////////////////////////////////////////////////////
            // Envia o e-mail de "Novo Pedido Gerado no Checkout".

            if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

            await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/sendMail').post({
                message: {
                    subject: 'Novo Pedido Gerado no Checkout',
                    body: {
                        contentType: 'HTML',
                        content: `
                            <p><b>Aluno:</b> ${NomeCompleto}</p>
                            <p>Atenciosamente,</p>
                            <p><img src="https://plataforma-backend-v3.azurewebsites.net/img/ASSINATURA_E-MAIL.png"/></p>
                        `
                    },
                    toRecipients: [{ emailAddress: { address: 'contato@ivyroom.com.br' } }]
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
                        country_code: 55,
                        area_code: Campo_de_Preenchimento_DDD,
                        number: Campo_de_Preenchimento_Celular_Dígitos
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

                await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/drive/items/0172BBJB2FRGLQQA7KHNCYUNGTVRU3HTG7/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows/itemAt(index=' + Número_Linha_Adicionada_à_BD_Cobranças + ')').update({values: [[null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, Retorno_Processamento_Cobrança_Cartão_PIX_CARTAO_PagarMe, Status_Cobrança_Cartão_PIX_CARTAO, null, null, null, null, null, null, null, null, null ]]})

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
                                country_code: 55,
                                area_code: Campo_de_Preenchimento_DDD,
                                number: Campo_de_Preenchimento_Celular_Dígitos
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

                        await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/drive/items/0172BBJB2FRGLQQA7KHNCYUNGTVRU3HTG7/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows/itemAt(index=' + Número_Linha_Adicionada_à_BD_Cobranças + ')').update({values: [[null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, Retorno_Processamento_Cobrança_Pix_PIX_CARTAO_PagarMe, Status_Cobrança_Pix, null, null, null, null, null, null ]]})

                        .then(async (response) => {

                            ////////////////////////////////////////////////////////////////////////////////////////
                            // Envia as instruções de pagamento via PIX ao comprador.

                            if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

                            await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/sendMail').post({
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

                                            <img src="https://plataforma-backend-v3.azurewebsites.net/img/ASSINATURA_E-MAIL.png"/>

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


////////////////////////////////////////////////////////////////////////////////////////
// Webhook: Processa Pedido (Checkout) Autorizado pela PagaLeve.
////////////////////////////////////////////////////////////////////////////////////////

app.post('/checkout/webhook_pagaleve', async (req, res) => {

    res.status(200).send();
    
    let PagaLeve_Checkout_Status = req.body.state;
    
    ////////////////////////////////////////////////////////////////////////////////////////
    // Insere o pedido na BD - PEDIDOS.

    if (PagaLeve_Checkout_Status === "CANCELED") { // "AUTHORIZED"

        let PagaLeve_Checkout_Response = req.body;
        let PagaLeve_Checkout_ID = req.body.id;
        let PagaLeve_Checkout_Amount = req.body.amount;
        let NomeCompleto = req.body.metadata.full_name;
        let Email_do_Cliente = req.body.metadata.email;
        let Campo_de_Preenchimento_DDD = req.body.metadata.ddd;
        let Campo_de_Preenchimento_Celular = req.body.metadata.phone;
        let Campo_de_Preenchimento_CPF = req.body.metadata.cpf;
        let Endereço_Rua = req.body.metadata.shipping_address.street;
        let Endereço_Número = req.body.metadata.shipping_address.number;
        let Endereço_Complemento = req.body.metadata.shipping_address.complement;
        let Endereço_Bairro = req.body.metadata.shipping_address.neighborhood;
        let Endereço_Cidade = req.body.metadata.shipping_address.city;
        let Endereço_Estado = req.body.metadata.shipping_address.state;
        let Endereço_CEP = req.body.metadata.shipping_address.zip_code;
        let Nome_Produto = req.body.metadata.product;
        let Tipo_de_Pagamento_Escolhido = "PIX_PARCELADO"
        let Valor_Nominal_da_Compra_no_PIX_PARCELADO = req.body.metadata.amount;

        if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

        await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/drive/items/0172BBJB2FRGLQQA7KHNCYUNGTVRU3HTG7/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows')
        
        .post({"values": [[ ConverteData2(new Date()), NomeCompleto, Email_do_Cliente, Campo_de_Preenchimento_DDD, Campo_de_Preenchimento_Celular, Campo_de_Preenchimento_CPF, Endereço_Rua, Endereço_Número, Endereço_Complemento, Endereço_Bairro, Endereço_Cidade, Endereço_Estado, Endereço_CEP, Nome_Produto, Tipo_de_Pagamento_Escolhido, Valor_Nominal_da_Compra_no_PIX_PARCELADO, "-", "-", "-", "-", "-", "-", "-", "-", "-", "-", "-", "-", "-", "-", "-", "-", PagaLeve_Checkout_Response, PagaLeve_Checkout_Status ]]})  

        .then(async (response) => {

            res.status(200).send();

            ////////////////////////////////////////////////////////////////////////////////////////
            // Envia o e-mail de "Novo Pedido Gerado no Checkout".

            if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

            await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/sendMail').post({
                message: {
                    subject: 'Novo Pedido Gerado no Checkout',
                    body: {
                        contentType: 'HTML',
                        content: `
                            <p><b>Aluno:</b> ${NomeCompleto}</p>
                            <p>Atenciosamente,</p>
                            <p><img src="https://plataforma-backend-v3.azurewebsites.net/img/ASSINATURA_E-MAIL.png"/></p>
                        `
                    },
                    toRecipients: [{ emailAddress: { address: 'contato@ivyroom.com.br' } }]
                }
            });

            ////////////////////////////////////////////////////////////////////////////////////////
            // Obtém o Access Token junto à PagaLeve (Endpoint: Criar uma Sessão Segura).

            fetch('https://api.pagaleve.com.br/v1/authentication', {
                method: 'POST',
                headers: {accept: 'application/json', 'content-type': 'application/json'},
                body: JSON.stringify({
                    password: PagaLeve_API_Secret,
                    username: PagaLeve_API_Key
                })
            })

            .then(response => response.json()).then(async json => {

                let PagaLeve_Session_Token = 'Bearer ' + json.token;

                let PagaLeve_Session_IdempotencyKey = uuidv4();

                ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                // Processa o Pagamento referente ao Pedido (Checkout) Autorizado junto à PagaLeve (Endpoint: Criar Pagamento).

                fetch('https://api.pagaleve.com.br/v1/payments', {
                    method: 'POST',
                    headers: {
                        accept: 'application/json',
                        'Idempotency-Key': PagaLeve_Session_IdempotencyKey,
                        'content-type': 'application/json',
                        authorization: PagaLeve_Session_Token
                    },
                    body: JSON.stringify({
                        currency: 'BRL',
                        intent: 'CAPTURE',
                        amount: PagaLeve_Checkout_Amount,
                        checkout_id: PagaLeve_Checkout_ID
                    })
                })

                .then(response => response.json()).then(async json => {

                    console.log(json);

                });
                
            });

        });

    }
    
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

    let { Data_Abertura_Turma } = req.body;

    let [Dia_Abertura_Turma,Mês_Abertura_Turma,Ano_Abertura_Turma] = Data_Abertura_Turma.split("/").map(num => parseInt(num, 10));

    let Dia_da_Semana_Data_Abertura_Turma = new Intl.DateTimeFormat('pt-BR', { weekday: 'long' }).format(new Date(Ano_Abertura_Turma, Mês_Abertura_Turma - 1, Dia_Abertura_Turma));
    
    ////////////////////////////////////////////////////////////////////////////////////////
    // Puxa os dados da BD - LEADS.
    
    if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

    const BD_Leads = await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/drive/items/0172BBJBYG24NEFOMGOJCLN5FMDILTSZTC/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{AC8C07F3-9A79-4ABD-8CE8-0C818B0EA1A7}/rows').get();

    const BD_Leads_Última_Linha = BD_Leads.value.length - 1;

    async function Envia_Email_Leads() {
    
        for (let LinhaAtual = 0; LinhaAtual <= BD_Leads_Última_Linha; LinhaAtual++) {
            
            Lead_Email = BD_Leads.value[LinhaAtual].values[0][1];
            Lead_PrimeiroNome = BD_Leads.value[LinhaAtual].values[0][2].split(" ")[0];

            ///////////////////////////////////////////////////////////////////////////////////////////////////
            // Cria o evento iCalendar para a Abertura de Turma, com alerta de 1 hora antes do início do encontro.

            const cal = new ICalCalendar({ domain: 'ivyroom.com.br', prodId: { company: 'Ivy | Escola de Gestão', product: 'Ivy - Abertura de Turma: Preparatório em Gestão Generalista', language: 'PT-BR' } });
            const event = cal.createEvent({
                start: new Date(Date.UTC(Ano_Abertura_Turma, Mês_Abertura_Turma - 1, Dia_Abertura_Turma + 1, 1, 0, 0)), // 22:00 BRT
                end: new Date(Date.UTC(Ano_Abertura_Turma, Mês_Abertura_Turma - 1, Dia_Abertura_Turma + 1, 2, 0, 0)), // 23:00 BRT
                summary: 'Abertura de Turma',
                description: 
`A próxima turma do Preparatório em Gestão Generalista abre 10/abril/2025 (quinta-feira) às 22:00, no horário de Brasília.
Acesse este link para adquirir o serviço: https://ivygestao.com/`,
                uid: `${new Date().getTime()}@ivyroom.com.br`,
                stamp: new Date()
            });

            event.createAlarm({
                type: 'display',
                trigger: 1 * 60 * 60 * 1000,
                description: 'Ivy - Abertura de Turma: Falta 1 hora.'
            });

            ////////////////////////////////////////////////////////////////////////////////////////
            // Envia o e-mail para o lead na LinhaAtual da BD - ALUNOS.

            if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

            await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/sendMail').post({

                message: {
                    subject: 'Ivy - Próxima Turma do Preparatório em Gestão Generalista: 10/abril 22:00',
                    body: {
                        contentType: 'HTML',
                        content: `
                            <p>Olá ${Lead_PrimeiroNome},</p>
                            <p>É com satisfação que informamos que a data de abertura da próxima turma do Preparatório em Gestão Generalista foi definida para:</p> 
                            <p><b>10/abril/2025 (${Dia_da_Semana_Data_Abertura_Turma}) às 22:00</b>, no horário de Brasília.</p>
                            <p>Abra o arquivo .ics em anexo para adicionar o evento a sua agenda.</p>
                            <p>No dia da abertura da turma, você poderá comprar o Preparatório por meio deste link: <a href="https://ivygestao.com/" target="_blank">https://ivygestao.com/</a></p>
                            <p>Precisamente às 22:00, o botão de "Entrar na Lista de Espera" será substituído pelo botão de compra, que dará acesso ao nosso checkout com diversas modalidades de pagamento e parcelamento.</p>
                            <p>Lembrando que ofereceremos um <b>bônus exclusivo</b>, muito diferenciado, aos <b>50 primeiros alunos</b>.</p>
                            <p>Os detalhes também estão <a href="https://ivygestao.com/" target="_blank">neste link</a>.</p>
                            <p>P.S. Há pouco postamos stories em <a href="https://www.instagram.com/ivy.escoladegestao/" target="_blank">nosso instagram</a> explicando os principais pontos sobre a abertura de turma. Vale conferir.</p>
                            <p>Qualquer dúvida ou insegurança, sempre à disposição.</p>
                            <p>Atenciosamente,</p>
                            <p><img src="https://plataforma-backend-v3.azurewebsites.net/img/ASSINATURA_E-MAIL.png"/></p>
                        `
                    },
                    toRecipients: [{ emailAddress: { address: Lead_Email } }],
                    attachments: [
                        {
                            "@odata.type": "#microsoft.graph.fileAttachment",
                            name: "Ivy - Abertura de Turma.ics",
                            contentBytes: Buffer.from(cal.toString()).toString('base64')
                        }
                    ]
                }
            
            });

            await new Promise(resolve => setTimeout(resolve, 2000));

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
    
    // Primeira linha Black Friday: 204.
    // Última linha Black Friday: 237.

    for (let LinhaAtual = 206; LinhaAtual <= 237; LinhaAtual++) {
        
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

        Aluno_Login = BD_Alunos.value[LinhaAtual].values[0][12];
        Aluno_Senha = BD_Alunos.value[LinhaAtual].values[0][13];

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

            // await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/sendMail').post({
            //     message: {
            //         subject: 'Ivy - É hoje! Link de Acesso: Atendimento ao Vivo #1',
            //         body: {
            //             contentType: 'HTML',
            //             content: `
            //                 <p>Boa tarde ${Aluno_PrimeiroNome},</p>
            //                 <p>Lembramos que <b>hoje</b>, pontualmente às <b>18h</b>, acontece o <b>primeiro atendimento ao vivo</b> para os participantes da Turma de Black Friday do Preparatório, via Microsoft Teams.</p>
            //                 <p>Neste encontro, o Lucas já iniciará algumas explicações importantes sobre Método Gerencial, com conhecimentos complementares ao conteúdo da plataforma. E estará à disposição para tirar dúvidas iniciais sobre o tema.</p>
            //                 <p>Acesse o encontro por meio <a href="https://teams.microsoft.com/l/meetup-join/19%3ameeting_ZGUxZjMwYjktYTU3Ny00Yzc5LTkzYWYtOWU4ZGQ5NDYzM2Y0%40thread.v2/0?context=%7b%22Tid%22%3a%2249342d16-0605-4267-b540-d1fe7756dbac%22%2c%22Oid%22%3a%22b4a93dcf-5946-4cb2-8368-5db4d242a236%22%7d" target="_blank">deste link</a>.</p>
            //                 <p>Qualquer dúvida ou insegurança, sempre à disposição.</p>
            //                 <p>Atenciosamente,</p>
            //                 <p><img src="https://plataforma-backend-v3.azurewebsites.net/img/ASSINATURA_E-MAIL.png"/></p>
            //             `
            //         },
            //         toRecipients: [{ emailAddress: { address: Aluno_Email } }]
            //     }
                
            // })

            await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/sendMail').post({
                message: {
                    subject: 'Encontro Especial de Black Friday - Reagendamento',
                    body: {
                        contentType: 'HTML',
                        content: `
                            <p>Bom dia ${Aluno_PrimeiroNome},</p>
                            <p>Quem escreve é Lucas Machado, fundador da Ivy. Tudo bem?</p>
                            <p>Pelo fato de que menos de 50% dos alunos da Turma de Black Friday já finalizaram o Preparatório, iremos reagendar o Encontro Especial de Black Friday que aconteceria amanhã (sábado, 08/fev às 09:00).</p>
                            <p>Vamos monitorar o progresso da turma na plataforma e reagendaremos o encontro no momento oportuno. Avisaremos vocês por e-mail.</p>
                            <p>P.S. Em instantes sairão os invites para as próximas Office Hours (encontros ao vivo mensais que fazem parte da entrega padrão do serviço).</p>
                            <p>Qualquer dúvida ou insegurança, à disposição.</p>
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

////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////
// Envia convites para as Office Hours.
////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////

app.post('/alunos/conviteOHs', async (req,res) => {

    let { Data_Início_Office_Hours, Link_Microsoft_Teams } = req.body;
    
    res.status(200).json({ message: "1. Request recebida." });

    let [Dia_Início_Office_Hours,Mês_Início_Office_Hours,Ano_Início_Office_Hours] = Data_Início_Office_Hours.split("/").map(num => parseInt(num, 10));

    let Dia_da_Semana_Data_Início_Office_Hours = new Intl.DateTimeFormat('pt-BR', { weekday: 'long' }).format(new Date(Ano_Início_Office_Hours, Mês_Início_Office_Hours - 1, Dia_Início_Office_Hours));

    ////////////////////////////////////////////////////////////////////////////////////////
    // Puxa os dados da BD - OFFICE HOURS.
    
    if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

    const BD_Office_Hours = await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/drive/items/0172BBJB4MCD3537W3HFGZXYMIIMCN5JQ2/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows').get();

    if (BD_Office_Hours !== null && client) client.send(JSON.stringify({ message: `2. BD - OFFICE HOURS obtida.`, origin: "ConviteOHs" }));
    
    const BD_Office_Hours_Última_Linha = BD_Office_Hours.value.length - 1;

    let Número_Invite_Enviado = 0;

    ////////////////////////////////////////////////////////////////////////////////////////
    // Aguarda 1s para iniciar o envio dos e-mails, para que o WebSocket possa enviar os dados de volta ao frontend.
    // Então envia um invite a cada 2s.
    
    async function Envia_Invites_Office_Hours() {

        for (let LinhaAtual = 0; LinhaAtual <= BD_Office_Hours_Última_Linha; LinhaAtual++) {

            ///////////////////////////////////////////////////////////////////////////////////////////////////
            // Puxa as variáveis do aluno da BD - OFFICE HOURS.
    
            Aluno_PrimeiroNome = BD_Office_Hours.value[LinhaAtual].values[0][1].split(" ")[0];
            Aluno_Email = BD_Office_Hours.value[LinhaAtual].values[0][2];
            Aluno_Status_Envio_Convite_Office_Hours = BD_Office_Hours.value[LinhaAtual].values[0][3];
    
            if (Aluno_Status_Envio_Convite_Office_Hours === "SIM") {
    
                Número_Invite_Enviado++;
    
                if (client) client.send(JSON.stringify({ message: `3. Invite #${Número_Invite_Enviado} enviado para: ${Aluno_PrimeiroNome}`, origin: "ConviteOHs" }));
    
                if (LinhaAtual === BD_Office_Hours_Última_Linha && client) client.send(JSON.stringify({ message: `--- fim ---`, origin: "ConviteOHs" }));
    
                ///////////////////////////////////////////////////////////////////////////////////////////////////
                // Cria o evento iCalendar para as Office Hours, com alerta de 1 hora antes do início do encontro.
    
                const cal = new ICalCalendar({ domain: 'ivyroom.com.br', prodId: { company: 'Ivy | Escola de Gestão', product: 'Ivy - Office Hours', language: 'PT-BR' } });
                const event = cal.createEvent({
                    start: new Date(Date.UTC(Ano_Início_Office_Hours, Mês_Início_Office_Hours - 1, Dia_Início_Office_Hours, 21, 30, 0)), // 18:30 BRT
                    end: new Date(Date.UTC(Ano_Início_Office_Hours, Mês_Início_Office_Hours - 1, Dia_Início_Office_Hours, 23, 0, 0)), // 20:00 BRT
                    summary: 'Office Hours',
                    description: ` Link do Encontro (Microsoft Teams): ${Link_Microsoft_Teams}`,
                    uid: `${new Date().getTime()}@ivyroom.com.br`,
                    stamp: new Date()
                });
    
                event.createAlarm({
                    type: 'display',
                    trigger: 1 * 60 * 60 * 1000,
                    description: 'Office Hours (Ivy) - Inicia em 1 hora.'
                });
    
                ////////////////////////////////////////////////////////////////////////////////////////
                // Envia o e-mail para o lead na LinhaAtual da BD - ALUNOS.
    
                if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();
    
                await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/sendMail').post({
    
                    message: {
                        subject: 'Ivy - Convite: Office Hours (Atendimento ao Vivo)',
                        body: {
                            contentType: 'HTML',
                            content: `
                                <p>Olá ${Aluno_PrimeiroNome},</p>
                                <p>Informamos que as próximas Office Hours com o Lucas, nosso fundador, acontecerão <b>${Dia_da_Semana_Data_Início_Office_Hours} (${Data_Início_Office_Hours}) às 18:30</b>, via Microsoft Teams, por meio <a href=${Link_Microsoft_Teams} target="_blank">deste link</a>.</p>
                                <p><b>Por favor abra o arquivo .ics em anexo e adicione o evento a sua agenda.</b></p>
                                <p>Reforçamos que você é o protagonista destes encontros. Por isto, se prepare previamente e traga suas dúvidas, anotações e materiais impressos.</p> 
                                <p>Qualquer dúvida ou insegurança, sempre à disposição.</p>
                                <p>Atenciosamente,</p>
                                <p><img src="https://plataforma-backend-v3.azurewebsites.net/img/ASSINATURA_E-MAIL.png"/></p>
                            `
                        },
                        toRecipients: [{ emailAddress: { address: Aluno_Email } }],
                        attachments: [
                            {
                                "@odata.type": "#microsoft.graph.fileAttachment",
                                name: "Ivy - Office Hours.ics",
                                contentBytes: Buffer.from(cal.toString()).toString('base64')
                            }
                        ]
                    }
                
                });

                await new Promise(resolve => setTimeout(resolve, 2000));
    
            } else {

                await new Promise(resolve => setTimeout(resolve, 0));

                if (LinhaAtual === BD_Office_Hours_Última_Linha && client) client.send(JSON.stringify({ message: `--- fim ---`, origin: "ConviteOHs" }));

            }
    
        }

    }

    setTimeout(Envia_Invites_Office_Hours, 1000);

});

////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////
// Envia lembretes para as Office Hours.
////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////

app.post('/alunos/lembretesOHs', async (req,res) => {

    let { Data_Início_Office_Hours, Link_Microsoft_Teams } = req.body;

    res.status(200).json({ message: "1. Request recebida." });

    console.log('1. Request recebida.');

    let [Dia_Início_Office_Hours,Mês_Início_Office_Hours,Ano_Início_Office_Hours] = Data_Início_Office_Hours.split("/").map(num => parseInt(num, 10));

    let Dia_da_Semana_Data_Início_Office_Hours = new Intl.DateTimeFormat('pt-BR', { weekday: 'long' }).format(new Date(Ano_Início_Office_Hours, Mês_Início_Office_Hours - 1, Dia_Início_Office_Hours));

    ////////////////////////////////////////////////////////////////////////////////////////
    // Puxa os dados da BD - OFFICE HOURS.
    
    if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

    const BD_Office_Hours = await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/drive/items/0172BBJB4MCD3537W3HFGZXYMIIMCN5JQ2/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows').get();

    //if (BD_Office_Hours !== null && client) client.send(JSON.stringify({ message: `2. BD - OFFICE HOURS obtida.`, origin: "LembreteOHs" }));
    
    if (BD_Office_Hours !== null) console.log('2. BD - OFFICE HOURS obtida.');

    const BD_Office_Hours_Última_Linha = BD_Office_Hours.value.length - 1;

    let Número_Lembrete_Enviado = 0;

    ////////////////////////////////////////////////////////////////////////////////////////
    // Aguarda 1s para iniciar o envio dos e-mails, para que o WebSocket possa enviar os dados de volta ao frontend.
    // Então envia um invite a cada 2s.
    
    async function Envia_Lembretes_Office_Hours() {

        for (let LinhaAtual = 0; LinhaAtual <= BD_Office_Hours_Última_Linha; LinhaAtual++) {

            ///////////////////////////////////////////////////////////////////////////////////////////////////
            // Puxa as variáveis do aluno da BD - OFFICE HOURS.
    
            Aluno_PrimeiroNome = BD_Office_Hours.value[LinhaAtual].values[0][1].split(" ")[0];
            Aluno_Email = BD_Office_Hours.value[LinhaAtual].values[0][2];
            Aluno_Status_Envio_Convite_Office_Hours = BD_Office_Hours.value[LinhaAtual].values[0][3];
    
            if (Aluno_Status_Envio_Convite_Office_Hours === "SIM") {
    
                Número_Lembrete_Enviado++;
    
                // if (client) client.send(JSON.stringify({ message: `3. Lembrete #${Número_Lembrete_Enviado} enviado para: ${Aluno_PrimeiroNome}`, origin: "LembreteOHs" }));
                
                console.log(`3. Lembrete #${Número_Lembrete_Enviado} enviado para: ${Aluno_PrimeiroNome}`);

                // if (LinhaAtual === BD_Office_Hours_Última_Linha && client) client.send(JSON.stringify({ message: `--- fim ---`, origin: "LembreteOHs" }));
                
                if (LinhaAtual === BD_Office_Hours_Última_Linha) console.log(`--- fim ---`)
    
                ////////////////////////////////////////////////////////////////////////////////////////
                // Envia o e-mail para o aluno na LinhaAtual da BD - OFFICE HOURS.
    
                if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();
    
                await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/sendMail').post({
    
                    message: {
                        subject: 'Ivy - Lembrete: Office Hours (Atendimento ao Vivo)',
                        body: {
                            contentType: 'HTML',
                            content: `
                                <p>Olá ${Aluno_PrimeiroNome},</p>
                                <p>Reforçamos que as próximas Office Hours com o Lucas, nosso fundador, acontecerão <b>hoje, ${Dia_da_Semana_Data_Início_Office_Hours} (${Data_Início_Office_Hours}) às 18:30</b>, via Microsoft Teams, por meio <a href=${Link_Microsoft_Teams} target="_blank">deste link</a>.</p>
                                <p>Lembramos que você é o protagonista destes encontros. Por isto, se prepare previamente e traga suas dúvidas, anotações e materiais impressos.</p> 
                                <p>Qualquer dúvida ou insegurança, sempre à disposição.</p>
                                <p>Atenciosamente,</p>
                                <p><img src="https://plataforma-backend-v3.azurewebsites.net/img/ASSINATURA_E-MAIL.png"/></p>
                            `
                        },
                        toRecipients: [{ emailAddress: { address: Aluno_Email } }]
                    }
                
                });

                await new Promise(resolve => setTimeout(resolve, 2000));
    
            } else {

                await new Promise(resolve => setTimeout(resolve, 0));

                // if (LinhaAtual === BD_Office_Hours_Última_Linha && client) client.send(JSON.stringify({ message: `--- fim ---`, origin: "LembreteOHs" }));

                if(LinhaAtual === BD_Office_Hours_Última_Linha) console.log(`--- fim ---`);

            }
    
        }

    }

    setTimeout(Envia_Lembretes_Office_Hours, 1000);

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
// Endpoint: Postagem de Reels e Stories
////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////

app.post('/meta/postar', async (req, res) => {

    let { Reel_Código, Reel_Legenda, Incluir_Stories } = req.body;

    ////////////////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////////////////
    // Reels: Início do processo de postagem.
    ////////////////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////////////////

    res.status(200).json({ message: "Reels - 1. Request recebida." });

    ////////////////////////////////////////////////////////////////////////////////////////
    // Cria o Media Container (Reel).

    if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

    let Reel_Video_URL = (await Microsoft_Graph_API_Client.api(`/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/drive/root:/SISTEMA DE GESTÃO/1. VENDA/2. POSTAR REELS/PUBLICAÇÕES/${Reel_Código}/PUBLICAÇÃO.mp4`).get())['@microsoft.graph.downloadUrl'];

    fetch(`https://graph.facebook.com/${Meta_Graph_API_Latest_Version}/${Meta_Graph_API_Instagram_Business_Account_ID}/media`, {
        
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        
        body: JSON.stringify({
            media_type: 'REELS',
            video_url: Reel_Video_URL,
            caption: Reel_Legenda,
            access_token: Meta_Graph_API_Access_Token
        })
    
    })

    .then(response => response.json()).then(data => {

        let Reel_IG_Media_Container_ID = data.id;

        if (Reel_IG_Media_Container_ID !== null && client) client.send(JSON.stringify({ message: `Reels - 2. Media Container ID ${Reel_IG_Media_Container_ID} criado.`, origin: "postar" }));

        ////////////////////////////////////////////////////////////////////////////////////////
        // Verifica o status do Media Container (Reels) a cada 5s.

        const VerificaStatusReels = () => {
            
            fetch(`https://graph.facebook.com/${Meta_Graph_API_Latest_Version}/${Reel_IG_Media_Container_ID}?fields=status_code`, {
                method: 'GET',
                headers: {
                    'Content-Type': 'application/json',
                    'Authorization': `Bearer ${Meta_Graph_API_Access_Token}`
                }
            })

            .then(response => response.json()).then(data => {

                let Reel_IG_Media_Container_Status = data.status_code;
                
                if (Reel_IG_Media_Container_Status === "FINISHED") {

                    clearInterval(Reel_IG_Media_Container_Status_Verificação_ID);

                    if (client) client.send(JSON.stringify({ message: `Reels - 3. Media Container Status atualizado: ${Reel_IG_Media_Container_Status}.`, origin: "postar" }));

                    ////////////////////////////////////////////////////////////////////////////////////////
                    // Publica o Media Container (Reels).
                    ////////////////////////////////////////////////////////////////////////////////////////
                    
                    fetch(`https://graph.facebook.com/${Meta_Graph_API_Latest_Version}/${Meta_Graph_API_Instagram_Business_Account_ID}/media_publish`, {
                        
                        method: 'POST',
                        headers: {'Content-Type': 'application/json'},
                        
                        body: JSON.stringify({
                            creation_id: Reel_IG_Media_Container_ID,
                            access_token: Meta_Graph_API_Access_Token
                        })
                    
                    })
                    
                    .then(response => response.json()).then(data => {
                        
                        let Reel_IG_Media_ID = data.id;

                        if (Reel_IG_Media_ID !== null && client) client.send(JSON.stringify({ message: `Reels - 4. Reel ID ${Reel_IG_Media_ID} publicado.`, origin: "postar" }));

                        ////////////////////////////////////////////////////////////////////////////////////////
                        // Cria a Custom Audience (Video View 3s) para o Reel.
                        ////////////////////////////////////////////////////////////////////////////////////////

                        fetch(`https://graph.facebook.com/${Meta_Graph_API_Latest_Version}/${Meta_Graph_API_Ad_Account_ID}/customaudiences`, {
                            method: 'POST',
                            headers: {'Content-Type': 'application/json'},
                            body: JSON.stringify({
                                name: 'Video View 3s - ' + Reel_Código,
                                description: 'Audiência criada via Meta Graph API.',
                                subtype: 'ENGAGEMENT',
                                retention_days: 365,
                                rule: [{
                                    event_name: 'video_watched',
                                    object_id: Reel_IG_Media_ID,
                                    context_id: Meta_Graph_API_Facebook_Page_ID
                                }],
                                prefill: 'true',
                                access_token: Meta_Graph_API_Access_Token
                            })
                        })
                        
                        .then(response => response.json()).then(data => {

                            let Reel_Audience_ID = data.id;

                            if (Reel_Audience_ID !== null && client) client.send(JSON.stringify({ message: `Reels - 5. Audiência ID ${Reel_Audience_ID} criada.`, origin: "postar" }));

                            ////////////////////////////////////////////////////////////////////////////////////////
                            // Obtém o Número de Seguidores atualizado.
                            ////////////////////////////////////////////////////////////////////////////////////////

                            fetch(`https://graph.facebook.com/${Meta_Graph_API_Latest_Version}/${Meta_Graph_API_Instagram_Business_Account_ID}?fields=followers_count&access_token=${Meta_Graph_API_Access_Token}`, { method: 'GET'})

                            .then(response => response.json()).then(async data => {

                                let Número_Seguidores = data.followers_count;

                                if (Número_Seguidores !== null && client) client.send(JSON.stringify({ message: `Reels - 6. Número de Seguidores obtido: ${Número_Seguidores}`, origin: "postar" }));
                            
                                ///////////////////////////////////////////////////////////////////////////////////////
                                // Adiciona o Reel na BD - RESULTADOS (RELACIONAMENTO).
                                ///////////////////////////////////////////////////////////////////////////////////////

                                if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

                                await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/drive/items/0172BBJB5JTOTCSWCLGBB2HKLEFJVR7AUC/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{122865F8-2E2D-4B60-A34C-E02E001E835E}/rows')
                                
                                .post({"values": [[ Reel_Código, `'${Reel_IG_Media_ID}`, `'${Reel_Audience_ID}`, ConverteData2(new Date()), Número_Seguidores, null, null, null, null ]]})
                                
                                .then(async response => {

                                    if (client) client.send(JSON.stringify({ message: `Reels - 7. BD - RESULTADOS atualizada.`, origin: "postar" }));

                                    /////////////////////////////////////////////////////////////////////////////////////////////////////
                                    // Cria o evento na agenda (calendário) para criação da campanha de RL (72h depois).
                                    /////////////////////////////////////////////////////////////////////////////////////////////////////

                                    let Horário_Início_Criação_Campanha_RL = new Date(new Date().setMinutes(0, 0, 0) + 3 * 24 * 60 * 60 * 1000);
                                    let Horário_Término_Criação_Campanha_RL = new Date(Horário_Início_Criação_Campanha_RL.getTime() + 60 * 60 * 1000);

                                    if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

                                    await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/calendar/events').post({
                                        
                                        subject: "CAMPANHA RL - " + Reel_Código,
                                        
                                        start: {
                                            "dateTime": Horário_Início_Criação_Campanha_RL,
                                            "timeZone": "UTC"
                                        },
                                        
                                        end: {
                                            "dateTime": Horário_Término_Criação_Campanha_RL,
                                            "timeZone": "UTC"
                                        }
                                        
                                    })

                                    .then(async () => {

                                        if (client) client.send(JSON.stringify({ message: `Reels - 8. Criação da campanha agendada.`, origin: "postar" }));

                                        ////////////////////////////////////////////////////////////////////////////////////////
                                        ////////////////////////////////////////////////////////////////////////////////////////
                                        // Stories: Início do processo de postagem.
                                        ////////////////////////////////////////////////////////////////////////////////////////
                                        ////////////////////////////////////////////////////////////////////////////////////////

                                        if (Incluir_Stories === true) {

                                            ////////////////////////////////////////////////////////////////////////////////////////
                                            // Cria o Media Container (Stories).
                                            
                                            fetch(`https://graph.facebook.com/${Meta_Graph_API_Latest_Version}/${Meta_Graph_API_Instagram_Business_Account_ID}/media`, {
                                                
                                                method: 'POST',
                                                headers: {'Content-Type': 'application/json'},
                                                
                                                body: JSON.stringify({
                                                    media_type: 'STORIES',
                                                    video_url: Reel_Video_URL,
                                                    access_token: Meta_Graph_API_Access_Token
                                                })
                                            
                                            })

                                            .then(response => response.json()).then(data => {

                                                let Stories_IG_Media_Container_ID = data.id;

                                                if (Stories_IG_Media_Container_ID !== null && client) client.send(JSON.stringify({ message: `Stories - 1. Media Container ID ${Stories_IG_Media_Container_ID} criado.`, origin: "postar" }));

                                                ////////////////////////////////////////////////////////////////////////////////////////
                                                // Verifica o status do Media Container (Stories) a cada 5s.

                                                const VerificaStatusStories = () => {
                                                
                                                    fetch(`https://graph.facebook.com/${Meta_Graph_API_Latest_Version}/${Stories_IG_Media_Container_ID}?fields=status_code`, {
                                                        method: 'GET',
                                                        headers: {
                                                            'Content-Type': 'application/json',
                                                            'Authorization': `Bearer ${Meta_Graph_API_Access_Token}`
                                                        }
                                                    })
                                        
                                                    .then(response => response.json()).then(data => {
                                        
                                                        let Stories_IG_Media_Container_Status = data.status_code;
                                                        
                                                        if (Stories_IG_Media_Container_Status === "FINISHED") {
                                        
                                                            clearInterval(Stories_IG_Media_Container_Status_Verificação_ID);

                                                            if (client) client.send(JSON.stringify({ message: `Stories - 2. Media Container Status atualizado: ${Stories_IG_Media_Container_Status}.`, origin: "postar" }));

                                                            ////////////////////////////////////////////////////////////////////////////////////////
                                                            // Publica o Media Container (Stories).
                                                            
                                                            fetch(`https://graph.facebook.com/${Meta_Graph_API_Latest_Version}/${Meta_Graph_API_Instagram_Business_Account_ID}/media_publish`, {
                                                                
                                                                method: 'POST',
                                                                headers: {'Content-Type': 'application/json'},
                                                                
                                                                body: JSON.stringify({
                                                                    creation_id: Stories_IG_Media_Container_ID,
                                                                    access_token: Meta_Graph_API_Access_Token
                                                                })
                                                            
                                                            })
                                                            
                                                            .then(response => response.json()).then(data => {
                                                                
                                                                let Stories_IG_Media_ID = data.id;

                                                                if (Stories_IG_Media_ID !== null && client) client.send(JSON.stringify({ message: `Stories - 3. Stories ID ${Stories_IG_Media_ID} publicado.`, origin: "postar" }));
                                                                
                                                                if (Stories_IG_Media_ID !== null && client) client.send(JSON.stringify({ message: `--- fim ---`, origin: "postar" }));

                                                            })
                                                        
                                                        } else {

                                                            if (client) client.send(JSON.stringify({ message: `Stories - 2. Media Container Status atualizado: ${Stories_IG_Media_Container_Status}.`, origin: "postar" }));
                                        
                                                        }

                                                    });

                                                }

                                                const Stories_IG_Media_Container_Status_Verificação_ID = setInterval(VerificaStatusStories, 5000);

                                            })

                                        }
                                        
                                    });

                                });

                            });

                        });

                    });
                
                } else {

                    if (client) client.send(JSON.stringify({ message: `Reels - 3. Media Container Status atualizado: ${Reel_IG_Media_Container_Status}.`, origin: "postar" }));

                }

            })
            
        };

        const Reel_IG_Media_Container_Status_Verificação_ID = setInterval(VerificaStatusReels, 5000);
        
    })
    
});

////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////
// Endpoint: Cria Campanha de RL
////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////

app.post('/meta/CriaCampanhaRL', async (req, res) => {

    let { Reel_Código } = req.body;

    res.status(200).json({ message: "1. Request recebida." });

    ///////////////////////////////////////////////////////////////////////////////////////
    // Obtém os dados do Reel.
    ///////////////////////////////////////////////////////////////////////////////////////

    if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();
    
    let BD_Resultados_RL = await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/drive/items/0172BBJB5JTOTCSWCLGBB2HKLEFJVR7AUC/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{122865F8-2E2D-4B60-A34C-E02E001E835E}/rows').get();

    ////////////////////////////////////////////////////////////////////////////////////////////////
    // Obtém o Reel_IG_Media_ID e obtém o Reel_Audience_ID.

    let Index_Verificado = 0;
    let Reel_Código_Verificado;
    let Reel_IG_Media_ID;
    let Reel_Audience_ID;

    Obtém_Dados_Reel();

    function Obtém_Dados_Reel() {

        Reel_Código_Verificado = BD_Resultados_RL.value[Index_Verificado].values[0][0];

        if (Reel_Código_Verificado === Reel_Código) {

            Reel_IG_Media_ID =  BD_Resultados_RL.value[Index_Verificado].values[0][1];
            Reel_Audience_ID = BD_Resultados_RL.value[Index_Verificado].values[0][2];


            if(client) client.send(JSON.stringify({ message: `2. Reel_IG_Media_ID e Reel_Audience_ID encontrados.`, origin: "CriaCampanhaRL" }));
        
        } else {

            Index_Verificado++;
            Obtém_Dados_Reel();

        }
    
    }  

    ///////////////////////////////////////////////////////////////////////////////////////
    // Cria a Campanha.
    ///////////////////////////////////////////////////////////////////////////////////////

    fetch(`https://graph.facebook.com/${Meta_Graph_API_Latest_Version}/${Meta_Graph_API_Ad_Account_ID}/campaigns`, {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({
            name: 'RL.' + Reel_Código,
            status: 'ACTIVE',
            buying_type: 'AUCTION',
            objective: 'OUTCOME_ENGAGEMENT',
            special_ad_categories: [],
            campaign_budget_optimization: false,
            is_ab_test: false,
            access_token: Meta_Graph_API_Access_Token
        })
    })

    .then(response => response.json()).then(async data => {

        let Reel_Campanha_Relacionamento_ID = data.id;

        if(Reel_Campanha_Relacionamento_ID !== null && client) client.send(JSON.stringify({ message: `3. Campanha de RL criada.`, origin: "CriaCampanhaRL" }));

        ///////////////////////////////////////////////////////////////////////////////////////
        // Obtém o orçamento mínimo diário atualizado, em BRL,
        // para setar o Orçamento do Conjunto de Anúncios.
        ///////////////////////////////////////////////////////////////////////////////////////

        fetch(`https://graph.facebook.com/${Meta_Graph_API_Latest_Version}/${Meta_Graph_API_Ad_Account_ID}?fields=min_daily_budget`, {
            method: 'GET',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${Meta_Graph_API_Access_Token}`
            }
        })

        .then(response => response.json()).then(data => {

            let Reel_Conjunto_Anuncios_Orçamento = data.min_daily_budget;

            if(Reel_Conjunto_Anuncios_Orçamento !== null && client) client.send(JSON.stringify({ message: `4. Orçamento do Conjunto de Anúncios obtido.`, origin: "CriaCampanhaRL" }));

            ///////////////////////////////////////////////////////////////////////////////////////
            // Cria o Conjunto de Anúncios.
            ///////////////////////////////////////////////////////////////////////////////////////

            fetch(`https://graph.facebook.com/${Meta_Graph_API_Latest_Version}/${Meta_Graph_API_Ad_Account_ID}/adsets`, {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify({
                    status: 'ACTIVE',
                    campaign_id: Reel_Campanha_Relacionamento_ID,
                    name: 'RL.' + Reel_Código,
                    destination_type: 'ON_VIDEO',
                    optimization_goal: 'THRUPLAY',
                    frequency_control_specs: [{
                        event:'IMPRESSIONS',
                        interval_days: 90,
                        max_frequency: 1
                    }],
                    daily_budget: Reel_Conjunto_Anuncios_Orçamento,
                    billing_event: 'IMPRESSIONS',
                    bid_strategy: 'LOWEST_COST_WITHOUT_CAP',
                    start_time: Math.floor(Date.now() / 1000),
                    targeting:{
                        "geo_locations": {
                            "countries":["BR"]
                        },
                        "age_min":18,
                        "age_max":65,
                        "custom_audiences": [{"id": Meta_Graph_API_Custom_Audience_ID_Seguidores}],
                        "excluded_custom_audiences": [{"id": Reel_Audience_ID}],
                        "targeting_relaxation_types": {
                            "lookalike": 0,
                            "custom_audience": 0
                        },
                        "publisher_platforms": ["instagram"],
                        "instagram_positions": ["stream","profile_reels","explore","reels", "explore_home", "profile_feed"]
                    },
                    access_token: Meta_Graph_API_Access_Token
                })
            })

            .then(response => response.json()).then(async data => {

                let Reel_Conjunto_Anuncios_Relacionamento_ID = data.id;

                if (Reel_Conjunto_Anuncios_Relacionamento_ID !== null && client) client.send(JSON.stringify({ message: `5. Conjunto de Anúncios criado.`, origin: "CriaCampanhaRL" }));

                ///////////////////////////////////////////////////////////////////////////////////////
                // Cria o Criativo.
                ///////////////////////////////////////////////////////////////////////////////////////

                fetch(`https://graph.facebook.com/${Meta_Graph_API_Latest_Version}/${Meta_Graph_API_Ad_Account_ID}/adcreatives`, {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({
                        object_id: Meta_Graph_API_Facebook_Page_ID,
                        instagram_user_id: Meta_Graph_API_Instagram_Business_Account_ID,
                        source_instagram_media_id: Reel_IG_Media_ID,
                        call_to_action: {
                            type: 'LEARN_MORE',
                            value: {
                                link: 'https://ivygestao.com/'
                            }
                        },
                        contextual_multi_ads: {
                            enroll_status: 'OPT_OUT'
                        },
                        access_token: Meta_Graph_API_Access_Token
                    })
                })

                .then(response => response.json()).then(async data => {

                    let Reel_Criativo_Relacionamento_ID = data.id;

                    console.log(response);

                    console.log("-----------");

                    console.log(data);

                    if (Reel_Criativo_Relacionamento_ID !== null && client) client.send(JSON.stringify({ message: `6. Criativo criado. ${Reel_Criativo_Relacionamento_ID}`, origin: "CriaCampanhaRL" }));

                    ///////////////////////////////////////////////////////////////////////////////////////
                    // Cria o Anúncio (após 2s, para garantir a existência / carregamento do Criativo).
                    ///////////////////////////////////////////////////////////////////////////////////////
                        
                    fetch(`https://graph.facebook.com/${Meta_Graph_API_Latest_Version}/${Meta_Graph_API_Ad_Account_ID}/ads`, {
                        method: 'POST',
                        headers: {'Content-Type': 'application/json'},
                        body: JSON.stringify({
                            name: 'RL.' + Reel_Código,
                            adset_id: Reel_Conjunto_Anuncios_Relacionamento_ID,
                            status: 'ACTIVE',
                            creative: {creative_id: Reel_Criativo_Relacionamento_ID},
                            access_token: Meta_Graph_API_Access_Token
                        })
                    })

                    .then(response => response.json()).then(async data => {

                        let Reel_Anúncio_Relacionamento_ID = data.id;

                        if (Reel_Anúncio_Relacionamento_ID !== null && client) client.send(JSON.stringify({ message: `7. Anúncio criado. ${Reel_Anúncio_Relacionamento_ID}`, origin: "CriaCampanhaRL" }));

                        if (Reel_Anúncio_Relacionamento_ID !== null && client) client.send(JSON.stringify({ message: '--- Fim ---', origin: "CriaCampanhaRL" }));

                    });
                    
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
    // Puxa os dados da BD - RESULTADOS.
    
    if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

    let BD_Resultados_RL = await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/drive/items/0172BBJB5JTOTCSWCLGBB2HKLEFJVR7AUC/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{122865F8-2E2D-4B60-A34C-E02E001E835E}/rows').get();

    ///////////////////////////////////////////////////////////////////////////////////////////////////////////
    // Verifica se há criativos com:
    // --> O número de CONTAS ALCANÇADAS (5%) e INTERAÇÕES (5%) em branco.
    // --> O número de CONTAS ALCANÇADAS (72h) e INTERAÇÕES (72h) em branco.

    let BD_Resultados_RL_Número_Linhas = BD_Resultados_RL.value.length;
    let BD_Resultados_RL_Última_Linha = BD_Resultados_RL_Número_Linhas - 1;
    
    for (let LinhaVerificada = 0; LinhaVerificada <= BD_Resultados_RL_Última_Linha; LinhaVerificada++) {

        let Reel_IG_Media_ID = BD_Resultados_RL.value[LinhaVerificada].values[0][1];
        let Reel_Data_e_Hora_Postagem = ConverteData4(BD_Resultados_RL.value[LinhaVerificada].values[0][3]);
        let Reel_Número_de_Seguidores_Momento_Postagem = BD_Resultados_RL.value[LinhaVerificada].values[0][4];
        let Reel_Contas_Alcançadas_5Porcento_Registrado = BD_Resultados_RL.value[LinhaVerificada].values[0][5];
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

                    await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/drive/items/0172BBJB5JTOTCSWCLGBB2HKLEFJVR7AUC/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{122865F8-2E2D-4B60-A34C-E02E001E835E}/rows/itemAt(index=' + LinhaVerificada + ')').update({values: [[null, null, null, null, null, Reel_Organic_Reach_Atual, Reel_Organic_Interactions_Atual, null, null ]]})

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

                    await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/drive/items/0172BBJB5JTOTCSWCLGBB2HKLEFJVR7AUC/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{122865F8-2E2D-4B60-A34C-E02E001E835E}/rows/itemAt(index=' + LinhaVerificada + ')').update({values: [[null, null, null, null, null, Reel_Organic_Reach_Atual, Reel_Organic_Interactions_Atual, Reel_Organic_Reach_Atual, Reel_Organic_Interactions_Atual ]]});
                    
                } 
                
                ///////////////////////////////////////////////////////////////////////////////////////////////////////////
                // Caso número de CONTAS ALCANÇADAS (5%) e INTERAÇÕES (5%) não esteja em branco.
                // --> Registra as informações do criativo na BD - RESULTADOS:
                //       - Somente nas colunas CONTAS ALCANÇADAS (72h) e INTERAÇÕES (72h).
                
                else {

                    if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();
                    
                    await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/drive/items/0172BBJB5JTOTCSWCLGBB2HKLEFJVR7AUC/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{122865F8-2E2D-4B60-A34C-E02E001E835E}/rows/itemAt(index=' + LinhaVerificada + ')').update({values: [[null, null, null, null, null, null, null, Reel_Organic_Reach_Atual, Reel_Organic_Interactions_Atual ]]});
                    
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
    // Obtém a BD - STATUS CAMPANHAS.
    ///////////////////////////////////////////////////////////////////////////////////////

    if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();
    
    let BD_Status_Campanhas_DB = await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/drive/items/0172BBJBYTOUDIQ5V5KBEIMADJCDNO2S4Z/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{93C2A633-D78C-42B0-9A68-937848657884}/rows').get();

    const BD_Status_Campanhas_DB_Última_Linha = BD_Status_Campanhas_DB.value.length - 1;

    ///////////////////////////////////////////////////////////////////////////////////////
    // Registra um de desempenho de Campanha de DB a cada 2s.
    ///////////////////////////////////////////////////////////////////////////////////////

    for (let LinhaAtual = 0; LinhaAtual <= BD_Status_Campanhas_DB_Última_Linha; LinhaAtual++) {

        ///////////////////////////////////////////////////////////////////////////////////////////////////
        // Obtém as variáveis do aluno da BD - OFFICE HOURS.

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
                    // Adiciona as informações à BD - RESULTADOS CAMPANHAS.

                    if (!Microsoft_Graph_API_Client) await Conecta_ao_Microsoft_Graph_API();

                    await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/drive/items/0172BBJBY6STB7R6BSQBFKAY5LO6W3TFRR/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{93C2A633-D78C-42B0-9A68-937848657884}/rows')
                    
                    .post({"values": [[ "-", ConverteData2(new Date(new Date().setDate(new Date().getDate() - 1))), Campanha_DB_Reel_Código, `'${Campanha_DB_Ad_ID}`, Campanha_DB_Descrição, Campanha_DB_Qualidade_Clique, Campanha_DB_Ad_Spend, Campanha_DB_Ad_Reach, Campanha_DB_Ad_Impressions, Campanha_DB_Ad_Link_Clicks, "-", "-", `=${Campanha_DB_Campaign_Daily_Budget}/100`, "-", "-", "-", "-", "-", "-", "-" ]]})

                });

            });

            await new Promise(resolve => setTimeout(resolve, 2000));

        } else {

            await new Promise(resolve => setTimeout(resolve, 0));

        }

    }

});