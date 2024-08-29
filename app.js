// ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// // IMPORTA TODAS AS BIBLIOTECAS NECESSÁRIAS.
// ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// Importa as bibliotecas para comunicação via HTTP Posts (incluindo pelo Live Server) e cria o endpoint no servidor.
const express = require('express');
const cors = require('cors');
const app = express();
const port = process.env.PORT || 3000;
app.use(cors());
app.use(express.json());
app.listen(port);

// Importa as bibliotecas de comunicação com o Microsoft Graph API.
const { Client } = require('@microsoft/microsoft-graph-client');
const { ConfidentialClientApplication } = require('@azure/msal-node');
var accessToken;
var Microsoft_Graph_API_Client;

async function Conecta_ao_Microsoft_Graph_API() {
    const cca = new ConfidentialClientApplication({ auth: { clientId: process.env.CLIENT_ID, authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`, clientSecret: process.env.CLIENT_SECRET } });
    accessToken = (await cca.acquireTokenByClientCredential({scopes: ['https://graph.microsoft.com/.default']})).accessToken;
    Microsoft_Graph_API_Client = Client.init({authProvider:(done)=>{done(null, accessToken)}});
    console.log("#. AccessToken do Microsoft Graph API renovado.");
}


// ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// // PROCESSA O LOGIN DO USUÁRIO
// ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

app.post('/login', async (req, res) => {
    
    // //Retorna aviso de sucesso no recebimento da mensagem à request.
    // res.status(200).send('EmailUsuário e SenhaUsuário recebidos pelo backend.');

    // // Obtém os dados da BD - PLATAFORMA.xlsx.
    // if (!Microsoft_Graph_API_Client) { await Conecta_ao_Microsoft_Graph_API() };
    // const BD_Plataforma = await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/drive/items/0172BBJB7TUZJNIWDVWFE2MIW7MNKHMWLL/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows').get();    

    // console.log(BD_Plataforma);

    console.log(req.body);

});







// ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// // PROCESSA MENSAGENS DE ENTRADA
// ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// // Importa a biblioteca para comunicação com variáveis de ambiente.
// const dotenv = require('dotenv');
// dotenv.config();

// // Importa a biblioteca para comunicação HTTP Posts e cria o endpoint no servidor.
// const express = require('express');
// const app = express();
// const port = process.env.PORT || 3000;
// app.listen(port);
// app.use(express.json());

// // Importa as bibliotecas de comunicação com o Microsoft Graph API e renova o acesso (AccessToken e Client) a cada 30min.
// const { Client } = require('@microsoft/microsoft-graph-client');
// const { ConfidentialClientApplication } = require('@azure/msal-node');
// var accessToken;
// var Microsoft_Graph_API_Client;

// async function Conecta_ao_Microsoft_Graph_API() {
//     const cca = new ConfidentialClientApplication({ auth: { clientId: process.env.CLIENT_ID, authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`, clientSecret: process.env.CLIENT_SECRET } });
//     accessToken = (await cca.acquireTokenByClientCredential({scopes: ['https://graph.microsoft.com/.default']})).accessToken;
//     Microsoft_Graph_API_Client = Client.init({authProvider:(done)=>{done(null, accessToken)}});
//     console.log("#. AccessToken do Microsoft Graph API renovado.");
// }

// // Função auxiliar para formatação de DataHora como "DD/MMM/AAAA HH:mm" em BRT. 
// function FormataDataHora(datahora_a_formatar) {
//   const day = ('0' + datahora_a_formatar.getDate()).slice(-2);
//   const month = datahora_a_formatar.toLocaleString('pt-BR', { month: 'short' }).replace('.', '');
//   const year = datahora_a_formatar.getFullYear();
//   const hours = ('0' + datahora_a_formatar.getHours()).slice(-2);
//   const minutes = ('0' + datahora_a_formatar.getMinutes()).slice(-2);
//   return `${day}/${month}/${year} ${hours}:${minutes}`;
// }

// // Cria a Fila que controla o processamento das Mensagens de Entrada vindas do ManyChat.
// let Fila_Processamento_Mensagens_Entrada = [];

// //////////////////////////////////////////////////////////////////////////////////
// //////////////////////////////////////////////////////////////////////////////////
// // Recebe Mensagem de Entrada vinda do ManyChat.
// //////////////////////////////////////////////////////////////////////////////////
// //////////////////////////////////////////////////////////////////////////////////

// app.post('/manychat-webhook/:VariaveisConsolidadas', async (req, res) => {  
 
//     // Retorna aviso de sucesso no recebimento da mensagem à request.
//     res.status(200).send('Mensagem de Entrada recebida pelo app.js com sucesso.');
    
//     const MensagemEntrada_VariaveisConsolidadas = req.params.VariaveisConsolidadas;
//     const MensagemEntrada_Mensagem = req.body.Mensagem;

//     //////////////////////////////////////////////////////////////////////////////////
//     //////////////////////////////////////////////////////////////////////////////////
//     // Adiciona a Mensagem de Entrada à fila de processamento.
//     //////////////////////////////////////////////////////////////////////////////////
//     //////////////////////////////////////////////////////////////////////////////////

//     Fila_Processamento_Mensagens_Entrada.push({ MensagemEntrada_VariaveisConsolidadas, MensagemEntrada_Mensagem });

//     console.log(`ME.1. Mensagem de Entrada adicionada à fila de processamento.`);

//     //////////////////////////////////////////////////////////////////////////////////
//     //////////////////////////////////////////////////////////////////////////////////
//     // Há outras Mensagens de Entrada na fila de processamento?
//     //////////////////////////////////////////////////////////////////////////////////
//     //////////////////////////////////////////////////////////////////////////////////

//     if (Fila_Processamento_Mensagens_Entrada.length === 1) {
        
//         // Não:

//         //////////////////////////////////////////////////////////////////////////////////
//         //////////////////////////////////////////////////////////////////////////////////
//         // Processa a nova Mensagem de Entrada.
//         //////////////////////////////////////////////////////////////////////////////////
//         //////////////////////////////////////////////////////////////////////////////////

//         await ProcessaMensagemEntrada(Fila_Processamento_Mensagens_Entrada[0]);
    
//     } else {
        
//         // Sim:

//         console.log(`ME.2.B. Mensagem de Entrada aguardando para ser processada.`);
    
//     }

// });

// async function ProcessaMensagemEntrada(Mensagem_a_Processar) {

//     // Define as variáveis extraídas da DM.
//     const { MensagemEntrada_VariaveisConsolidadas, MensagemEntrada_Mensagem } = Mensagem_a_Processar;
//     const MensagemEntrada_ManyChatSubscriberID = parseInt(MensagemEntrada_VariaveisConsolidadas.split("@")[0]);
//     const MensagemEntrada_Perfil = MensagemEntrada_VariaveisConsolidadas.split("@")[1].split("*")[0];
//     const MensagemEntrada_NomeCompleto = MensagemEntrada_VariaveisConsolidadas.split("@")[1].split("*")[1].replace(/'/g, '*');
//     const MensagemEntrada_DataHora_Formatada = FormataDataHora(new Date(Date.now() - 3 * 60 * 60 * 1000));

//     console.log(`ME.2.A. Processamento da Mensagem de Entrada iniciado.`);
    
//     //////////////////////////////////////////////////////////////////////////////////
//     //////////////////////////////////////////////////////////////////////////////////
//     // Conversa é nova?
//     //////////////////////////////////////////////////////////////////////////////////
//     //////////////////////////////////////////////////////////////////////////////////

//     // Obtém os dados da BD - POTENCIAIS ALUNOS.xlsx.

//     if (!Microsoft_Graph_API_Client) { await Conecta_ao_Microsoft_Graph_API() };

//     const BD_Potenciais_Alunos = await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/drive/items/0172BBJB3BYUNDDIMYQJF37AQWCYSADSTT/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{AC8C07F3-9A79-4ABD-8CE8-0C818B0EA1A7}/rows').get();    

//     // Verifica se a conversa é nova.

//     const BD_Potenciais_Alunos_Número_Linhas = BD_Potenciais_Alunos.value.length;

//     var ConversaNova = "Sim";
//     var Index_LinhaVerificada = 0;
//     var ManyChatSubscriberIDVerificado; 

//     VerificaManyChatSubscriberIDs();

//     function VerificaManyChatSubscriberIDs() {

//         if (Index_LinhaVerificada < BD_Potenciais_Alunos_Número_Linhas) {
        
//             ManyChatSubscriberIDVerificado = BD_Potenciais_Alunos.value[Index_LinhaVerificada].values[0][1];
        
//             if (ManyChatSubscriberIDVerificado === MensagemEntrada_ManyChatSubscriberID){

//                 ConversaNova = "Não";

//             } else {

//                 Index_LinhaVerificada++;
        
//                 VerificaManyChatSubscriberIDs();

//             }
    
//         }

//     }

//     // Conversa é nova? Sim:

//     if (ConversaNova === "Sim") {

//         //////////////////////////////////////////////////////////////////////////////////
//         //////////////////////////////////////////////////////////////////////////////////
//         // Adiciona Potencial Aluno na BD - POTENCIAIS ALUNOS.
//         //////////////////////////////////////////////////////////////////////////////////
//         //////////////////////////////////////////////////////////////////////////////////

//         if (!Microsoft_Graph_API_Client) { await Conecta_ao_Microsoft_Graph_API() };

//         await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/drive/items/0172BBJB3BYUNDDIMYQJF37AQWCYSADSTT/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{AC8C07F3-9A79-4ABD-8CE8-0C818B0EA1A7}/rows/add')

//             .post({ index: null, values:[[MensagemEntrada_DataHora_Formatada, MensagemEntrada_ManyChatSubscriberID, MensagemEntrada_Perfil, MensagemEntrada_NomeCompleto, '-', '-', '1', null, null, null]]});

//         console.log('ME.3.A.1 Potencial Aluno adicionado à BD - POTENCIAIS ALUNOS.');

//         //////////////////////////////////////////////////////////////////////////////////
//         //////////////////////////////////////////////////////////////////////////////////
//         // Cria a BD - MENSAGENS do PA.
//         //////////////////////////////////////////////////////////////////////////////////
//         //////////////////////////////////////////////////////////////////////////////////

//         // Cria o arquivo.

//         if (!Microsoft_Graph_API_Client) { await Conecta_ao_Microsoft_Graph_API() };

//         await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/drive/items/0172BBJB2ABWOTQRVAWRDKN2B3U5544ABK/copy')

//             .post({"parentReference": { "id": "0172BBJB7JFBO2TLUKOZCYCXO7UTPPUZ5W" },"name": MensagemEntrada_ManyChatSubscriberID + ".xlsx"});

//         console.log('ME.3.A.2 BD - MENSAGENS (' + MensagemEntrada_ManyChatSubscriberID + '.xlsx) do Potencial Aluno criada.');

//     }

//     // Conversa é nova? Não:

//     else if (ConversaNova === "Não") {

//         console.log(`ME.3.B. Potencial Aluno já estava listado na BD - POTENCIAIS ALUNOS.`);

//     }

//     //////////////////////////////////////////////////////////////////////////////////
//     //////////////////////////////////////////////////////////////////////////////////
//     // Adiciona mensagem nas duas tabelas da BD - MENSAGENS do PA.
//     //////////////////////////////////////////////////////////////////////////////////
//     //////////////////////////////////////////////////////////////////////////////////

//     // Obtém os IDs necessários.

//     if (!Microsoft_Graph_API_Client) { await Conecta_ao_Microsoft_Graph_API() };
//     var ID_Arquivo = (await Microsoft_Graph_API_Client.api("/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/drive/items/0172BBJB7JFBO2TLUKOZCYCXO7UTPPUZ5W/children?$filter=name eq '" + MensagemEntrada_ManyChatSubscriberID + ".xlsx'").get()).value[0].id;
    
//     if (!Microsoft_Graph_API_Client) { await Conecta_ao_Microsoft_Graph_API() };
//     var ID_Aba_BD_GPT = (await Microsoft_Graph_API_Client.api("/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/drive/items/" + ID_Arquivo + "/workbook/worksheets").get()).value[0].id;
    
//     if (!Microsoft_Graph_API_Client) { await Conecta_ao_Microsoft_Graph_API() };
//     var ID_Aba_BD_Padrão = (await Microsoft_Graph_API_Client.api("/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/drive/items/" + ID_Arquivo + "/workbook/worksheets").get()).value[1].id;
    
//     if (!Microsoft_Graph_API_Client) { await Conecta_ao_Microsoft_Graph_API() };
//     var ID_Tabela_BD_GPT = (await Microsoft_Graph_API_Client.api("/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/drive/items/" + ID_Arquivo + "/workbook/worksheets/" + ID_Aba_BD_GPT + "/tables").get()).value[0].id;
    
//     if (!Microsoft_Graph_API_Client) { await Conecta_ao_Microsoft_Graph_API() };
//     var ID_Tabela_BD_Padrão = (await Microsoft_Graph_API_Client.api("/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/drive/items/" + ID_Arquivo + "/workbook/worksheets/" + ID_Aba_BD_Padrão + "/tables").get()).value[0].id;

//     // Adiciona mensagem à BD_GPT.

//     if (!Microsoft_Graph_API_Client) { await Conecta_ao_Microsoft_Graph_API() };

//     await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/drive/items/' + ID_Arquivo + '/workbook/worksheets/' + ID_Aba_BD_GPT + '/tables/' + ID_Tabela_BD_GPT + '/rows/add')

//         .post({ index: null, values:[['-', 'user', MensagemEntrada_Mensagem]]});

//     // Adiciona mensagem à BD_Padrão.

//     if (!Microsoft_Graph_API_Client) { await Conecta_ao_Microsoft_Graph_API() };

//     await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/drive/items/' + ID_Arquivo + '/workbook/worksheets/' + ID_Aba_BD_Padrão + '/tables/' + ID_Tabela_BD_Padrão + '/rows/add')

//         .post({ index: null, values:[[MensagemEntrada_DataHora_Formatada, MensagemEntrada_Mensagem]]});

//     console.log('ME.4 Mensagem adicionada às abas BD_GPT e BD_Padrão da BD - MENSAGENS (' + MensagemEntrada_ManyChatSubscriberID + '.xlsx) do Potencial Aluno.');


//     //////////////////////////////////////////////////////////////////////////////////
//     //////////////////////////////////////////////////////////////////////////////////
//     // Atualiza a BD - FIFO.
//     //////////////////////////////////////////////////////////////////////////////////
//     //////////////////////////////////////////////////////////////////////////////////

//     //////////////////////////////////////////////////////////////////////////////////
//     // PA já listado no FIFO?

//     // Obtém os dados da BD - FIFO.xlsx.

//     if (!Microsoft_Graph_API_Client) { await Conecta_ao_Microsoft_Graph_API() };

//     const BD_FIFO = await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/drive/items/0172BBJB37QFH4WTOM5RC3JNYHC4IY4PVB/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{4877C290-D071-46E4-A784-BC2793A37396}/rows').get();    

//     // Verifica se a conversa é nova.

//     const BD_FIFO_Número_Linhas = BD_FIFO.value.length;

//     var PA_Já_Listado = "Não";
//     var Index_LinhaVerificada_BD_FIFO = 0;
//     var PerfilVerificado_BD_FIFO; 

//     VerificaManyChatSubscriberIDs_BD_FIFO();

//     function VerificaManyChatSubscriberIDs_BD_FIFO() {

//         if (Index_LinhaVerificada_BD_FIFO < BD_FIFO_Número_Linhas) {
        
//             PerfilVerificado_BD_FIFO = BD_FIFO.value[Index_LinhaVerificada_BD_FIFO].values[0][0];
        
//             if (PerfilVerificado_BD_FIFO === MensagemEntrada_Perfil){

//                 PA_Já_Listado = "Sim";

//             } else {

//                 Index_LinhaVerificada_BD_FIFO++;
        
//                 VerificaManyChatSubscriberIDs_BD_FIFO();

//             }
    
//         }

//     }

//     //////////////////////////////////////////////////////////////////////////////////
//     // PA já listado no FIFO? Não.

//     if (PA_Já_Listado === "Não") {

//         //////////////////////////////////////////////////////////////////////////////////
//         // Adiciona PA ao FIFO.

//         if (!Microsoft_Graph_API_Client) { await Conecta_ao_Microsoft_Graph_API() };

//         await Microsoft_Graph_API_Client.api('/users/b4a93dcf-5946-4cb2-8368-5db4d242a236/drive/items/0172BBJB37QFH4WTOM5RC3JNYHC4IY4PVB/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{4877C290-D071-46E4-A784-BC2793A37396}/rows/add')

//             .post({ index: null, values:[[MensagemEntrada_Perfil]]});

//         console.log('ME.5.A PA adicionado ao FIFO.')

//     }

//     //////////////////////////////////////////////////////////////////////////////////
//     // PA já listado no FIFO? Sim.

//     if (PA_Já_Listado === "Sim") {

//         console.log('ME.5.B PA já estava listado no FIFO.')

//     }

//     //////////////////////////////////////////////////////////////////////////////////
//     //////////////////////////////////////////////////////////////////////////////////
//     // Retira a Mensagem de Entrada já processada da fila de processamento.
//     //////////////////////////////////////////////////////////////////////////////////
//     //////////////////////////////////////////////////////////////////////////////////

//     Fila_Processamento_Mensagens_Entrada.shift();

//     console.log(`ME.6. Mensagem de Entrada retirada da fila de processamento.`);

//     //////////////////////////////////////////////////////////////////////////////////
//     //////////////////////////////////////////////////////////////////////////////////
//     // Processa a próxima Mensagem de Entrada na fila de processamento.
//     //////////////////////////////////////////////////////////////////////////////////
//     //////////////////////////////////////////////////////////////////////////////////

//     if (Fila_Processamento_Mensagens_Entrada.length > 0) {

//         await ProcessaMensagemEntrada(Fila_Processamento_Mensagens_Entrada[0]);
    
//     }

// };

// ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// // PROCESSA MENSAGENS DE SAÍDA
// ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// // Importa a biblioteca necessária para fazer os HTTP POST ao ManyChat e configura a conexão.
// const axios = require('axios');
// const url = 'https://api.manychat.com/fb/subscriber/setCustomField';
// const headers = {
//   'accept': 'application/json',
//   'Authorization': 'Bearer 881250:76d2acb1dcd7ddf429f70d584df3a07c',
//   'Content-Type': 'application/json',
// };

// app.post('/excel-webhook/:PA_ManyChatSubscriberID', async (req, res) => {

//     // Retorna mensagem de sucesso à request.
//     res.status(200).send('Mensagens de Saída recebidas pelo app.js com sucesso.');

//     const NúmeroMensagensSaída = 5;
//     const MensagensSaída = [];

//     // Define as variáveis extraídas da Mensagem de Saída.
//     const MensagemSaída_ManyChatSubscriberID = req.params.PA_ManyChatSubscriberID;
//     const MensagensSaída_JSON_Body = req.body;
//     MensagensSaída[0] = req.body.mensagem_1;
//     MensagensSaída[1] = req.body.mensagem_2;
//     MensagensSaída[2] = req.body.mensagem_3;
//     MensagensSaída[3] = req.body.mensagem_4;
//     MensagensSaída[4] = req.body.mensagem_5;

//     console.log(`MS.1. JSON das Mensagens de Saída recebido com sucesso:`);
//     console.log(MensagensSaída_JSON_Body);

//     //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//     //Função que envia cada Mensagem de Saída ao ManyChat, a cada 35s.
    
//     let NúmeroMensagemSaídaProcessada = 0;

//     Processa_Próxima_Mensagem_Saída();

//     function Processa_Próxima_Mensagem_Saída() {

//         const data = {
//             subscriber_id: MensagemSaída_ManyChatSubscriberID,
//             field_id: 10238769,
//             field_value: MensagensSaída[NúmeroMensagemSaídaProcessada],
//         };

//         axios.post(url, data, { headers })

//         .then(response => {
            
//             console.log(`MS.2. Mensagem de Saída enviada ao ManyChat. Status: ${response.status}.`);

//             if (MensagensSaída[NúmeroMensagemSaídaProcessada + 1] === "" || NúmeroMensagemSaídaProcessada + 1 === NúmeroMensagensSaída){

//             console.log("MS.3. Todas as Mensagens de Saída foram enviadas ao ManyChat com sucesso.");
//             return;

//             } else {

//                 NúmeroMensagemSaídaProcessada++;
//                 setTimeout(Processa_Próxima_Mensagem_Saída, 35000);

//             }
        
//         })

//     }

// });












// ////////////////////////////////////////////////////////////////////////////////////////////////////////////














// var ContainerVídeoPrincipal = document.getElementById("Container-Vídeo-Principal");
// var BotãoTelaCheia1 = document.getElementById("Botão-Tela-Cheia-1");
// var TextoTelaCheia1 = document.getElementById("Texto-Tela-Cheia-1");

// var Seção1 = document.getElementById("Seção-1");
// var ContainerExternoSeção1 = document.getElementById("Container-Externo-Seção-1");
// var ContainerInternoSeção1 = document.getElementById("Container-Interno-Seção-1");
// var SubseçãoCadastro1 = document.getElementById("Subseção-Cadastro-1");
// var SetaFechamentoSeção1 = document.getElementById("Seta-Fechamento-Seção-1");

// var Seção2 = document.getElementById("Seção-2");
// var ContainerExternoSeção2 = document.getElementById("Container-Externo-Seção-2");
// var ContainerInternoSeção2 = document.getElementById("Container-Interno-Seção-2");
// var SubseçãoCadastro2 = document.getElementById("Subseção-Cadastro-2");
// var SetaFechamentoSeção2 = document.getElementById("Seta-Fechamento-Seção-2");

// var Seção3 = document.getElementById("Seção-3");
// var ContainerExternoSeção3 = document.getElementById("Container-Externo-Seção-3");
// var ContainerInternoSeção3 = document.getElementById("Container-Interno-Seção-3");
// var SubseçãoCadastro3 = document.getElementById("Subseção-Cadastro-3");
// var SetaFechamentoSeção3 = document.getElementById("Seta-Fechamento-Seção-3");

// var Seção4 = document.getElementById("Seção-4");
// var ContainerExternoSeção4 = document.getElementById("Container-Externo-Seção-4");
// var ContainerInternoSeção4 = document.getElementById("Container-Interno-Seção-4");
// var ContainerCondiçõesComerciais3 = document.getElementById("Container-Condições-Comerciais-3");
// var SetaFechamentoSeção4 = document.getElementById("Seta-Fechamento-Seção-4");
// var BotãoInstagramDirect = document.getElementById("Botão-Instagram-Direct");

// var ContainerVídeoDepoimentos = document.getElementById("Container-Vídeo-Depoimentos");
// var VídeoDepoimentos = document.getElementById("Vídeo-Depoimentos");
// var BotãoTelaCheia2 = document.getElementById("Botão-Tela-Cheia-2");
// var TextoTelaCheia2 = document.getElementById("Texto-Tela-Cheia-2");

// var BotãoCronograma = document.getElementById("Botão-Cronograma");
// var BotãoEmenta = document.getElementById("Botão-Ementa");
// var BotãoBibliografia = document.getElementById("Botão-Bibliografia");
// var BotãoSoftwares = document.getElementById("Botão-Softwares");

// var ContainerExternoSubseção41 = document.getElementById("Container-Externo-Subseção-4-1");
// var ContainerInternoSubseção41 = document.getElementById("Container-Interno-Subseção-4-1");
// var SetaFechamentoSubseção41 = document.getElementById("Seta-Fechamento-Subseção-4-1");

// var ContainerExternoSubseção42 = document.getElementById("Container-Externo-Subseção-4-2");
// var ContainerInternoSubseção42 = document.getElementById("Container-Interno-Subseção-4-2");
// var SetaFechamentoSubseção42 = document.getElementById("Seta-Fechamento-Subseção-4-2");

// var ContainerExternoSubseção43 = document.getElementById("Container-Externo-Subseção-4-3");
// var ContainerInternoSubseção43 = document.getElementById("Container-Interno-Subseção-4-3");
// var SetaFechamentoSubseção43 = document.getElementById("Seta-Fechamento-Subseção-4-3");

// var EspaçoFinalContainerBotãoCadastro = document.getElementById("Espaço-Final-Container-Botão-Cadastro");
// var ContainerBotãoCadastro = document.getElementById("Container-Botão-Cadastro");
// var BotãoCadastro = document.getElementById("Botão-Cadastro");

// var userAgent = navigator.userAgent;

// /*//////////////////////////////////////////////////////////////////////////////////////////////////////*/
// /*/// Retira os botões de girar os vídeos se o usuário não estiver usando o Instagram In-App Browser. //*/
// /*//////////////////////////////////////////////////////////////////////////////////////////////////////*/

// if (userAgent.indexOf('Instagram') === -1) {
    
//     BotãoTelaCheia1.style.display = 'none';
//     BotãoTelaCheia2.style.display = 'none';
//     ContainerVídeoPrincipal.style.marginBottom = '40px';
//     ContainerVídeoDepoimentos.style.marginBottom = '40px';

// }


// /*/////////////////////////////////////////////////////////////////////////////////////*/
// /*//// Carrega o Vídeo-Principal somente ao visualizar o Container-Vídeo-Principal ////*/
// /*/////////////////////////////////////////////////////////////////////////////////////*/

// function handleIntersection(entries, observer) {
    
//     entries.forEach(entry => {
      
//         if (entry.isIntersecting) {

//         var VídeoPrincipal = document.createElement("iframe");
//         VídeoPrincipal.setAttribute("id", "Vídeo-Principal");
//         VídeoPrincipal.setAttribute("frameborder", "0");
//         VídeoPrincipal.setAttribute("src", "https://www.youtube.com/embed/z9tmj6jxhrY?si=QwcMiypQLzI9bSVf");

//         /* Adiciona o allowfullscreen para os dois vídeos se o usuário estiver no Laptop ou num navegador que não seja o Instagram In-App Browser. */
//         if (window.innerWidth > 430 || userAgent.indexOf('Instagram') === -1) {
//             VídeoPrincipal.setAttribute("allowfullscreen", "true");
//             VídeoDepoimentos.setAttribute("allowfullscreen", "true");
//         }
        
//         ContainerVídeoPrincipal.appendChild(VídeoPrincipal);
  
//         observer.unobserve(entry.target);
//       }
    
//     });
  
// }
  
// var observer = new IntersectionObserver(handleIntersection);

// observer.observe(ContainerVídeoPrincipal);


// /*/////////////////////////////////////////////////////////////////////////////////////*/
// /*///////////////////////////////// Botão Tela Cheia //////////////////////////////////*/
// /*/////////////////////////////////////////////////////////////////////////////////////*/

// BotãoTelaCheia1.addEventListener("click", function(event) {
    
//     if (TextoTelaCheia1.innerHTML !== 'Tela Padrão') {

//         ContainerVídeoPrincipal.style.width = 'calc(var(--considered-screen-width) * 0.95)';
//         ContainerVídeoPrincipal.style.height = 'calc(var(--considered-screen-width) * 1.68858)';

//         var VídeoPrincipal = document.getElementById("Vídeo-Principal");

//         VídeoPrincipal.style.width = 'calc(var(--considered-screen-width) * 1.68858)';
//         VídeoPrincipal.style.height = 'calc(var(--considered-screen-width) * 0.95)';
//         VídeoPrincipal.style.transformOrigin = 'top left';
//         VídeoPrincipal.style.transform = 'rotate(90deg) translateY(-100%)';

//         TextoTelaCheia1.innerHTML = 'Tela Padrão';

//         VídeoPrincipal.scrollIntoView({behavior: 'smooth'});

//     } else {
        
//         ContainerVídeoPrincipal.style.width = 'calc(var(--considered-screen-width) * 0.90)';
//         ContainerVídeoPrincipal.style.height = 'calc(var(--considered-screen-width) * 0.50634)';

//         var VídeoPrincipal = document.getElementById("Vídeo-Principal");

//         VídeoPrincipal.style.width = 'calc(var(--considered-screen-width) * 0.90)';
//         VídeoPrincipal.style.height = 'calc(var(--considered-screen-width) * 0.50634)';
//         VídeoPrincipal.style.transformOrigin = '';
//         VídeoPrincipal.style.transform = '';

//         TextoTelaCheia1.innerHTML = 'Tela Cheia';

//         VídeoPrincipal.scrollIntoView({behavior: 'smooth'});

//     }

// })

// /*/////////////////////////////////////////////////////////////////////////////////////*/
// /*///////////////////// Controla as posições do Itens Dinâmicos ///////////////////////*/
// /*/////////////////////////////////////////////////////////////////////////////////////*/

// window.onscroll = function() {

//     var Posição_Seção1 = Seção1.offsetTop;

//     /*////////////////////// Controla a posição do Container-Botão-Cadastro ///////////////////////*/

//     if (window.scrollY + window.innerHeight >= Posição_Seção1){

//         ContainerBotãoCadastro.style.display = 'block';
//         EspaçoFinalContainerBotãoCadastro.style.height = ContainerBotãoCadastro.offsetHeight + 'px';

//         if (window.scrollY + window.innerHeight < Posição_Seção1 + ContainerBotãoCadastro.offsetHeight){

//             ContainerBotãoCadastro.style.position = 'absolute';
//             ContainerBotãoCadastro.style.top = Posição_Seção1 + 'px';
//             ContainerBotãoCadastro.style.bottom = '';

//         } else {

//             ContainerBotãoCadastro.style.position = 'fixed';
//             ContainerBotãoCadastro.style.top = '';
//             ContainerBotãoCadastro.style.bottom = '0px';

//         }
 
//     }

//     /*////////////////////// Controla a posição das Setas de Fechamento das Seções ///////////////////////*/

//     var Posição_Seção2 = Seção2.offsetTop;
//     var Posição_Seção3 = Seção3.offsetTop;
//     var Posição_Seção4 = Seção4.offsetTop;
//     var Posição_EspaçoFinalContainerBotãoCadastro = EspaçoFinalContainerBotãoCadastro.offsetTop;

//     /* Seção 1 */

//     if (window.scrollY <= Posição_Seção1) {
    
//         SetaFechamentoSeção1.style.display = 'none';
    
//     } else if (window.scrollY > Posição_Seção1 && window.scrollY <= (Posição_Seção2 - window.innerHeight + ContainerBotãoCadastro.offsetHeight)) {

//         SubseçãoCadastro1.style.marginBottom = '20px';
        
//         SetaFechamentoSeção1.style.display = 'flex';
//         SetaFechamentoSeção1.style.position = 'fixed';
//         SetaFechamentoSeção1.style.bottom = ContainerBotãoCadastro.offsetHeight + 15 + 'px';
//         SetaFechamentoSeção1.style.marginBottom = '0px';
//         SetaFechamentoSeção1.style.marginLeft = 'calc((var(--considered-screen-width) * 0.50) - 40px)';

//     } else if (window.scrollY > (Posição_Seção2 - window.innerHeight + ContainerBotãoCadastro.offsetHeight)) {

//         SubseçãoCadastro1.style.marginBottom = '15px';

//         SetaFechamentoSeção1.style.display = 'flex';
//         SetaFechamentoSeção1.style.position = 'relative';
//         SetaFechamentoSeção1.style.bottom = '0px';
//         SetaFechamentoSeção1.style.marginBottom = '-25px';
//         SetaFechamentoSeção1.style.marginLeft = 'calc((var(--considered-screen-width) * 0.50) - 40px)';

//     }

//     /* Seção 2 */

//     if (window.scrollY <= Posição_Seção2) {

//         SetaFechamentoSeção2.style.display = 'none';
    
//     } else if (window.scrollY > Posição_Seção2 && window.scrollY <= (Posição_Seção3 - window.innerHeight + ContainerBotãoCadastro.offsetHeight)) {

//         SubseçãoCadastro2.style.marginBottom = '20px';
        
//         SetaFechamentoSeção2.style.display = 'flex';
//         SetaFechamentoSeção2.style.position = 'fixed';
//         SetaFechamentoSeção2.style.bottom = ContainerBotãoCadastro.offsetHeight + 15 + 'px';
//         SetaFechamentoSeção2.style.marginBottom = '0px';
//         SetaFechamentoSeção2.style.marginLeft = 'calc((var(--considered-screen-width) * 0.50) - 40px)';

//     } else if (window.scrollY > (Posição_Seção3 - window.innerHeight + ContainerBotãoCadastro.offsetHeight)) {

//         SubseçãoCadastro2.style.marginBottom = '15px';

//         SetaFechamentoSeção2.style.display = 'flex';
//         SetaFechamentoSeção2.style.position = 'relative';
//         SetaFechamentoSeção2.style.bottom = '0px';
//         SetaFechamentoSeção2.style.marginBottom = '-25px';
//         SetaFechamentoSeção2.style.marginLeft = 'calc((var(--considered-screen-width) * 0.50) - 40px)';

//     }

//     /* Seção 3 */

//     if (window.scrollY <= Posição_Seção3) {

//         SetaFechamentoSeção3.style.display = 'none';
    
//     } else if (window.scrollY > Posição_Seção3 && window.scrollY <= (Posição_Seção4 - window.innerHeight + ContainerBotãoCadastro.offsetHeight)) {

//         SubseçãoCadastro3.style.marginBottom = '20px';
        
//         SetaFechamentoSeção3.style.display = 'flex';
//         SetaFechamentoSeção3.style.position = 'fixed';
//         SetaFechamentoSeção3.style.bottom = ContainerBotãoCadastro.offsetHeight + 15 + 'px';
//         SetaFechamentoSeção3.style.marginBottom = '0px';
//         SetaFechamentoSeção3.style.marginLeft = 'calc((var(--considered-screen-width) * 0.50) - 40px)';

//     } else if (window.scrollY > (Posição_Seção4 - window.innerHeight + ContainerBotãoCadastro.offsetHeight)) {

//         SubseçãoCadastro3.style.marginBottom = '15px';

//         SetaFechamentoSeção3.style.display = 'flex';
//         SetaFechamentoSeção3.style.position = 'relative';
//         SetaFechamentoSeção3.style.bottom = '0px';
//         SetaFechamentoSeção3.style.marginBottom = '-25px';
//         SetaFechamentoSeção3.style.marginLeft = 'calc((var(--considered-screen-width) * 0.50) - 40px)';

//     }

//     /* Seção 4 */

//     if (window.scrollY <= Posição_Seção4) {
    
//         SetaFechamentoSeção4.style.display = 'none';
//         BotãoInstagramDirect.style.display = 'none';

//     } else if (window.scrollY > Posição_Seção4 && window.scrollY <= (Posição_EspaçoFinalContainerBotãoCadastro - window.innerHeight + ContainerBotãoCadastro.offsetHeight)) {

//         ContainerCondiçõesComerciais3.style.marginBottom = '20px';
    
//         if (userAgent.indexOf('Instagram') === -1) {

//             BotãoInstagramDirect.style.display = 'block';
    
//         }

//         SetaFechamentoSeção4.style.display = 'flex';
//         SetaFechamentoSeção4.style.position = 'fixed';
//         SetaFechamentoSeção4.style.bottom = ContainerBotãoCadastro.offsetHeight + 15 + 'px';
//         SetaFechamentoSeção4.style.marginBottom = '0px';
//         SetaFechamentoSeção4.style.marginLeft = 'calc((var(--considered-screen-width) * 0.50) - 40px)';

//     } else if (window.scrollY > (Posição_EspaçoFinalContainerBotãoCadastro - window.innerHeight + ContainerBotãoCadastro.offsetHeight)) {

//         ContainerCondiçõesComerciais3.style.marginBottom = '15px';

//         SetaFechamentoSeção4.style.display = 'flex';
//         SetaFechamentoSeção4.style.position = 'relative';
//         SetaFechamentoSeção4.style.bottom = '0px';
//         SetaFechamentoSeção4.style.marginBottom = '-25px';
//         SetaFechamentoSeção4.style.marginLeft = 'calc((var(--considered-screen-width) * 0.50) - 40px)';

//     }

// }

// /*Abre o Cadastro*/
// BotãoCadastro.addEventListener("click", function(event){
//     event.preventDefault();
//     window.location.href = "cadastro/";
// })


// /*/////////////////////////////////////////////////////////////////////////////////////*/
// /*////////////////////////////////////// Seção 1 //////////////////////////////////////*/
// /*/////////////////////////////////////////////////////////////////////////////////////*/

// /*Abre a Seção 1*/
// ContainerExternoSeção1.addEventListener("click", function(event) {
//     ContainerExternoSeção1.style.display = 'none';
//     ContainerInternoSeção1.style.display = 'block';
//     ContainerExternoSeção2.style.display = 'block';
//     ContainerInternoSeção2.style.display = 'none';
//     ContainerExternoSeção3.style.display = 'block';
//     ContainerInternoSeção3.style.display = 'none';
//     ContainerExternoSeção4.style.display = 'block';
//     ContainerInternoSeção4.style.display = 'none';
//     ContainerInternoSeção1.scrollIntoView({behavior: 'smooth'});
//     SetaFechamentoSeção1.style.position = 'fixed';
//     SetaFechamentoSeção1.style.bottom = '25px';
// })



// /*Fecha a Seção 1*/
// SetaFechamentoSeção1.addEventListener("click", function(event) {
//     ContainerExternoSeção1.style.display = 'block';
//     ContainerInternoSeção1.style.display = 'none';
//     ContainerExternoSeção1.scrollIntoView({behavior: 'smooth'});
// })


// /*/////////////////////////////////////////////////////////////////////////////////////*/
// /*////////////////////////////////////// Seção 2 //////////////////////////////////////*/
// /*/////////////////////////////////////////////////////////////////////////////////////*/

// /*Abre a Seção 2*/
// ContainerExternoSeção2.addEventListener("click", function(event) {
//     ContainerExternoSeção1.style.display = 'block';
//     ContainerInternoSeção1.style.display = 'none';
//     ContainerExternoSeção2.style.display = 'none';
//     ContainerInternoSeção2.style.display = 'block';
//     ContainerExternoSeção3.style.display = 'block';
//     ContainerInternoSeção3.style.display = 'none';
//     ContainerExternoSeção4.style.display = 'block';
//     ContainerInternoSeção4.style.display = 'none';
//     ContainerInternoSeção2.scrollIntoView({behavior: 'smooth'});
//     SetaFechamentoSeção2.style.position = 'fixed';
//     SetaFechamentoSeção2.style.bottom = '25px';
// })

// /*Fecha a Seção 2*/
// SetaFechamentoSeção2.addEventListener("click", function(event) {
//     ContainerExternoSeção2.style.display = 'block';
//     ContainerInternoSeção2.style.display = 'none';
//     ContainerExternoSeção2.scrollIntoView({behavior: 'smooth'});
// })

// /*/////////////////////////////////////////////////////////////////////////////////////*/
// /*////////////////////////////////////// Seção 3 //////////////////////////////////////*/
// /*/////////////////////////////////////////////////////////////////////////////////////*/

// /*Abre a Seção 3*/
// ContainerExternoSeção3.addEventListener("click", function(event) {
//     ContainerExternoSeção1.style.display = 'block';
//     ContainerInternoSeção1.style.display = 'none';
//     ContainerExternoSeção2.style.display = 'block';
//     ContainerInternoSeção2.style.display = 'none';
//     ContainerExternoSeção3.style.display = 'none';
//     ContainerInternoSeção3.style.display = 'block';
//     ContainerExternoSeção4.style.display = 'block';
//     ContainerInternoSeção4.style.display = 'none';
//     ContainerInternoSeção3.scrollIntoView({behavior: 'smooth'});
//     SetaFechamentoSeção3.style.position = 'fixed';
//     SetaFechamentoSeção3.style.bottom = '25px';
// })

// /*Fecha a Seção 3*/
// SetaFechamentoSeção3.addEventListener("click", function(event) {
//     ContainerExternoSeção3.style.display = 'block';
//     ContainerInternoSeção3.style.display = 'none';
//     ContainerExternoSeção3.scrollIntoView({behavior: 'smooth'});
// })


// /*/////////////////////////////////////////////////////////////////////////////////////*/
// /*////////////////////////////////////// Seção 4 //////////////////////////////////////*/
// /*/////////////////////////////////////////////////////////////////////////////////////*/

// /*Abre a Seção 4*/
// ContainerExternoSeção4.addEventListener("click", function(event) {
//     ContainerExternoSeção1.style.display = 'block';
//     ContainerInternoSeção1.style.display = 'none';
//     ContainerExternoSeção2.style.display = 'block';
//     ContainerInternoSeção2.style.display = 'none';
//     ContainerExternoSeção3.style.display = 'block';
//     ContainerInternoSeção3.style.display = 'none';
//     ContainerExternoSeção4.style.display = 'none';
//     ContainerInternoSeção4.style.display = 'block';
//     ContainerInternoSeção4.scrollIntoView({behavior: 'smooth'});
//     SetaFechamentoSeção4.style.position = 'fixed';
//     SetaFechamentoSeção4.style.bottom = '25px';
//     BotãoInstagramDirect.style.position = 'fixed';
// })

// /*///////////////////////////////////// Seção 4-1 /////////////////////////////////////*/

// /*Abre a Seção 4-1*/
// ContainerExternoSubseção41.addEventListener("click", function(event) {
//     ContainerExternoSubseção41.style.display = 'none';
//     ContainerInternoSubseção41.style.display = 'block';
//     ContainerInternoSubseção41.scrollIntoView({behavior: 'smooth'});
// })

// /*Fecha a Subseção 4-1*/
// SetaFechamentoSubseção41.addEventListener("click", function(event) {
//     ContainerExternoSubseção41.style.display = 'block';
//     ContainerInternoSubseção41.style.display = 'none';
//     ContainerExternoSubseção41.scrollIntoView({behavior: 'smooth'});
// })

// /*///////////////////////////////////// Seção 4-2 /////////////////////////////////////*/

// /*Abre a Seção 4-2*/
// ContainerExternoSubseção42.addEventListener("click", function(event) {
//     ContainerExternoSubseção42.style.display = 'none';
//     ContainerInternoSubseção42.style.display = 'block';
//     ContainerInternoSubseção42.scrollIntoView({behavior: 'smooth'});
// })

// /*Fecha a Subseção 4-2*/
// SetaFechamentoSubseção42.addEventListener("click", function(event) {
//     ContainerExternoSubseção42.style.display = 'block';
//     ContainerInternoSubseção42.style.display = 'none';
//     ContainerExternoSubseção42.scrollIntoView({behavior: 'smooth'});
// })

// /*///////////////////////////////////// Seção 4-3 /////////////////////////////////////*/

// /*Abre a Seção 4-3*/
// ContainerExternoSubseção43.addEventListener("click", function(event) {
//     ContainerExternoSubseção43.style.display = 'none';
//     ContainerInternoSubseção43.style.display = 'block';
//     ContainerInternoSubseção43.scrollIntoView({behavior: 'smooth'});
// })

// /*Fecha a Subseção 4-3*/
// SetaFechamentoSubseção43.addEventListener("click", function(event) {
//     ContainerExternoSubseção43.style.display = 'block';
//     ContainerInternoSubseção43.style.display = 'none';
//     ContainerExternoSubseção43.scrollIntoView({behavior: 'smooth'});
// })

// /*/////////////////////////////////////////////////////////////////////////////////////*/
// /*///////////////////////////////// Botão Tela Cheia //////////////////////////////////*/
// /*/////////////////////////////////////////////////////////////////////////////////////*/

// BotãoTelaCheia2.addEventListener("click", function(event) {
    
//     if (TextoTelaCheia2.innerHTML !== 'Tela Padrão') {

//         ContainerVídeoDepoimentos.style.width = 'calc(var(--considered-screen-width) * 0.95)';
//         ContainerVídeoDepoimentos.style.height = 'calc(var(--considered-screen-width) * 1.68858)';

//         VídeoDepoimentos.style.width = 'calc(var(--considered-screen-width) * 1.68858)';
//         VídeoDepoimentos.style.height = 'calc(var(--considered-screen-width) * 0.95)';
//         VídeoDepoimentos.style.transformOrigin = 'top left';
//         VídeoDepoimentos.style.transform = 'rotate(90deg) translateY(-100%)';

//         TextoTelaCheia2.innerHTML = 'Tela Padrão';

//         VídeoDepoimentos.scrollIntoView({behavior: 'smooth'});

//     } else {
        
//         ContainerVídeoDepoimentos.style.width = 'calc(var(--considered-screen-width) * 0.90)';
//         ContainerVídeoDepoimentos.style.height = 'calc(var(--considered-screen-width) * 0.50634)';

//         VídeoDepoimentos.style.width = 'calc(var(--considered-screen-width) * 0.90)';
//         VídeoDepoimentos.style.height = 'calc(var(--considered-screen-width) * 0.50634)';
//         VídeoDepoimentos.style.transformOrigin = '';
//         VídeoDepoimentos.style.transform = '';

//         TextoTelaCheia2.innerHTML = 'Tela Cheia';

//         VídeoDepoimentos.scrollIntoView({behavior: 'smooth'});

//     }

// })

// /*////////////////////////////////// Demais Funções //////////////////////////////////*/

// /*Abre documentos em .pdf*/
// BotãoCronograma.addEventListener("click", function(event) {
//     event.preventDefault();
//     window.open("pdf/GESTÃO GENERALISTA - CRONOGRAMA DE ESTUDOS_v3.pdf","_blank");
// })

// BotãoEmenta.addEventListener("click", function(event) {
//     event.preventDefault();
//     window.open("pdf/GESTÃO GENERALISTA - EMENTA_v5.pdf","_blank");
// })

// BotãoBibliografia.addEventListener("click", function(event) {
//     event.preventDefault();
//     window.open("pdf/GESTÃO GENERALISTA - BIBLIOGRAFIA_v3.pdf","_blank");
// })

// BotãoSoftwares.addEventListener("click", function(event) {
//     event.preventDefault();
//     window.open("pdf/GESTÃO GENERALISTA - SOFTWARES_v3.pdf","_blank");
// })

// /*Fecha a Seção 4*/
// SetaFechamentoSeção4.addEventListener("click", function(event) {
//     ContainerExternoSeção4.style.display = 'block';
//     ContainerInternoSeção4.style.display = 'none';
//     ContainerExternoSeção4.scrollIntoView({behavior: 'smooth'});
// })

// /*Abre o Direct*/
// BotãoInstagramDirect.addEventListener("click", function(event) {
//     event.preventDefault();
//     window.open("https://ig.me/m/ivy.escoladegestao","_blank");
// })
