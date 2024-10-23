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
////////////////////////////////////////////////////////////////////////////////////////
// Importa todas as bibliotecas e faz as configurações iniciais necessárias.
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
    const date = new Date((DataExcel - 25568) * 86400 * 1000);
    return date.toLocaleDateString('pt-BR', { day: '2-digit', month: 'short', year: 'numeric' }).replace(/ de|\.|,/g, '').replace(/\s+/g, '/');
}

////////////////////////////////////////////////////////////////////////////////////////
// Processa submissão do formulário de login.
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
// Processa carregamento da aba /estudos no Frontend.
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
// Atualiza a BD - Plataforma.
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