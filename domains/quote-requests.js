'use strict';

function createQuoteRequestHandler({ microsoftGraph, retry }) {
    return async function processQuoteRequest(req, res) {
        let { Solicitante_NomeCompleto, Solicitante_Email, Solicitante_Telefone, Solicitante_Cargo, Solicitante_NomeEmpresa, Solicitante_CNPJ, Solicitante_NúmerodeParticipantes, Solicitante_Observações } = req.body;

        try { await retry(() => microsoftGraph.sendMail({ subject: 'Machado - Nova Solicitação de Orçamento', body: { contentType: 'HTML', content: `<p><b>Dados do Solicitante:</b></p><p>${Solicitante_NomeCompleto}</p><p>${Solicitante_Email}</p><p>${Solicitante_Telefone}</p><p>${Solicitante_Cargo}</p><p><b>Dados da Empresa:</b></p><p>${Solicitante_NomeEmpresa}</p><p>${Solicitante_CNPJ}</p><p>${Solicitante_NúmerodeParticipantes}</p><p>${Solicitante_Observações}</p><p><img width="500" height="auto" src="https://plataforma-backend-v3.azurewebsites.net/img/ASSINATURA_E-MAIL.jpg"/></p>`}, toRecipients: [{ emailAddress: { address: 'contato@machadogestao.com' } }] })); }
        catch (err) { return res.status(500).json({}); }

        return res.status(200).json({});
    };
}

module.exports = { createQuoteRequestHandler };
