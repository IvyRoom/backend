'use strict';

const { escapeHtml } = require('../shared/escape-html');

// Colunas da BD - RECOMENDAÇÕES (0-based), verificadas contra a planilha em 13/jul/2026. PRIMEIRO NOME é coluna calculada — deixar null em linhas novas.
const RECOMENDACOES_COLUMNS = { benefitedCompany: 0, recommenderFullName: 1, recommenderFirstName: 2, recommenderEmail: 3, dateTime: 4, recommendedCompany: 5, recommendedProfessional: 6, recommendedWhatsapp: 7, stage: 8, status: 9, updateDateTime: 10, nextContactDateTime: 11, participantsCount: 12 };
const RECOMENDACOES_ROW_WIDTH = 13;

// Valores devem espelhar as listas da aba AUXILIAR da planilha.
const RECOMENDACOES_INITIAL_STAGE = '1. REALIZAR CONTATO INICIAL';
const RECOMENDACOES_INITIAL_STATUS = 'A INICIAR';

const CONECTA_WHATSAPP_PATTERN = /^\+\d{2} \d{2} \d{5}-\d{4}$/;

function normalizeMatchKey(value) {
    return String(value == null ? '' : value).trim().replace(/\s+/g, ' ').toLowerCase();
}

function isPlaceholderCell(value) {
    return String(value == null ? '' : value).trim() === '-';
}

function createConectaRecommendationHandler({ microsoftGraph, retry, now }) {
    // Serial Excel de data e hora atuais no fuso de Brasília — a exibição fica na formatação da planilha.
    function nowBrazilSerial() {
        const [datePart, timePart] = now().toLocaleString('pt-BR', { timeZone: 'America/Sao_Paulo', hour12: false }).split(', ');
        const [day, month, year] = datePart.split('/').map(Number);
        const [hour, minute, second] = timePart.split(':').map(Number);
        return Math.floor(Date.UTC(year, month - 1, day) / 86400000) + 25569 + (hour * 3600 + minute * 60 + second) / 86400;
    }

    return async function processConectaRecommendation(req, res) {
        const body = req.body || {};
        const isNonEmptyString = (value) => typeof value === 'string' && value.trim() !== '';
        const requiredFields = ['recommenderFullName', 'benefitedCompany', 'recommendedCompany', 'recommendedProfessional', 'recommendedWhatsapp'];
        if (!requiredFields.every((field) => isNonEmptyString(body[field]))) return res.status(400).json({ error: 'Erro_014' });
        if (!CONECTA_WHATSAPP_PATTERN.test(body.recommendedWhatsapp.trim())) return res.status(400).json({ error: 'Erro_014' });

        let recomendacoesResponse;
        try { recomendacoesResponse = await retry(() => microsoftGraph.readRecommendationRows()); }
        catch (err) { console.error('conecta Erro_015:', err); return res.status(500).json({ error: 'Erro_015' }); }
        const recomendacoesData = microsoftGraph.extractRows(recomendacoesResponse);

        const columns = RECOMENDACOES_COLUMNS;
        const recommenderNameKey = normalizeMatchKey(body.recommenderFullName);
        const benefitedCompanyKey = normalizeMatchKey(body.benefitedCompany);

        const recommenderRows = recomendacoesData
            .map((row, index) => ({ index, cells: microsoftGraph.extractRowCells(row) }))
            .filter(({ cells }) => normalizeMatchKey(cells[columns.recommenderFullName]) === recommenderNameKey && normalizeMatchKey(cells[columns.benefitedCompany]) === benefitedCompanyKey);

        if (recommenderRows.length === 0) return res.status(404).json({ error: 'Erro_016' });

        // Reenvio idêntico (ex.: retry após falha de e-mail) não duplica linha; e-mails seguem adiante mesmo assim.
        const isDuplicate = recommenderRows.some(({ cells }) =>
            normalizeMatchKey(cells[columns.recommendedCompany]) === normalizeMatchKey(body.recommendedCompany)
            && normalizeMatchKey(cells[columns.recommendedProfessional]) === normalizeMatchKey(body.recommendedProfessional)
            && normalizeMatchKey(cells[columns.recommendedWhatsapp]) === normalizeMatchKey(body.recommendedWhatsapp));

        const recommenderCells = recommenderRows[0].cells;

        if (!isDuplicate) {
            const recommendationColumns = [columns.dateTime, columns.recommendedCompany, columns.recommendedProfessional, columns.recommendedWhatsapp, columns.stage, columns.status, columns.updateDateTime, columns.nextContactDateTime, columns.participantsCount];
            const slotRow = recommenderRows.find(({ cells }) => recommendationColumns.every((column) => isPlaceholderCell(cells[column])));

            const currentTime = nowBrazilSerial();
            const cells = new Array(RECOMENDACOES_ROW_WIDTH).fill(null);
            cells[columns.dateTime] = currentTime;
            cells[columns.recommendedCompany] = body.recommendedCompany.trim();
            cells[columns.recommendedProfessional] = body.recommendedProfessional.trim();
            cells[columns.recommendedWhatsapp] = body.recommendedWhatsapp.trim();
            cells[columns.stage] = RECOMENDACOES_INITIAL_STAGE;
            cells[columns.status] = RECOMENDACOES_INITIAL_STATUS;
            cells[columns.updateDateTime] = currentTime;
            cells[columns.nextContactDateTime] = currentTime;

            // Escritas deliberadamente sem retry(): uma falha ambígua após inserção bem-sucedida duplicaria a linha.
            try {
                if (slotRow) {
                    await microsoftGraph.updateRecommendationRow(slotRow.index, cells);
                } else {
                    cells[columns.benefitedCompany] = recommenderCells[columns.benefitedCompany];
                    cells[columns.recommenderFullName] = recommenderCells[columns.recommenderFullName];
                    cells[columns.recommenderEmail] = recommenderCells[columns.recommenderEmail];
                    cells[columns.participantsCount] = '-';
                    await microsoftGraph.appendRecommendationRow(cells);
                }
            } catch (err) { console.error('conecta Erro_017:', err); return res.status(500).json({ error: 'Erro_017' }); }
        }

        const recommenderEmail = String(recommenderCells[columns.recommenderEmail] == null ? '' : recommenderCells[columns.recommenderEmail]).trim();
        const recommenderFirstNameCell = String(recommenderCells[columns.recommenderFirstName] == null ? '' : recommenderCells[columns.recommenderFirstName]).trim();
        const recommenderFirstName = recommenderFirstNameCell && recommenderFirstNameCell !== '-' ? recommenderFirstNameCell : String(recommenderCells[columns.recommenderFullName]).trim().split(/\s+/)[0];
        const signatureHTML = '<p><img width="600" height="auto" src="https://plataforma-backend-v3.azurewebsites.net/img/ASSINATURA_E-MAIL.jpg"/></p>';

        const internalEmailContent = `<p><b>Dados do Recomendante:</b></p><p>Nome Completo: ${escapeHtml(recommenderCells[columns.recommenderFullName])}</p><p>E-mail: ${escapeHtml(recommenderEmail)}</p><p>Empresa Beneficiada: ${escapeHtml(recommenderCells[columns.benefitedCompany])}</p><p><b>Dados da Recomendação:</b></p><p>Empresa Recomendada: ${escapeHtml(body.recommendedCompany.trim())}</p><p>Profissional Contatado: ${escapeHtml(body.recommendedProfessional.trim())}</p><p>WhatsApp do Profissional: ${escapeHtml(body.recommendedWhatsapp.trim())}</p>${signatureHTML}`;

        const confirmationEmailContent = `<p>Olá ${escapeHtml(recommenderFirstName)},</p><p>Recebemos sua recomendação da Machado para a empresa <b>${escapeHtml(body.recommendedCompany.trim())}</b>. Obrigado pela confiança.</p><p>Logo entraremos em contato com ${escapeHtml(body.recommendedProfessional.trim())}. Assim que houver atualizações relevantes, sinalizaremos a você.</p><p>Atenciosamente,</p>${signatureHTML}`;

        try {
            await retry(() => microsoftGraph.sendMail({ subject: 'Machado Conecta - Nova Recomendação Recebida', body: { contentType: 'HTML', content: internalEmailContent }, toRecipients: [{ emailAddress: { address: 'contato@machadogestao.com' } }] }));
            await retry(() => microsoftGraph.sendMail({ subject: 'Machado Conecta - Recomendação Registrada', body: { contentType: 'HTML', content: confirmationEmailContent }, toRecipients: [{ emailAddress: { address: recommenderEmail } }] }));
        } catch (err) { console.error('conecta Erro_018:', err); return res.status(500).json({ error: 'Erro_018' }); }

        return res.status(200).json({});
    };
}

module.exports = { createConectaRecommendationHandler };
