'use strict';

const { escapeHtml } = require('../shared/escape-html');

const CERTIFICATE_ID_ALPHABET = '0123456789ABCDEFGHJKMNPQRSTVWXYZ';

function createClientOnboardingHandlers({
    microsoftGraph,
    retry,
    now,
    randomInt,
    sleep,
    schedule,
}) {
    function accessDeadlineSerial(daysFromToday) {
        const today = now();
        const utcMidnight = Date.UTC(today.getFullYear(), today.getMonth(), today.getDate());
        return Math.floor(utcMidnight / 86400000) + 25569 + daysFromToday;
    }

    function generateCertificateId() {
        let suffix = '';
        for (let i = 0; i < 8; i++) suffix += CERTIFICATE_ID_ALPHABET[randomInt(CERTIFICATE_ID_ALPHABET.length)];
        return `FMG-${suffix.slice(0, 4)}-${suffix.slice(4)}`;
    }

    async function processClientIntake(req, res) {
        const participants = Array.isArray(req.body && req.body.participants) ? req.body.participants : [];
        const company = (req.body && req.body.company) || {};
        const shipping = (req.body && req.body.shippingAddress) || {};
        const legalRep = (req.body && req.body.legalRepresentative) || {};
        const adminAssistant = (req.body && req.body.adminAssistant) || {};

        // Limite de 25 espelha MAX_PARTICIPANTS do formulário (sistemas/formulario/main.js).
        const isNonEmptyString = (value) => typeof value === 'string' && value.trim() !== '';
        const validPayload = isNonEmptyString(company.legalName)
            && participants.length >= 1 && participants.length <= 25
            && participants.every((p) => p && isNonEmptyString(p.fullName) && isNonEmptyString(p.email) && isNonEmptyString(p.cpf));
        if (!validPayload) return res.status(400).json({ error: 'Erro_013' });

        let plataformaResponse, clientesResponse;
        try { plataformaResponse = await retry(() => microsoftGraph.readPlatformRows()); }
        catch (err) { return res.status(500).json({ error: 'Erro_001' }); }
        try { clientesResponse = await retry(() => microsoftGraph.readClientRows()); }
        catch (err) { return res.status(500).json({ error: 'Erro_011' }); }
        const plataformaData = microsoftGraph.extractRows(plataformaResponse);
        const clientesData = microsoftGraph.extractRows(clientesResponse);

        const onlyDigits = (value) => String(value == null ? '' : value).replace(/\D/g, '');
        const existingEmails = new Set(plataformaData.map((row) => {
            const cells = microsoftGraph.extractRowCells(row);
            return String(cells[2] == null ? '' : cells[2]).trim().toLowerCase();
        }));
        const existingCpfs = new Set(clientesData.map((row) => onlyDigits(microsoftGraph.extractRowCells(row)[4])));
        const existingCertificateIds = new Set(plataformaData.map((row) => {
            const cells = microsoftGraph.extractRowCells(row);
            return String(cells[21] == null ? '' : cells[21]).trim().toUpperCase();
        }).filter(Boolean));

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
                cells[3] = randomInt(100000000000, 1000000000000);
                cells[4] = 'Ativo';
                cells[5] = 'Não';
                cells[6] = deadline;
                cells[8] = 0;
                for (let module = 10; module <= 19; module++) cells[module] = 0;
                let certificateId;
                do { certificateId = generateCertificateId(); } while (existingCertificateIds.has(certificateId));
                existingCertificateIds.add(certificateId);
                cells[21] = certificateId;
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
            try { await microsoftGraph.appendPlatformRows(plataformaRows); }
            catch (err) { return res.status(500).json({ error: 'Erro_008' }); }
        }

        if (clientesRows.length > 0) {
            try { await microsoftGraph.appendClientRows(clientesRows); }
            catch (err) { return res.status(500).json({ error: 'Erro_010' }); }
        }

        const companyAddress = company.address || {};
        const pessoaHTML = (rotulo, p) => `<p><b>${rotulo}</b></p><p>Nome Completo: ${escapeHtml(p.fullName)}</p><p>CPF: ${escapeHtml(p.cpf)}</p><p>Cargo: ${escapeHtml(p.role)}</p><p>DDD: ${escapeHtml(p.areaCode)}</p><p>WhatsApp: ${escapeHtml(p.whatsapp)}</p><p>E-mail: ${escapeHtml(p.email)}</p>`;
        const participantesHTML = participants.map((p, i) => `<p>${i + 1}. ${escapeHtml(p.fullName)} — Cargo: ${escapeHtml(p.role)} · DDD: ${escapeHtml(p.areaCode)} · WhatsApp: ${escapeHtml(p.whatsapp)}</p>`).join('');
        const emailContent = `<p>Um novo Formulário de Informações Iniciais foi preenchido.</p><p><b>Pessoa Jurídica Contratante</b></p><p>Razão Social: ${escapeHtml(company.legalName)}</p><p>CNPJ: ${escapeHtml(company.cnpj)}</p><p>CEP: ${escapeHtml(companyAddress.postalCode)}</p><p>Rua: ${escapeHtml(companyAddress.street)}</p><p>Número: ${escapeHtml(companyAddress.number)}</p><p>Complemento: ${escapeHtml(companyAddress.complement)}</p><p>Bairro: ${escapeHtml(companyAddress.neighborhood)}</p><p>Cidade: ${escapeHtml(companyAddress.city)}</p><p>Estado: ${escapeHtml(companyAddress.state)}</p>${pessoaHTML('Representante Jurídico', legalRep)}${pessoaHTML('Auxiliar Administrativo Financeiro', adminAssistant)}<p><b>Participantes</b></p>${participantesHTML}<p><img width="500" height="auto" src="https://plataforma-backend-v3.azurewebsites.net/img/ASSINATURA_E-MAIL.jpg"/></p>`;

        try { await retry(() => microsoftGraph.sendMail({ subject: 'Machado: novo Formulário de Informações Iniciais preenchido', body: { contentType: 'HTML', content: emailContent }, toRecipients: [{ emailAddress: { address: 'contato@machadogestao.com' } }] })); }
        catch (err) { return res.status(500).json({ error: 'Erro_012' }); }

        return res.status(200).json({});
    }

    async function releasePlatformAccess(req, res) {
        res.status(200).send();
        console.log(`1. Request recebida.`);

        const BD_PlataformaResponse = await microsoftGraph.readPlatformRows();
        if (BD_PlataformaResponse !== null) console.log(`2. BD_Plataforma obtida.`);

        let Número_Email_Enviado = 0;
        let Linha_Inicial = 39;
        let Linha_Final = 45;

        async function Envia_Email_Clientes() {
            const BD_Plataforma = microsoftGraph.extractRows(BD_PlataformaResponse);
            for (let LinhaAtual = (Linha_Inicial - 4); LinhaAtual <= (Linha_Final - 4); LinhaAtual++) {
                const cells = microsoftGraph.extractRowCells(BD_Plataforma[LinhaAtual]);
                let Cliente_PrimeiroNome = cells[1];
                let Cliente_Email = cells[2];
                let Cliente_Senha = cells[3];

                Número_Email_Enviado++;

                console.log(`3. E-mail #${Número_Email_Enviado} enviado para: ${Cliente_PrimeiroNome}`);

                if (LinhaAtual === (Linha_Final - 4)) console.log(`--- fim ---`);

                await microsoftGraph.sendMail({
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
                        `,
                    },
                    toRecipients: [{ emailAddress: { address: Cliente_Email } }],
                });

                await sleep(2000);
            }
        }

        schedule(Envia_Email_Clientes, 1000);
    }

    return {
        processClientIntake,
        releasePlatformAccess,
    };
}

module.exports = { createClientOnboardingHandlers };
