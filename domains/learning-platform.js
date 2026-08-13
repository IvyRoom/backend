'use strict';

const PLATFORM_TABLE_PATH = '/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECSBYCZNYGEWFFDLEOZ36WI2PDWO/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}';
const PLATFORM_ROWS_PATH = `${PLATFORM_TABLE_PATH}/rows`;
const FEEDBACK_ROWS_PATH = '/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECXO7I5R6LKLXJD3VWXORUAF7J37/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows/add';

function ConverteData(DataExcel) {
    const date = new Date((DataExcel - 25569) * 86400 * 1000);
    return date.toLocaleDateString('pt-BR', { day: '2-digit', month: 'short', year: 'numeric' }).replace(/\bde\b|\./g, '').replace(/\s+/g, '/');
}

function referencePhotoPath(platformRowIndex) {
    return `/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/root:/2. ENTREGA/1. CONTROLAR PLATAFORMA/PG - FOTOS DE REFERÊNCIA/${platformRowIndex}.jpg:/content`;
}

function createLearningPlatformHandlers({
    graphClient,
    faceClient,
    retry,
    uuid: createUuid,
    createPlatformRowAuthorizationHandle,
}) {
    async function loginWithFaceId(req, res) {
        let { Usuário_Login, Usuário_Senha } = req.body;

        let BD_Plataforma;
        try { BD_Plataforma = await retry(() => graphClient.api(PLATFORM_ROWS_PATH).get()); }
        catch (err) { return res.status(500).json({ error: 'Erro_001' }); }

        for (let i = 0; i < BD_Plataforma.value.length; i++) {
            let LinhaVerificada = BD_Plataforma.value[i].values[0];
            if (Usuário_Login === LinhaVerificada[2] && Usuário_Senha === LinhaVerificada[3].toString()) {
                const RespostaLogin = { Usuário_Status_FaceID: LinhaVerificada[4], Usuário_Foto_Cadastrada: LinhaVerificada[5], Usuário_PrazoAcesso: ConverteData(LinhaVerificada[6]), Usuário_Status_Login: LinhaVerificada[7] };
                if (LinhaVerificada[7] === 'Ativo') RespostaLogin.IndexVerificado = createPlatformRowAuthorizationHandle(i);
                return res.status(200).json(RespostaLogin);
            }
        }

        return res.status(401).json({ error: 'credenciais_inválidas' });
    }

    async function registerPhotoAndFaceId(req, res) {
        const platformRowIndex = res.locals.platformRowIndex;
        let FotoReferência = req.file.buffer;

        try { await retry(() => graphClient.api(referencePhotoPath(platformRowIndex)).put(FotoReferência)); }
        catch (err) { return res.status(500).json({ error: 'Erro_002' }); }

        try { await retry(() => graphClient.api(`${PLATFORM_ROWS_PATH}/itemAt(index=${platformRowIndex})`).update({ values: [[null, null, null, null, null, 'Sim', null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null]] })); }
        catch (err) { return res.status(500).json({ error: 'Erro_003' }); }

        let Azure_Face_API_LivenessSession;
        try { Azure_Face_API_LivenessSession = await retry(() => faceClient.path('/detectLivenessWithVerify-sessions').post({ contentType: 'multipart/form-data', body: [{ name: 'VerifyImage', body: FotoReferência }, { name: 'livenessOperationMode', body: 'Passive' }, { name: 'deviceCorrelationId', body: createUuid() }] })); }
        catch (err) { return res.status(500).json({ error: 'Erro_004' }); }

        let Azure_Face_API_LivenessSession_authToken = Azure_Face_API_LivenessSession.body.authToken;
        let Azure_Face_API_LivenessSession_sessionID = Azure_Face_API_LivenessSession.body.sessionId;

        return res.status(200).json({ Azure_Face_API_LivenessSession_authToken, Azure_Face_API_LivenessSession_sessionID });
    }

    async function createFaceIdSession(req, res) {
        const platformRowIndex = res.locals.platformRowIndex;

        let FotoReferência;
        try { FotoReferência = await retry(() => graphClient.api(referencePhotoPath(platformRowIndex)).get()); }
        catch (err) { return res.status(500).json({ error: 'Erro_005' }); }

        let Azure_Face_API_LivenessSession;
        try { Azure_Face_API_LivenessSession = await retry(() => faceClient.path('/detectLivenessWithVerify-sessions').post({ contentType: 'multipart/form-data', body: [{ name: 'VerifyImage', body: FotoReferência }, { name: 'livenessOperationMode', body: 'Passive' }, { name: 'deviceCorrelationId', body: createUuid() }] })); }
        catch (err) { return res.status(500).json({ error: 'Erro_004' }); }

        let Azure_Face_API_LivenessSession_authToken = Azure_Face_API_LivenessSession.body.authToken;
        let Azure_Face_API_LivenessSession_sessionID = Azure_Face_API_LivenessSession.body.sessionId;

        return res.status(200).json({ Azure_Face_API_LivenessSession_authToken, Azure_Face_API_LivenessSession_sessionID });
    }

    async function getFaceIdResult(req, res) {
        let Azure_Face_API_LivenessSession_sessionID = req.params.Azure_Face_API_LivenessSession_sessionID;

        let Azure_Face_API_LivenessSession;
        try { Azure_Face_API_LivenessSession = await retry(() => faceClient.path('/detectLivenessWithVerify-sessions/{sessionId}', Azure_Face_API_LivenessSession_sessionID).get()); }
        catch (err) { return res.status(500).json({ error: 'Erro_007' }); }

        let Azure_Face_API_LivenessSession_LivenessDecision = Azure_Face_API_LivenessSession.body.results.attempts[0].result.livenessDecision;
        let Azure_Face_API_LivenessSession_MatchConfidence = Azure_Face_API_LivenessSession.body.results.attempts[0].result.verifyResult.matchConfidence;
        let Azure_Face_API_LivenessSession_MatchDecision = Azure_Face_API_LivenessSession.body.results.attempts[0].result.verifyResult.isIdentical;

        return res.status(200).json({ Azure_Face_API_LivenessSession_LivenessDecision, Azure_Face_API_LivenessSession_MatchConfidence, Azure_Face_API_LivenessSession_MatchDecision });
    }

    async function refresh(req, res) {
        const platformRowIndex = res.locals.platformRowIndex;

        let BD_Plataforma;
        try { BD_Plataforma = await retry(() => graphClient.api(PLATFORM_ROWS_PATH).get()); }
        catch (err) { return res.status(500).json({ error: 'Erro_001' }); }

        let Usuário_NomeCompleto = BD_Plataforma.value[platformRowIndex].values[0][0];
        let Usuário_PrimeiroNome = BD_Plataforma.value[platformRowIndex].values[0][1];
        let Usuário_Email = BD_Plataforma.value[platformRowIndex].values[0][2];
        let Usuário_PrazoAcesso = ConverteData(BD_Plataforma.value[platformRowIndex].values[0][6]);
        let Usuário_Status_Login = BD_Plataforma.value[platformRowIndex].values[0][7];
        let Usuário_Formação_NúmeroTópicosConcluídos = BD_Plataforma.value[platformRowIndex].values[0][8];
        let Usuário_Formação_NotaMódulo1 = BD_Plataforma.value[platformRowIndex].values[0][10];
        let Usuário_Formação_NotaMódulo2 = BD_Plataforma.value[platformRowIndex].values[0][11];
        let Usuário_Formação_NotaMódulo3 = BD_Plataforma.value[platformRowIndex].values[0][12];
        let Usuário_Formação_NotaMódulo4 = BD_Plataforma.value[platformRowIndex].values[0][13];
        let Usuário_Formação_NotaMódulo5 = BD_Plataforma.value[platformRowIndex].values[0][14];
        let Usuário_Formação_NotaMódulo6 = BD_Plataforma.value[platformRowIndex].values[0][15];
        let Usuário_Formação_NotaMódulo7 = BD_Plataforma.value[platformRowIndex].values[0][16];
        let Usuário_Formação_NotaMódulo8 = BD_Plataforma.value[platformRowIndex].values[0][17];
        let Usuário_Formação_NotaMódulo9 = BD_Plataforma.value[platformRowIndex].values[0][18];
        let Usuário_Formação_NotaMódulo10 = BD_Plataforma.value[platformRowIndex].values[0][19];
        let Usuário_Formação_NotaAcumulado = BD_Plataforma.value[platformRowIndex].values[0][20];
        let Usuário_Formação_CertificadoID = BD_Plataforma.value[platformRowIndex].values[0][21];

        return res.status(200).json({ Usuário_NomeCompleto, Usuário_PrimeiroNome, Usuário_Email, Usuário_PrazoAcesso, Usuário_Status_Login, Usuário_Formação_NúmeroTópicosConcluídos, Usuário_Formação_NotaMódulo1, Usuário_Formação_NotaMódulo2, Usuário_Formação_NotaMódulo3, Usuário_Formação_NotaMódulo4, Usuário_Formação_NotaMódulo5, Usuário_Formação_NotaMódulo6, Usuário_Formação_NotaMódulo7, Usuário_Formação_NotaMódulo8, Usuário_Formação_NotaMódulo9, Usuário_Formação_NotaMódulo10, Usuário_Formação_NotaAcumulado, Usuário_Formação_CertificadoID });
    }

    async function updateProgress(req, res) {
        const platformRowIndex = res.locals.platformRowIndex;
        let { TipoAtualização, NúmeroTópicosConcluídos, NúmeroMódulo, NotaTeste } = req.body;

        let DadosaInserir = new Array(22).fill(null);
        DadosaInserir[8] = NúmeroTópicosConcluídos;

        if (TipoAtualização === 'NúmeroTópicosConcluídos-e-NotaTeste') DadosaInserir[NúmeroMódulo + 9] = NotaTeste;

        try { await retry(() => graphClient.api(`${PLATFORM_ROWS_PATH}/itemAt(index=${platformRowIndex})`).update({ values: [DadosaInserir] })); }
        catch (err) { return res.status(500).json({ error: 'Erro_008' }); }

        return res.status(200).json({});
    }

    async function processFeedback(req, res) {
        const platformRowIndex = res.locals.platformRowIndex;
        let { NúmeroTópicosConcluídos, Usuário_NomeCompleto, Usuário_Email, Feedback_DataPreenchimento, NúmeroMódulo, Feedback_TamanhoMódulo, Feedback_QualidadeConteúdo, Feedback_QualidadePlataforma, Feedback_QualidadeMateriaisImpressos, Feedback_Comentários } = req.body;

        try { await retry(() => graphClient.api(`${PLATFORM_ROWS_PATH}/itemAt(index=${platformRowIndex})`).update({ values: [[null, null, null, null, null, null, null, null, NúmeroTópicosConcluídos, null, null, null, null, null, null, null, null, null, null, null, null, null]] })); }
        catch (err) { return res.status(500).json({ error: 'Erro_008' }); }

        try { await retry(() => graphClient.api(FEEDBACK_ROWS_PATH).post({ values: [[Usuário_NomeCompleto, Usuário_Email, Feedback_DataPreenchimento, NúmeroMódulo, Feedback_TamanhoMódulo, Feedback_QualidadeConteúdo, Feedback_QualidadePlataforma, Feedback_QualidadeMateriaisImpressos, Feedback_Comentários]] })); }
        catch (err) { return res.status(500).json({ error: 'Erro_009' }); }

        return res.status(200).json({});
    }

    async function getStatusReport(req, res) {
        let { linha_inicial, linha_final } = req.body;

        let BD_Plataforma;
        try { BD_Plataforma = await retry(() => graphClient.api(PLATFORM_ROWS_PATH).get()); }
        catch (err) { return res.status(500).json({ error: 'Erro_001' }); }

        const Dados_Extraídos_BD_Plataforma = BD_Plataforma.value.slice(linha_inicial, linha_final + 1).map(({ values }) => [values[0][0], values[0][8], ...values[0].slice(10, 22)]);

        return res.status(200).json({ Dados_Extraídos_BD_Plataforma });
    }

    return {
        loginWithFaceId,
        registerPhotoAndFaceId,
        createFaceIdSession,
        getFaceIdResult,
        refresh,
        updateProgress,
        processFeedback,
        getStatusReport,
    };
}

module.exports = { createLearningPlatformHandlers };
