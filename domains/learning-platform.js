'use strict';

const learningPlatformErrorValues = Object.freeze({
    platformDataReadFailure: 'learning_platform.read_platform_data_failed',
    referencePhotoUploadFailure: 'learning_platform.upload_reference_photo_failed',
    referencePhotoRegistrationUpdateFailure: 'learning_platform.update_reference_photo_registration_failed',
    faceLivenessSessionCreationFailure: 'learning_platform.create_face_liveness_session_failed',
    referencePhotoReadFailure: 'learning_platform.read_reference_photo_failed',
    faceLivenessResultReadFailure: 'learning_platform.read_face_liveness_result_failed',
    platformDataWriteFailure: 'learning_platform.update_platform_data_failed',
    feedbackAppendFailure: 'learning_platform.append_feedback_failed',
});

function ConverteData(DataExcel) {
    const date = new Date((DataExcel - 25569) * 86400 * 1000);
    return date.toLocaleDateString('pt-BR', { day: '2-digit', month: 'short', year: 'numeric' }).replace(/\bde\b|\./g, '').replace(/\s+/g, '/');
}

function createLearningPlatformHandlers({
    microsoftGraph,
    azureFace,
    retry,
    uuid: createUuid,
    createPlatformRowAuthorizationHandle,
}) {
    async function loginWithFaceId(req, res) {
        let { Usuário_Login, Usuário_Senha } = req.body;

        let BD_PlataformaResponse;
        try { BD_PlataformaResponse = await retry(() => microsoftGraph.readPlatformRows()); }
        catch (err) { return res.status(500).json({ error: learningPlatformErrorValues.platformDataReadFailure }); }
        const BD_Plataforma = microsoftGraph.extractRows(BD_PlataformaResponse);

        for (let i = 0; i < BD_Plataforma.length; i++) {
            let LinhaVerificada = microsoftGraph.extractRowCells(BD_Plataforma[i]);
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

        try { await retry(() => microsoftGraph.uploadReferencePhoto(platformRowIndex, FotoReferência)); }
        catch (err) { return res.status(500).json({ error: learningPlatformErrorValues.referencePhotoUploadFailure }); }

        try { await retry(() => microsoftGraph.updatePlatformRow(platformRowIndex, [null, null, null, null, null, 'Sim', null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null])); }
        catch (err) { return res.status(500).json({ error: learningPlatformErrorValues.referencePhotoRegistrationUpdateFailure }); }

        let Azure_Face_API_LivenessSessionResponse;
        try { Azure_Face_API_LivenessSessionResponse = await retry(() => azureFace.createLivenessSession(FotoReferência, createUuid())); }
        catch (err) { return res.status(500).json({ error: learningPlatformErrorValues.faceLivenessSessionCreationFailure }); }
        const Azure_Face_API_LivenessSession = azureFace.extractLivenessSession(Azure_Face_API_LivenessSessionResponse);

        let Azure_Face_API_LivenessSession_authToken = Azure_Face_API_LivenessSession.authToken;
        let Azure_Face_API_LivenessSession_sessionID = Azure_Face_API_LivenessSession.sessionId;

        return res.status(200).json({ Azure_Face_API_LivenessSession_authToken, Azure_Face_API_LivenessSession_sessionID });
    }

    async function createFaceIdSession(req, res) {
        const platformRowIndex = res.locals.platformRowIndex;

        let FotoReferência;
        try { FotoReferência = await retry(() => microsoftGraph.downloadReferencePhoto(platformRowIndex)); }
        catch (err) { return res.status(500).json({ error: learningPlatformErrorValues.referencePhotoReadFailure }); }

        let Azure_Face_API_LivenessSessionResponse;
        try { Azure_Face_API_LivenessSessionResponse = await retry(() => azureFace.createLivenessSession(FotoReferência, createUuid())); }
        catch (err) { return res.status(500).json({ error: learningPlatformErrorValues.faceLivenessSessionCreationFailure }); }
        const Azure_Face_API_LivenessSession = azureFace.extractLivenessSession(Azure_Face_API_LivenessSessionResponse);

        let Azure_Face_API_LivenessSession_authToken = Azure_Face_API_LivenessSession.authToken;
        let Azure_Face_API_LivenessSession_sessionID = Azure_Face_API_LivenessSession.sessionId;

        return res.status(200).json({ Azure_Face_API_LivenessSession_authToken, Azure_Face_API_LivenessSession_sessionID });
    }

    async function getFaceIdResult(req, res) {
        let Azure_Face_API_LivenessSession_sessionID = req.params.Azure_Face_API_LivenessSession_sessionID;

        let Azure_Face_API_LivenessSessionResponse;
        try { Azure_Face_API_LivenessSessionResponse = await retry(() => azureFace.readLivenessSessionResult(Azure_Face_API_LivenessSession_sessionID)); }
        catch (err) { return res.status(500).json({ error: learningPlatformErrorValues.faceLivenessResultReadFailure }); }
        const Azure_Face_API_LivenessSession = azureFace.extractLivenessSessionResult(Azure_Face_API_LivenessSessionResponse);

        let Azure_Face_API_LivenessSession_LivenessDecision = Azure_Face_API_LivenessSession.livenessDecision;
        let Azure_Face_API_LivenessSession_MatchConfidence = Azure_Face_API_LivenessSession.matchConfidence;
        let Azure_Face_API_LivenessSession_MatchDecision = Azure_Face_API_LivenessSession.matchDecision;

        return res.status(200).json({ Azure_Face_API_LivenessSession_LivenessDecision, Azure_Face_API_LivenessSession_MatchConfidence, Azure_Face_API_LivenessSession_MatchDecision });
    }

    async function refresh(req, res) {
        const platformRowIndex = res.locals.platformRowIndex;

        let BD_PlataformaResponse;
        try { BD_PlataformaResponse = await retry(() => microsoftGraph.readPlatformRows()); }
        catch (err) { return res.status(500).json({ error: learningPlatformErrorValues.platformDataReadFailure }); }
        const BD_Plataforma = microsoftGraph.extractRows(BD_PlataformaResponse);
        const LinhaVerificada = microsoftGraph.extractRowCells(BD_Plataforma[platformRowIndex]);

        let Usuário_NomeCompleto = LinhaVerificada[0];
        let Usuário_PrimeiroNome = LinhaVerificada[1];
        let Usuário_Email = LinhaVerificada[2];
        let Usuário_PrazoAcesso = ConverteData(LinhaVerificada[6]);
        let Usuário_Status_Login = LinhaVerificada[7];
        let Usuário_Formação_NúmeroTópicosConcluídos = LinhaVerificada[8];
        let Usuário_Formação_NotaMódulo1 = LinhaVerificada[10];
        let Usuário_Formação_NotaMódulo2 = LinhaVerificada[11];
        let Usuário_Formação_NotaMódulo3 = LinhaVerificada[12];
        let Usuário_Formação_NotaMódulo4 = LinhaVerificada[13];
        let Usuário_Formação_NotaMódulo5 = LinhaVerificada[14];
        let Usuário_Formação_NotaMódulo6 = LinhaVerificada[15];
        let Usuário_Formação_NotaMódulo7 = LinhaVerificada[16];
        let Usuário_Formação_NotaMódulo8 = LinhaVerificada[17];
        let Usuário_Formação_NotaMódulo9 = LinhaVerificada[18];
        let Usuário_Formação_NotaMódulo10 = LinhaVerificada[19];
        let Usuário_Formação_NotaAcumulado = LinhaVerificada[20];
        let Usuário_Formação_CertificadoID = LinhaVerificada[21];

        return res.status(200).json({ Usuário_NomeCompleto, Usuário_PrimeiroNome, Usuário_Email, Usuário_PrazoAcesso, Usuário_Status_Login, Usuário_Formação_NúmeroTópicosConcluídos, Usuário_Formação_NotaMódulo1, Usuário_Formação_NotaMódulo2, Usuário_Formação_NotaMódulo3, Usuário_Formação_NotaMódulo4, Usuário_Formação_NotaMódulo5, Usuário_Formação_NotaMódulo6, Usuário_Formação_NotaMódulo7, Usuário_Formação_NotaMódulo8, Usuário_Formação_NotaMódulo9, Usuário_Formação_NotaMódulo10, Usuário_Formação_NotaAcumulado, Usuário_Formação_CertificadoID });
    }

    async function updateProgress(req, res) {
        const platformRowIndex = res.locals.platformRowIndex;
        let { TipoAtualização, NúmeroTópicosConcluídos, NúmeroMódulo, NotaTeste } = req.body;

        let DadosaInserir = new Array(22).fill(null);
        DadosaInserir[8] = NúmeroTópicosConcluídos;

        if (TipoAtualização === 'NúmeroTópicosConcluídos-e-NotaTeste') DadosaInserir[NúmeroMódulo + 9] = NotaTeste;

        try { await retry(() => microsoftGraph.updatePlatformRow(platformRowIndex, DadosaInserir)); }
        catch (err) { return res.status(500).json({ error: learningPlatformErrorValues.platformDataWriteFailure }); }

        return res.status(200).json({});
    }

    async function processFeedback(req, res) {
        const platformRowIndex = res.locals.platformRowIndex;
        let { NúmeroTópicosConcluídos, Usuário_NomeCompleto, Usuário_Email, Feedback_DataPreenchimento, NúmeroMódulo, Feedback_TamanhoMódulo, Feedback_QualidadeConteúdo, Feedback_QualidadePlataforma, Feedback_QualidadeMateriaisImpressos, Feedback_Comentários } = req.body;

        try { await retry(() => microsoftGraph.updatePlatformRow(platformRowIndex, [null, null, null, null, null, null, null, null, NúmeroTópicosConcluídos, null, null, null, null, null, null, null, null, null, null, null, null, null])); }
        catch (err) { return res.status(500).json({ error: learningPlatformErrorValues.platformDataWriteFailure }); }

        try { await retry(() => microsoftGraph.appendFeedbackRow([Usuário_NomeCompleto, Usuário_Email, Feedback_DataPreenchimento, NúmeroMódulo, Feedback_TamanhoMódulo, Feedback_QualidadeConteúdo, Feedback_QualidadePlataforma, Feedback_QualidadeMateriaisImpressos, Feedback_Comentários])); }
        catch (err) { return res.status(500).json({ error: learningPlatformErrorValues.feedbackAppendFailure }); }

        return res.status(200).json({});
    }

    async function getStatusReport(req, res) {
        let { linha_inicial, linha_final } = req.body;

        let BD_PlataformaResponse;
        try { BD_PlataformaResponse = await retry(() => microsoftGraph.readPlatformRows()); }
        catch (err) { return res.status(500).json({ error: learningPlatformErrorValues.platformDataReadFailure }); }
        const BD_Plataforma = microsoftGraph.extractRows(BD_PlataformaResponse);

        const Dados_Extraídos_BD_Plataforma = BD_Plataforma.slice(linha_inicial, linha_final + 1).map((row) => {
            const cells = microsoftGraph.extractRowCells(row);
            return [cells[0], cells[8], ...cells.slice(10, 22)];
        });

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
