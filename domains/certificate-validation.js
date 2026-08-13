'use strict';

const PLATFORM_ROWS_PATH = '/users/a8f570ff-a292-4b2f-a1e4-629ccd7a26be/drive/items/01OSXVECSBYCZNYGEWFFDLEOZ36WI2PDWO/workbook/worksheets/{00000000-0001-0000-0000-000000000000}/tables/{7C4EBF15-124A-4107-9867-F83E9C664B31}/rows';

function createCertificateValidationHandler({ graphClient, retry }) {
    return async function validateCertificate(req, res) {
        const Solicitante_CertificadoID = String(req.params.Solicitante_CertificadoID || '').trim().toUpperCase();

        let BD_Plataforma;
        try { BD_Plataforma = await retry(() => graphClient.api(PLATFORM_ROWS_PATH).get()); }
        catch (err) { return res.status(500).json({ error: 'Erro_001' }); }

        const Linha = Solicitante_CertificadoID
            ? BD_Plataforma.value.find((row) => String(row.values[0][21] == null ? '' : row.values[0][21]).trim().toUpperCase() === Solicitante_CertificadoID)
            : undefined;

        if (!Linha) return res.status(200).json({ Certificado_Válido: false });

        const Acumulado_Bruto = Number(Linha.values[0][20]);
        const Acumulado_Percentual = !isFinite(Acumulado_Bruto) ? 0 : (Acumulado_Bruto <= 1 ? Acumulado_Bruto * 100 : Acumulado_Bruto);

        if (Acumulado_Percentual < 70) return res.status(200).json({ Certificado_Válido: false });

        return res.status(200).json({
            Certificado_Válido: true,
            Titular_NomeCompleto: Linha.values[0][0],
            Acumulado_Percentual: Math.round(Acumulado_Percentual),
            Certificado_ID: String(Linha.values[0][21]).trim().toUpperCase(),
        });
    };
}

module.exports = { createCertificateValidationHandler };
