'use strict';

function createCertificateValidationHandler({ microsoftGraph, retry }) {
    return async function validateCertificate(req, res) {
        const Solicitante_CertificadoID = String(req.params.Solicitante_CertificadoID || '').trim().toUpperCase();

        let BD_PlataformaResponse;
        try { BD_PlataformaResponse = await retry(() => microsoftGraph.readPlatformRows()); }
        catch (err) { return res.status(500).json({ error: 'Erro_001' }); }
        const BD_Plataforma = microsoftGraph.extractRows(BD_PlataformaResponse);

        const Linha = Solicitante_CertificadoID
            ? BD_Plataforma.find((row) => {
                const cells = microsoftGraph.extractRowCells(row);
                return String(cells[21] == null ? '' : cells[21]).trim().toUpperCase() === Solicitante_CertificadoID;
            })
            : undefined;

        if (!Linha) return res.status(200).json({ Certificado_Válido: false });

        const cells = microsoftGraph.extractRowCells(Linha);
        const Acumulado_Bruto = Number(cells[20]);
        const Acumulado_Percentual = !isFinite(Acumulado_Bruto) ? 0 : (Acumulado_Bruto <= 1 ? Acumulado_Bruto * 100 : Acumulado_Bruto);

        if (Acumulado_Percentual < 70) return res.status(200).json({ Certificado_Válido: false });

        return res.status(200).json({
            Certificado_Válido: true,
            Titular_NomeCompleto: cells[0],
            Acumulado_Percentual: Math.round(Acumulado_Percentual),
            Certificado_ID: String(cells[21]).trim().toUpperCase(),
        });
    };
}

module.exports = { createCertificateValidationHandler };
