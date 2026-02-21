require('dotenv').config();
const fs = require('fs');

const token = process.env.AZURE_FACE_API_NPM_TEMPORARY_TOKEN || '';
if (!token) {
    console.log('Token not set — skipping .npmrc creation');
    process.exit(1); // fail loudly
}

const content = `registry=https://pkgs.dev.azure.com/msface/SDK/_packaging/AzureAIVision/npm/registry/

; begin auth token
//pkgs.dev.azure.com/msface/SDK/_packaging/AzureAIVision/npm/registry/:username=msface
//pkgs.dev.azure.com/msface/SDK/_packaging/AzureAIVision/npm/registry/:_password=${token}
//pkgs.dev.azure.com/msface/SDK/_packaging/AzureAIVision/npm/registry/:email=contato@machadogestao.com
//pkgs.dev.azure.com/msface/SDK/_packaging/AzureAIVision/npm/:username=msface
//pkgs.dev.azure.com/msface/SDK/_packaging/AzureAIVision/npm/:_password=${token}
//pkgs.dev.azure.com/msface/SDK/_packaging/AzureAIVision/npm/:email=contato@machadogestao.com
; end auth token`;

fs.writeFileSync('.npmrc', content);
console.log('.npmrc written');