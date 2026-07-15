# Monitore e atualize a versão do Face liveness client-side SDK sendo utilizada.

Versão sendo utilizada atualmente: 1.4.8.

1) Visite https://learn.microsoft.com/en-us/azure/ai-services/face/whats-new-face. Só se houver uma nova versão, prossiga aos próximos passos.

2) Abra https://github.com/Azure-Samples/azure-ai-vision-sdk/blob/main/samples/web/README.md para auxiliá-lo com eventuais consultas, com eventual suporte.

3) Realize o curl via **npmrc_password.http** e copie o token em Base64.

4) No terminal atrelado ao backend, armazene o token apenas para aquela sessão:

```powershell
$env:AZURE_AI_VISION_NPM_TOKEN_BASE64 = '<TOKEN_BASE64>'
```

O arquivo **.npmrc** utiliza essa variável nas duas configurações `_password`. O token não deve ser colado no arquivo nem versionado. Fechar o terminal remove a variável.

5) No mesmo terminal, rode **npm install @azure/ai-vision-face-ui@latest** e aguarde a instalação concluir.

6) Copie o arquivo **backend/node_modules/@azure/ai-vision-face-ui/FaceLivenessDetector.js** e cole em **sistemas/plataforma_v2/azure-ai-vision-face-ui**

7) Copie a pasta **backend/azure-ai-vision-face/ui-assets/facelivenessdetector-assets** e cole em **sistemas/plataforma_v2/azure-ai-vision-face-ui**

8) Substitua a imagem **sistemas/plataforma_v2/azure-ai-vision-face-ui/facelivenessdetector-assets/images/Brightness.svg** pela imagem **sistemas/plataforma_v2/login/img/Brightness.svg**. 

9) Atualize os seguintes valores no arquivo **sistemas/plataforma_v2/azure-ai-vision-face-ui/facelivenessdetector-assets/i18n/pt-BR/en.json**:

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"AZAIF_IncreaseBrightness": "Coloque o brilho da tela no máximo e afaste-se de janelas muito iluminadas."

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"AZAIF_IncreaseBrightnessHighestSetting": "A tela piscará algumas vezes para processar o FaceID."

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"AZAIF_IncreaseBrightnessTurnedUp": "Coloquei o brilho no máximo e me afastei de janelas muito iluminadas."

10) Delete a pasta **backend/node_modules** e o arquivo **backend/package-lock.json**.

11) Delete a linha **"@azure/ai-vision-face-ui": "^1.4.8",** no arquivo **backend/package.json**.

12) Abra um novo terminal atrelado ao backend e rode **npm install**.

13) Edite altere o URL_Base_Backend para o endereço local no frontend. Então teste localmente se o FaceID está funcionando corretamente.

14) Caso afirmativo, retorne o URL_Base_Backend ao endereço do Azure. Então deploye o backend e o frontend à produção.
