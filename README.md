# DriveAutomate

Aplicativo desktop para inventariar pastas do Google Drive em Excel e baixar
somente os arquivos escolhidos pelo usuário. Funciona com **Meu Drive**,
**Compartilhados comigo** e **Drives compartilhados**.

> Status: beta. Teste com uma pasta pequena antes de processar um acervo grande.

## Recursos

- Interface em PySide6 voltada a usuários não técnicos.
- Gerenciamento de várias contas Google no mesmo computador.
- Tokens protegidos pelo DPAPI do Windows e arquivos de conta sem e-mail no nome.
- Detecção automática da origem da pasta.
- Inventário paginado e gravado em streaming, adequado a árvores grandes.
- Excel com links, metadados, aba de instruções e divisão automática por abas.
- Coluna segura `Baixar?`, configurada como `Não` por padrão.
- Download seletivo a partir do Excel depois que o usuário o editar.
- Download retomável para arquivos binários usando HTTP Range.
- Conversão automática de Docs, Sheets, Slides, Drawings, Forms, Apps Script,
  Sites, Vids e Jamboard para formatos locais compatíveis.
- Operações de download de longa duração para arquivos nativos grandes, sem o
  limite de 10 MB do método de exportação legado.
- Cancelamento cooperativo, arquivos parciais e relatório CSV de resultados.
- Proteções para path traversal, nomes reservados do Windows e colisões.
- Arquivos locais nunca são sobrescritos silenciosamente: o padrão é pular e a
  alternativa cria um novo nome.
- Nenhuma telemetria e nenhuma credencial incluída no repositório ou executável.

## Fluxo do usuário

### Primeira configuração

1. O responsável técnico cria um OAuth Client ID do tipo **Aplicativo para
   computador** no Google Cloud e ativa a Google Drive API.
2. No DriveAutomate, clica em **Configurar OAuth** e seleciona o JSON baixado.
3. Clica em **Adicionar conta** e conclui a autorização no navegador.

O arquivo [`credentials.example.json`](credentials.example.json) é apenas um
modelo público. Copie-o para um arquivo local, substitua os placeholders pelo
JSON emitido pelo Google Cloud e nunca publique o arquivo preenchido.

O JSON OAuth e os tokens ficam somente no perfil local do Windows:

```text
%APPDATA%\DriveAutomate
```

Nenhuma credencial é gravada no diretório do programa.
Os tokens são cifrados pelo DPAPI para o usuário atual do Windows. Tokens JSON
de versões anteriores são migrados para o formato protegido quando usados; por
segurança, reconecte a conta se um arquivo legado tiver sido copiado ou exposto.

### Inventário

1. Selecione a conta com acesso à pasta.
2. Cole o ID ou link da pasta.
3. Use **Testar acesso** para confirmar a origem detectada.
4. Escolha o destino e clique em **Gerar inventário Excel**.

O tamanho em TB não é baixado nessa etapa: somente metadados são consultados. O
tempo depende principalmente da quantidade de pastas e arquivos.

### Download seletivo

1. Abra o Excel gerado.
2. Na coluna amarela `Baixar?`, marque `Sim` somente nas linhas desejadas.
3. Opcionalmente escolha `PDF`, `DOCX`, `XLSX` ou `PPTX` em
   `Formato de exportação`.
4. Salve e feche o Excel.
5. Na aba **Downloads**, selecione a planilha e a pasta de destino.
6. Clique em **Analisar planilha** e depois em **Baixar selecionados**.

Também existe o modo “Todos os arquivos restantes”, útil quando as linhas
desnecessárias foram apagadas. Esse modo sempre exige confirmação. Filtros
visuais do Excel não removem nem selecionam linhas.

## OAuth e publicação pública

O download exige o escopo:

```text
https://www.googleapis.com/auth/drive.readonly
```

Esse escopo permite somente leitura, mas é classificado pelo Google como
restrito. Um aplicativo destinado a usuários externos pode precisar de
verificação OAuth e, conforme o modo de distribuição/armazenamento de dados, de
avaliação de segurança. Consulte as regras atuais antes de publicar uma release.

Tokens criados por versões antigas com `drive.metadata.readonly` não permitem
download. O DriveAutomate detecta o escopo antigo e solicita nova autorização.

Nunca faça commit de:

- JSON OAuth;
- tokens ou refresh tokens;
- planilhas reais de inventário;
- executáveis compilados com credenciais;
- capturas contendo e-mails, IDs ou nomes de clientes.

Se uma credencial já foi publicada, removê-la do último commit não basta: revogue
o cliente no Google Cloud e limpe todo o histórico Git.

Há dois modelos de distribuição:

- **comunitário/BYO OAuth:** cada organização importa seu próprio JSON pelo botão
  `Configurar OAuth`; o repositório e o `.exe` permanecem neutros;
- **gerenciado:** o mantenedor usa um projeto Google Cloud dedicado, conclui a
  verificação exigida e entrega o JSON por um instalador ou canal administrativo
  separado — nunca por commit, issue ou release pública.

Assim, a etapa técnica acontece uma vez com o responsável; o uso diário (conta,
link, Excel e downloads) continua voltado a usuários leigos.

## Desenvolvimento

Requisitos: Windows, Python 3.11 ou superior e navegador padrão.

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements-dev.txt
$env:QT_QPA_PLATFORM = "offscreen"
python -m pytest
Remove-Item Env:QT_QPA_PLATFORM
python main.py
```

Modo CLI para inventário:

```powershell
python main.py --cli --credentials "C:\caminho\oauth.json" `
  --folder-id "ID_OU_LINK" --output "inventario.xlsx"
```

## Build do Windows

```powershell
.\build_exe.ps1
```

O script instala as dependências de desenvolvimento, procura segredos, executa
os testes, constrói o executável e imprime seu SHA-256. O resultado local fica em
`dist\DriveAutomate.exe`; `build/` e `dist/` não são versionados.

`requirements-lock.txt` registra o conjunto exato validado no build Windows;
`requirements.txt` e `requirements-dev.txt` mantêm faixas para desenvolvimento.

O executável público não contém OAuth. Cada organização fornece seu próprio JSON
na primeira configuração. Para releases comerciais, avalie assinatura
Authenticode para reduzir alertas do Windows SmartScreen.

Para publicar uma release no GitHub, crie e envie uma tag no formato `vX.Y.Z`.
O workflow `Windows release` executa os testes, compila o `.exe`, gera o SHA-256
e publica também o arquivo ZIP com `LICENSE` e `THIRD_PARTY_NOTICES.md`.

Distribua `LICENSE` e `THIRD_PARTY_NOTICES.md` junto ao `.exe` e conclua o
[checklist de release](docs/RELEASE_CHECKLIST.md), inclusive a revisão das
obrigações LGPLv3/GPLv3 ou da licença comercial escolhida para Qt/PySide6.

## Limitações conhecidas

- A preparação de arquivos nativos grandes acontece no Google e pode levar alguns
  minutos. Se uma execução for encerrada nessa etapa, a preparação será solicitada
  novamente na próxima tentativa.
- Tipos nativos sem formato de download compatível são registrados como não
  suportados.
- Pastas marcadas no Excel não iniciam recursão; selecione os arquivos filhos.
- O Excel contém nomes, IDs e possivelmente e-mails de proprietários. Trate-o como
  informação sensível.
- Alterações feitas no Drive durante uma execução longa podem produzir uma visão
  não atômica do acervo.
- Arquivos binários parciais são preservados para retomada; exportações Workspace
  reiniciam o arquivo atual após cancelamento.
- Para downloads muito grandes, use NTFS ou exFAT e confirme espaço livre no
  destino. FAT32 não aceita arquivos individuais acima de 4 GiB.

## Segurança e privacidade

Consulte [SECURITY.md](SECURITY.md) e [PRIVACY.md](PRIVACY.md). O código inclui uma
verificação local de segredos:

```powershell
python scripts/check_no_secrets.py
```

Mudanças por versão estão em [CHANGELOG.md](CHANGELOG.md).

## Licença

Código do DriveAutomate sob licença MIT. PySide6/Qt e as demais dependências
mantêm suas próprias licenças; consulte [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md).

Referências oficiais:

- [Download e exportação na Drive API](https://developers.google.com/workspace/drive/api/guides/manage-downloads)
- [Operações de download de longa duração](https://developers.google.com/workspace/drive/api/guides/long-running-operations)
- [Suporte a Drives compartilhados](https://developers.google.com/workspace/drive/api/guides/enable-shareddrives)
- [Escopos da Drive API](https://developers.google.com/workspace/drive/api/guides/api-specific-auth)
- [Qt for Python](https://doc.qt.io/qtforpython-6/)
