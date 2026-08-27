# Changelog

## Unreleased

- CodeQL atualizado e configurado para não falhar enquanto o repositório privado
  não tiver GitHub Code Security; a análise inicia automaticamente ao torná-lo
  público.
- Workflows atualizados para as versões atuais das actions de checkout, Python e
  CodeQL.
- Dependências travadas e actions oficiais atualizadas para as versões validadas
  mais recentes antes da publicação do repositório.
- Smoke tests do executável agora aguardam os processos e validam seus códigos de
  saída, com timeout para impedir falsos resultados positivos no CI.
- Scanner de segredos ampliado para detectar também chaves de API do Google.
- Metadados de licença ajustados para builds com versões atuais do setuptools.
- Interface com tema escuro, campos dimensionados e abas roláveis para evitar
  compressão ou sobreposição em janelas menores.
- Opções de conflito de nomes simplificadas, com explicação acessível no
  controle de downloads.

## 3.0.0 — 2026-08-18

- Interface reescrita em PySide6 com workers Qt e cancelamento cooperativo.
- Detecção separada de Meu Drive, Compartilhados comigo e Drives compartilhados.
- Suporte a links com resource key e paginação para árvores grandes.
- Inventário Excel schema 2, em streaming, com seleção explícita para download.
- Downloads binários retomáveis, verificação de tamanho/checksum e proteção de
  caminhos Windows.
- Operações longas `files.download` para conteúdo Google Workspace e Vids.
- Gerenciamento de múltiplas contas e tokens protegidos pelo DPAPI do Windows.
- Remoção de credenciais embutidas; OAuth passa a ser configuração local.
- CI, CodeQL, Dependabot, build PyInstaller, documentação de privacidade,
  segurança, licenças e release.
