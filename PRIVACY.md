# Privacidade

O DriveAutomate é executado localmente e não possui telemetria própria.

## Dados acessados

Com autorização do usuário, o aplicativo pode ler metadados e conteúdo de
arquivos do Google Drive. Ele não cria, altera, move nem exclui itens no Drive.

## Dados armazenados localmente

Em `%APPDATA%\DriveAutomate` ficam:

- a configuração OAuth desktop importada pelo responsável;
- tokens das contas conectadas, protegidos no Windows pelo DPAPI do usuário atual.

Os nomes dos novos arquivos de conta usam um identificador derivado e não expõem
o e-mail. O e-mail e o nome de exibição ficam dentro do conteúdo protegido para
que a interface consiga identificar a conta. Arquivos JSON legados são aceitos
para migração e protegidos quando a conta volta a ser usada.

Planilhas e downloads são gravados apenas nos destinos escolhidos pelo usuário.
O inventário pode conter nomes, IDs, estrutura de pastas e e-mails de
proprietários; cabe ao usuário protegê-lo e eliminá-lo quando não for mais
necessário.

## Compartilhamento e telemetria

O DriveAutomate não envia planilhas, tokens ou estatísticas a servidores do
projeto. A comunicação externa necessária ocorre diretamente com os serviços de
autenticação e com a Google Drive API.

## Remoção

O botão **Remover conta** elimina o token local, mas não revoga a autorização no
Google. A revogação completa pode ser feita na página de segurança da conta
Google. Desinstalar o aplicativo não apaga automaticamente planilhas ou arquivos
baixados em outros diretórios.
