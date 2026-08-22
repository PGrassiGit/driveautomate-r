# Política de segurança

Não abra uma issue pública com credenciais, tokens, IDs de pastas reais ou dados
de clientes. Use o **Private vulnerability reporting** da aba Security do GitHub;
o mantenedor deve habilitar esse recurso antes da publicação. Se ele estiver
indisponível, peça primeiro um canal privado sem anexar o dado sensível.

## Antes de publicar uma release

1. Execute `python scripts/check_no_secrets.py`.
2. Execute a suíte de testes em uma máquina limpa.
3. Confirme que `build/`, `dist/`, planilhas e JSONs OAuth não estão no commit.
4. Gere e publique o SHA-256 do executável.
5. Faça varredura antimalware e, quando possível, assine com Authenticode.
6. Revise as exigências atuais de verificação OAuth do Google.

Credenciais já expostas devem ser revogadas. Reescrever apenas o último commit
não remove o segredo do histórico.

## Modelo de ameaça resumido

- O Excel é entrada não confiável; IDs e caminhos são validados.
- Downloads são confinados à pasta escolhida.
- Links/junções existentes no caminho de destino são recusados.
- Arquivos binários são verificados por tamanho e checksum quando a API fornece.
- Arquivos parciais usam nomes próprios e substituição atômica.
- Tokens novos usam DPAPI no Windows e não trazem o e-mail no nome do arquivo.
- O aplicativo nunca aceita automaticamente downloads marcados como abusivos.
