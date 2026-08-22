# Checklist de release

Use uma máquina ou runner limpo para cada release pública.

- [ ] No GitHub, revisar descrição, topics, visibilidade, branch principal e
  habilitar **Private vulnerability reporting**.
- [ ] Escolher o modelo OAuth (BYO OAuth recomendado para o repositório público)
  e configurar a tela de consentimento, usuários de teste e publicação no Google
  Cloud sem commitar o JSON desktop.
- [ ] Revogar qualquer cliente OAuth que já tenha aparecido em código, binário,
  compactado ou histórico Git.
- [ ] Executar `python scripts/check_no_secrets.py` e toda a suíte de testes.
- [ ] Confirmar que a release não contém tokens, JSON OAuth, planilhas reais,
  IDs, e-mails, caminhos de perfil, `build/` ou caches.
- [ ] Revisar o histórico com uma ferramenta de detecção de segredos; quando
  necessário, reescrever o histórico e invalidar os segredos.
- [ ] Gerar o `.exe` em ambiente limpo, publicar o SHA-256 e realizar varredura
  antimalware.
- [ ] Distribuir `LICENSE` e `THIRD_PARTY_NOTICES.md` junto ao executável.
- [ ] Revisar e cumprir LGPLv3/GPLv3 ou usar a licença comercial apropriada para
  Qt/PySide6; confirmar também as licenças dos módulos Qt efetivamente empacotados.
- [ ] Publicar política de privacidade, canal privado de segurança e instruções de
  exclusão/revogação de dados.
- [ ] Concluir a verificação OAuth exigida pelo Google para o público e os escopos
  escolhidos.
- [ ] Substituir o marcador de link em `docs/LINKEDIN_POST.md` pelo URL público e
  usar somente capturas com dados fictícios.
- [ ] Assinar com Authenticode quando houver certificado disponível.

Referências oficiais:

- https://doc.qt.io/qtforpython-6/
- https://doc.qt.io/qtforpython-6/licenses.html
- https://developers.google.com/identity/protocols/oauth2/production-readiness/policy-compliance
