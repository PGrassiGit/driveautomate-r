# Checklist de release

Use uma máquina ou runner limpo para cada release pública.

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
- [ ] Assinar com Authenticode quando houver certificado disponível.

Referências oficiais:

- https://doc.qt.io/qtforpython-6/
- https://doc.qt.io/qtforpython-6/licenses.html
- https://developers.google.com/identity/protocols/oauth2/production-readiness/policy-compliance
