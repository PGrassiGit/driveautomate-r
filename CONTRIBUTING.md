# Como contribuir

1. Abra uma issue sem dados pessoais para discutir mudanças grandes.
2. Crie um ambiente virtual e instale `requirements-dev.txt`.
3. Mantenha lógica de Drive/Excel separada da interface PySide6.
4. Adicione testes para todo comportamento novo.
5. Execute:

```powershell
python scripts/check_no_secrets.py
$env:QT_QPA_PLATFORM = "offscreen"
python -m pytest
```

Não inclua credenciais, tokens, planilhas reais, caminhos de perfil, binários ou
capturas de clientes. Commits devem ser pequenos, descritivos e revisáveis.

