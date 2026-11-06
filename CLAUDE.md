# Compras Marco · Google Apps Script (web app)

## Regra obrigatória: versão a cada atualização

Toda alteração no programa muda a versão, no mesmo commit:

1. Subir `VERSAO` e `VERSAO_DATA` (dd/mm/aaaa), constantes no topo de `codemarco.gs` (fonte única).
   - correção/ajuste pequeno → 3º número (1.0.0 → 1.0.1)
   - funcionalidade nova ou regra alterada → 2º número (1.0.1 → 1.1.0)
   - mudança grande de estrutura → 1º número (1.1.0 → 2.0.0)
2. Acrescentar a linha no "Histórico de versões" de `NOTAS.md` (mais nova em cima): versão, data e o que mudou.
3. Dizer ao usuário, na resposta, qual é a nova versão.

A versão aparece numa faixa logo acima do programa — "Versão x.y.z · atualizado em dd/mm/aaaa" — em `indexmarco.html`, com os valores vindos de `doGet` em `codemarco.gs`. É assim que o usuário confere que a versão certa está rodando depois de publicar (se a faixa não mudou, a nova versão da implantação não foi publicada ou a página está em cache: Ctrl+F5). Não remover nem esconder a faixa; ela some na impressão/PDF.

### Primeira alteração (se o projeto ainda não tem a faixa)

Na primeira alteração, criar junto: `VERSAO = '1.0.0'` e `VERSAO_DATA` em `codemarco.gs`, repassadas ao HTML por `doGet`, a faixa acima do programa em `indexmarco.html` e o arquivo `NOTAS.md` com a tabela "Histórico de versões". Depois disso, seguir a regra normalmente. Vale só para o web app; scripts auxiliares (Python, firmware) ficam de fora, salvo pedido do usuário.
