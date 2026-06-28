# CORREÇÕES PONTUAIS - NÃO ALTERAR FUNCIONALIDADES JÁ FUNCIONAIS

IMPORTANTE

O sistema passou recentemente por uma reorganização da navegação e diversas funcionalidades já voltaram a funcionar.

NÃO realizar refatorações.

NÃO reorganizar arquitetura.

NÃO alterar Dashboard, Histórico, Supabase, Backup, Restore, Importação, Pódios, Rankings ou qualquer outra funcionalidade que não esteja explicitamente listada abaixo.

NÃO criar melhorias extras.

NÃO alterar layout.

NÃO alterar nomes de funções existentes.

NÃO modificar regras de negócio.

O objetivo é corrigir APENAS os problemas descritos neste documento.

Antes de implementar qualquer correção:

1. Analise o código atual.
2. Identifique a causa real do problema.
3. Corrija apenas o ponto necessário.
4. Preserve todo o restante do sistema.

Ao final:

* Entregue o ARQUIVO PYTHON COMPLETO atualizado.
* Garantindo que os problemas abaixo foram corrigidos.
* Sem remover funcionalidades já existentes.

═══════════════════════════════════════════
BUG 1 - GERENCIAR USUÁRIOS QUEBRA O SISTEMA
═══════════════════════════════════════════

Ao acessar:

👥 Gerenciar Usuários

ocorre:

NameError

Traceback aponta para:

pagina_gerenciar_usuarios()

e para a chamada:

gerenciar_usuarios_supabase()

Objetivo:

* Verificar se a função existe.
* Verificar se foi renomeada.
* Verificar se a chamada está incorreta.
* Corrigir apenas a referência.

IMPORTANTE:

Não recriar a funcionalidade.

Não reimplementar gerenciamento de usuários.

Apenas localizar a função correta e corrigir a chamada.

═══════════════════════════════════════════
BUG 2 - HISTÓRICO EXIBE GRÁFICO DUPLICADO
═══════════════════════════════════════════

Na tela:

📈 Histórico

estão sendo exibidos dois gráficos.

O primeiro:

"Evolução Mensal"

aparece vazio.

O segundo:

"Evolução Mensal - Nome do Analista"

aparece corretamente.

Objetivo:

* Identificar por que o gráfico vazio está sendo renderizado.
* Remover a renderização de gráficos sem dados.
* Manter apenas gráficos válidos.

IMPORTANTE:

Não alterar o gráfico que já está funcionando.

Não alterar cálculos.

Não alterar métricas.

Apenas impedir a exibição de gráficos vazios.

═══════════════════════════════════════════
BUG 3 - PERÍODO ACEITA DIGITAÇÃO LIVRE
═══════════════════════════════════════════

Foi possível cadastrar:

"Janwiro 2026"

em vez de:

"Janeiro 2026"

Isso gera períodos inválidos e pode comprometer:

* Dashboard
* Histórico
* Comparações
* Controle de duplicidade

Objetivo:

Eliminar qualquer digitação manual do mês.

Substituir por seleção obrigatória.

Meses permitidos:

Janeiro
Fevereiro
Março
Abril
Maio
Junho
Julho
Agosto
Setembro
Outubro
Novembro
Dezembro

O ano deve continuar sendo selecionado separadamente.

O sistema deve montar automaticamente:

"Mês + Ano"

Exemplo:

Janeiro + 2026

Resultado:

Janeiro 2026

IMPORTANTE:

Não permitir texto livre para o mês.

═══════════════════════════════════════════
BUG 4 - DASHBOARD SEM SELETOR DE PERÍODO
═══════════════════════════════════════════

Atualmente a Dashboard trabalha apenas com o período carregado automaticamente.

Objetivo:

Adicionar um seletor de período na Dashboard.

Exemplo:

Período:
[ Janeiro 2026 ▼ ]

Ao trocar o período:

* Atualizar indicadores
* Atualizar gráficos
* Atualizar rankings
* Atualizar pódios
* Atualizar métricas

IMPORTANTE:

Não alterar os cálculos atuais.

Não alterar as consultas atuais.

Apenas permitir selecionar qual período será exibido.

═══════════════════════════════════════════
VALIDAÇÃO OBRIGATÓRIA ANTES DA ENTREGA
═══════════════════════════════════════════

Antes de finalizar:

1. Confirmar que Gerenciar Usuários abre normalmente.
2. Confirmar que o Histórico não exibe gráfico vazio.
3. Confirmar que não é possível digitar manualmente o nome do mês.
4. Confirmar que a Dashboard possui seletor de período.
5. Confirmar que Dashboard, Histórico, Importação, Gerenciar Períodos, Configurações e Supabase continuam funcionando.

═══════════════════════════════════════════
ENTREGA
═══════════════════════════════════════════

Entregar:

* Arquivo Python completo atualizado.
* Sem trechos omitidos.
* Sem "...".
* Sem "restante do código igual".

Antes do código, apresentar um resumo das alterações realizadas em cada um dos 4 bugs corrigidos.
