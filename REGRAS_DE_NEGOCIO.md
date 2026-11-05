# Regras de Negócio — Sistema Integrado de Escalas & Cronograma

**Versão Atual:** 0.2  
**Repositório / Aplicação:** `gantt_disp-main`  
**Última Atualização:** Novembro/2026  

Este documento registra e detalha todas as regras de negócio, premissas de cálculo, integrações e fluxos operacionais implementados no sistema. Ele serve como especificação oficial para manutenção e evolução contínua do projeto.

---

## 1. Visão Geral e Arquitetura

O sistema une o acompanhamento macro de disponibilidade da equipe (Cronograma Gantt via Plotly/Pandas) com a gestão operacional de escalas contínuas matriciais das plataformas offshore (**P-80**, **P-82** e **P-83**).

### 1.1 Camada de Dados
- **Banco de Dados Local:** SQLite persistente em `escala_p80.db` (tabelas `p80_membros`, `p80_escalas`, `p80_config_vagas`, `p80_sugestoes` e `p80_audit_log`).
- **Fontes Externas (Cronograma):** Planilhas Excel sincronizadas via Google Drive ou armazenamento local (`Planejamento Geral.xlsx`, `Estaleiro`, `Férias`).
- **Sincronização:** Os colaboradores cadastrados nas planilhas são automaticamente sincronizados com o banco de dados das plataformas para manter a equipe atualizada.

---

## 2. Controle de Acesso e Perfis

O sistema opera com separação estrita de papéis:

1. **Modo Visualizador (Padrão):**
   - Acesso somente leitura a todas as abas e matrizes de 6 meses.
   - Navegação temporal livre (-6m, -1m, hoje, +1m, +6m).
   - Filtros de pesquisa em tempo real por colaborador ou disciplina.
   - Sem possibilidade de alteração acidental ou disparo de comandos de escrita.

2. **Modo Editor:**
   - Ativação mediante senha administrativa: `"admin"`.
   - Permite interação direta com a grade (clique em qualquer dia preenchido ou vazio para abrir o modal de edição).
   - Acesso ao painel administrativo expansível para lançamentos avulsos e cíclicos, cadastro/desativação de membros e limpeza controlada de escalas.
   - Configuração de metas de vagas mínimas a bordo por plataforma.
   - Opção explícita de logout ("🔒 Sair da Edição") que retorna com segurança ao modo visualizador.

---

## 3. Estrutura das Abas e Apresentação da Equipe

O aplicativo organiza a gestão em abas dedicadas:
- **Cronograma:** Visão Gantt temporal interativa com destaque de disponibilidade e conflitos globais.
- **Escala P-80:** Matriz contínua de 6 meses da unidade P-80.
- **Escala P-82:** Matriz contínua de 6 meses da unidade P-82.
- **Escala P-83:** Matriz contínua de 6 meses da unidade P-83.
- **Conflitos:** Tabela analítica de sobreposições de atividades futuras.
- **Sugestões:** Central colaborativa de sugestões com triagem por prioridade e acompanhamento de implementações.
- **Log de Edições:** Registro de auditoria cronológico e imutável de todas as inserções, edições, exclusões e permutas.
- **Equipe:** Cadastro geral com status instantâneo.
- **Exportar:** Exportação de relatórios em CSV.

### 3.1 Exibição de Linhas por Plataforma (Nativos x Externos)
- Cada aba de plataforma exibe, por padrão, **todos os membros ativos da sua respectiva equipe base**.
- As linhas são agrupadas pelas 4 disciplinas padrão:
  1. `Mecânica`
  2. `Instrumentação`
  3. `Elétrica`
  4. `Engenharia`
- **Membros Externos:** Se um colaborador pertencente a outra unidade ou cadastrado como "EXTERNO" for programado ou escalado naquela plataforma na janela visualizada, sua linha passa a ser exibida automaticamente na escala da unidade com uma etiqueta identificadora `Externo (<Origem>)`.

### 3.2 Exclusão Estrita de Supervisores e Coordenadores
- Colaboradores com função de **Supervisor** ou **Coordenador** (liderança) **não são puxados para as escalas das plataformas (P-80, P-82 e P-83)**.
- Esses cargos permanecem no acompanhamento de disponibilidade macro e liderança (abas Cronograma e Equipe), sem ocupar linhas operacionais nas grades offshore.

---

## 4. Regras de Regime de Folga: 14x21 (1,5x) vs 14x14 (1,0x)

Para atender colaboradores sob diferentes regimes operacionais (incluindo FSE e especialistas com contratos diferenciados), o sistema suporta a parametrização do regime de folga por colaborador e por período de embarque:

### 4.1 Proporções de Folga Suportadas
1. **Regime 14x21 (Fator 1,5x):**
   - Para cada 14 dias de embarque (`EMB`), confere direito a **21 dias de folga** (`FOL`):
     $$\text{Dias de Folga} = \lceil N \times 1.5 \rceil$$
   - Em escala contínua cíclica, cada ciclo perfaz 35 dias (14 embarcado + 21 de folga).
   - É o regime padrão atribuído automaticamente a todos os colaboradores próprios importados das planilhas.

2. **Regime 14x14 (Fator 1,0x):**
   - Para cada 14 dias de embarque (`EMB`), confere direito a **14 dias de folga** (`FOL`):
     $$\text{Dias de Folga} = N \times 1.0$$
   - Em escala contínua cíclica, cada ciclo perfaz 28 dias (14 embarcado + 14 de folga).
   - Aplicável a regimes específicos (como FSEs ou prestadores sob contrato 14x14).

### 4.2 Persistência e Aplicação
- **Cadastro do Colaborador:** O regime de folga base de cada colaborador é armazenado na tabela `p80_membros` (coluna `regime_folga`).
- **Lançamento de Período (Avulso ou Cíclico):** O formulário e o modal de edição apresentam o seletor `⚖️ Projeção de Regime de Folga (14x21 ou 14x14)`. Ao salvar, o sistema calcula as folgas subsequentes rigorosamente na proporção escolhida e atualiza a preferência cadastral do colaborador.
- **Redimensionamento na Edição:** Se um período for encurtado ou estendido, as folgas são recalculadas dinamicamente com base no regime selecionado.

---

## 5. Alocação Cíclica e Auto-Permuta de Folga

### 5.1 Alocação Cíclica Contínua
- Ao lançar um período de embarque (`EMB`), o editor pode ativar a opção de escala cíclica informando a quantidade de embarques ($K$, padrão 7 ciclos).
- O sistema projeta $K$ ciclos sequenciais ininterruptos de embarque e folga, respeitando o regime do colaborador (35 dias por ciclo em 14x21; 28 dias por ciclo em 14x14).

### 5.2 Auto-Permuta de Folga
- Ao clicar em qualquer período de folga (`FOL`), o modal ativa a interface de **Auto-Permuta**:
  1. Identifica o intervalo de folga original $[D_{\text{orig\_ini}}, D_{\text{orig\_fim}}]$ e o total de dias contínuos.
  2. O usuário escolhe a nova data de início da folga ($D_{\text{dest\_ini}}$).
  3. A data final é calculada automaticamente mantendo rigorosamente a mesma duração.
  4. **Trava de Segurança:** A nova folga deve ser posterior à data original e estar a no máximo 180 dias de distância do término original.

---

## 6. Integração com o Cronograma e Conflitos

### 6.1 Mapeamento de Eventos
- Eventos do cronograma geral (Excel) são projetados na grade sem gerar 1,5x de folga:
  - `Estaleiro` $\rightarrow$ Mapeado como `BASE` (Escritório / Base / Estaleiro).
  - `Férias` $\rightarrow$ Mapeado como `FÉR`.
  - `Folga` $\rightarrow$ Mapeado como `FOL`.

### 6.2 Detecção de Conflitos (`CONF`)
- Se em um mesmo dia o colaborador possuir escala própria na plataforma e houver atividade divergente no cronograma externo, a célula é sinalizada com **`CONF`** (vermelho `#dc2626`).
- **Hover Informativo:** Passar o mouse exibe a divergência exata (ex: `⚠️ CONFLITO em 15/11: Escala P-80 (EMB) vs Cronograma (Férias)`).
- **Relatório Global de Conflitos:** Filtra e exibe **apenas conflitos futuros** (com data de término maior que a data atual), organizados em sub-abas dedicadas por plataforma (`P-80`, `P-82`, `P-83`).

---

## 7. Dicionário de Status e Identidade Visual

| Código | Nome | Cor Hex | Descrição e Comportamento |
| :---: | :--- | :---: | :--- |
| **EMB** | Embarcado | `#0284c7` (Azul) | Embarque offshore. Gera folga proporcional (14x21 ou 14x14). |
| **FOL** | Folga | `#16a34a` (Verde) | Folga compensatória/escala. Permite Auto-Permuta (até 180 dias). |
| **PMB** | PMB | `#0891b2` (Ciano) | Pré-embarque e atividades preparatórias. |
| **FÉR** | Férias | `#d97706` (Laranja) | Férias regulamentares. Não gera folga compensatória. |
| **BASE** | Escritório / Base | `#7c3aed` (Roxo) | Estaleiro, base ou escritório. Não gera folga compensatória. |
| **TREIN** | Treinamento | `#059669` (Esmeralda) | Treinamentos e certificações. Não gera folga compensatória. |
| **LIC** | Licença | `#9f1239` (Vinho) | Licença médica, gala, paternidade, etc. Não gera folga compensatória. |
| **EMP** | Emprestado | `#4338ca` (Índigo) | Membro da base embarcado em outra plataforma. Não pontua na origem. |
| **CONF** | Conflito | `#dc2626` (Vermelho) | Sobreposição divergente entre escala própria e cronograma externo. |

### 7.1 Regras do Status de Empréstimo (`EMP`)
- Quando um colaborador nativo da plataforma $A$ é alocado com `EMB` na plataforma $B$:
  - Na plataforma $B$: ele é exibido como **Externo (`Ext`)**.
  - Na plataforma $A$ (origem nativa): ele é automaticamente exibido com o status **`EMP` (Emprestado)**.
  - **Soma da Turma a Bordo:** Células com status `EMP` **não pontuam** na soma diária de pessoas a bordo da plataforma de origem $A$ (pois o profissional está fisicamente a bordo na unidade $B$).
  - **Pop-up Informativo:** Ao clicar em qualquer célula `EMP` na unidade de origem, abre-se um modal dedicado informando o nome da plataforma de destino e o período, com botão de atalho para navegar diretamente à escala daquela unidade.

### 7.2 Regras do Status de Licença (`LIC`)
- Status dedicado para afastamentos formais (licença médica, licença paternidade/maternidade, etc.).
- Comporta-se de forma neutra em relação a projeções: ocupa os dias cadastrados **sem gerar dias de folga compensatória**.
- Disponível no modal de clique da grade e no painel administrativo do editor.

---

## 8. Janela Contínua de 6 Meses com Rolagem Horizontal

Para eliminar a necessidade de trocas constantes de mês e propiciar visão estratégica de médio prazo:

### 8.1 Visualização Matricial Contínua
- A grade exibe **6 meses consecutivos ininterruptos** a partir do mês/ano base selecionado (aproximadamente 182 dias).
- **Cabeçalho Duplo Fixo (Sticky):**
  - **Linha 1:** Banner mensal com nome do mês e ano em maiúsculas (`📅 JANEIRO 2027`), delimitando com precisão o intervalo de dias de cada mês.
  - **Linha 2:** Número do dia e sigla do dia da semana (ex: `01 SEG`).
- **Coluna de Profissionais Fixa (Sticky):** A coluna da esquerda (nomes, disciplinas e regimes) permanece travada durante a rolagem horizontal contínua.
- **Delimitação Visual Mensal:** O último dia de cada mês recebe uma borda vertical destacada em azul petróleo (`border-right: 2px solid #0284c7`), delimitando claramente a virada dos meses.

### 8.2 Controles de Navegação Temporal
- **`⏪ -6 Meses`:** Retrocede a janela em 6 meses.
- **`◀ -1 Mês`:** Retrocede a janela em 1 mês.
- **`📍 Mês Atual`:** Retorna imediatamente ao mês e ano vigentes.
- **`+1 Mês ▶`:** Avança a janela em 1 mês.
- **`+6 Meses ⏩`:** Avança a janela em 6 meses.
- **Seletores Diretos:** Selectbox de Mês Inicial e Ano Inicial para salto direto.

---

## 9. Formatação Condicional de Vagas Mínimas & Resumo Diário

Para monitorar e garantir o atendimento ao efetivo mínimo offshore:

### 9.1 Configuração de Vagas Mínimas
- No cabeçalho da escala, o botão com popover **`⚙️ Vagas Mín.`** permite configurar 4 parâmetros essenciais por plataforma:
  1. Mínimo de Mecânicos
  2. Mínimo de Instrumentistas
  3. Mínimo de Eletricistas
  4. Mínimo Total a Bordo
- Os valores são gravados na tabela `p80_config_vagas` do SQLite.

### 9.2 Resumo Diário da Turma a Bordo (4 Sublinhas)
Ao final da grade, a seção **Divisão de Turma a Bordo** apresenta 4 sublinhas diárias de monitoramento:
1. `📊 TOTAL A BORDO (SOMA)` (Meta: Mínimo Total)
2. `🔧 MECÂNICA A BORDO` (Meta: Mínimo Mecânica)
3. `⚡ ELÉTRICA A BORDO` (Meta: Mínimo Elétrica)
4. `🎛️ INSTRUMENTAÇÃO A BORDO` (Meta: Mínimo Instrumentação)

### 9.3 Sinalização Condicional de Alerta
- Caso em algum dia o efetivo a bordo de qualquer disciplina ou o total seja **inferior ao mínimo configurado** (e a meta seja $> 0$):
  - A célula do dia é destacada em **vermelho de alerta** (`background: #fee2e2`, `color: #b91c1c`, `border: 1px solid #fca5a5`).
  - É exibido o ícone de aviso: `⚠️ <qtd>`.
  - O tooltip detalha: `⚠️ Alerta: <qtd> a bordo em DD/MM/AAAA (Mínimo exigido: <meta>)`.

### 9.4 Identidade Visual dos Eletricistas na Turma a Bordo
- O algoritmo de classificação por disciplina identifica eletricistas independentemente de acentuação ou codificação de caracteres.
- **Código de Cores dos Balões Contínuos:**
  - **Mecânica:** Azul (`#0284c7`)
  - **Instrumentação:** Laranja (`#ea580c`)
  - **Elétrica:** Amarelo (`#eab308`)
  - **Engenharia / Outros:** Verde-azulado (`#0d9488`)

---

## 10. Aba de Sugestões de Implementação

Canal aberto para sugestões operacionais e melhorias técnicas no sistema:
1. **Submissão Livre:** Qualquer colaborador pode enviar sugestões informando seu nome ou optando pela caixa **"Anônimo"**.
2. **Triagem em Modo Edição:** Sugestões recém-enviadas ficam com status `Pendente` e só são visíveis pelo Editor. O editor define a prioridade (`Alta`, `Média`, `Baixa`) e aprova ou rejeita.
3. **Lista Pública por Prioridade:** Sugestões aprovadas constam publicamente na seção "Sugestões a Serem Implementadas", ordenadas estritamente por prioridade:
   $$\text{Alta} \longrightarrow \text{Média} \longrightarrow \text{Baixa}$$
4. **Conclusão:** O editor pode clicar em **"🚀 Marcar como Implementada"**, movendo o registro para a galeria pública de sugestões entregues.

---

## 11. Log de Auditoria & Histórico de Edições

A aba **Log de Edições** oferece rastreabilidade total, gravando de forma cronológica e imutável na tabela `p80_audit_log`:
- **Ações Registradas:** `INSERÇÃO`, `EDIÇÃO`, `EXCLUSÃO`, `AUTO-PERMUTA`, `LIMPEZA`, `EQUIPE` e `SUGESTÃO`.
- **Atributos:** Horário local exato de Brasília (`DD/MM/AAAA HH:MM:SS`), identificação do usuário editor, categoria da ação, plataforma, colaborador e detalhamento das datas/regimes.
- **Filtros e Exportação:** Filtros por plataforma, tipo de ação e busca textual, com botão de download completo em CSV.

---

## 12. Registro de Versões

- **v0.2 (Novembro/2026):**
  - **Janela de Tempo Contínua de 6 Meses:** Substituição da visualização mensal fechada por barra de rolagem horizontal contínua de 6 meses (~182 dias), com banner superior de mês/ano, bordas delimitadoras de virada de mês e botões rápidos de navegação temporal (`⏪ -6 Meses`, `◀ -1 Mês`, `📍 Mês Atual`, `+1 Mês ▶`, `+6 Meses ⏩`).
  - **Regimes de Folga 14x21 vs 14x14 (FSE):** Seleção de projeção de folga (1,5x ou 1,0x) tanto para embarques avulsos quanto cíclicos, persistida no cadastro do colaborador (`p80_membros.regime_folga`). Colaboradores próprios sincronizados recebem 14x21 por padrão.
  - **Status de Licença (`LIC`):** Novo status com chip Vinho (`#9f1239`), sem projeção de folga (comportamento análogo a Férias/Treinamento).
  - **Status de Empréstimo de Pessoal (`EMP`):** Novo status Índigo (`#4338ca`) exibido na plataforma de origem quando o colaborador estiver escalado em outra unidade offshore. Não pontua na Turma a Bordo da origem e conta com pop-up informativo com link para a escala da unidade de destino.
  - **Formatação Condicional de Vagas Mínimas:** Configuração de efetivo mínimo por plataforma (Mecânica, Instrumentação, Elétrica e Total a Bordo) via Popover `⚙️ Vagas Mín.` (`p80_config_vagas`). Sublinhas de monitoramento na Turma a Bordo destacando em vermelho (`#fee2e2`, `⚠️`) dias com contagem inferior à meta configurada.
  - **Correção da Cor dos Eletricistas:** Correção estrita na classificação de disciplinas, garantindo cor Amarela (`#eab308`) nos balões contínuos da Turma a Bordo.

- **v0.1 (Novembro/2026):**
  - Implementação das abas individuais de plataforma (**P-80**, **P-82**, **P-83**).
  - Exclusão estrita de supervisores e coordenadores das escalas operacionais de plataforma.
  - Regra proporcional estrita de 1,5x na criação e redimensionamento de embarques.
  - Alocação cíclica 14x21 e Auto-Permuta de folga ($\le 180$ dias).
  - Divisão de Turma a Bordo com balões contínuos por disciplina.
  - Integração do cronograma externo com mapeamento de Estaleiro como `BASE`.
  - Sinalização de conflito `CONF` (vermelho) com hover explicativo.
  - Relatório de conflitos futuros por plataforma (`P-80`, `P-82`, `P-83`).
  - Abas de Sugestões e Log de Edições com exportação CSV.
