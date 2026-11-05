# Regras de Negócio — Sistema Integrado de Escalas & Cronograma

**Versão Atual:** 0.1  
**Repositório / Aplicação:** `gantt_disp-main`  
**Última Atualização:** Novembro/2026  

Este documento registra e detalha todas as regras de negócio, premissas de cálculo, integrações e fluxos operacionais implementados no sistema. Ele serve como especificação oficial para manutenção e evolução do projeto.

---

## 1. Visão Geral e Arquitetura

O sistema une o acompanhamento macro de disponibilidade da equipe (Cronograma Gantt via Plotly/Pandas) com a gestão operacional de escalas mensais matriciais das plataformas offshore (**P-80**, **P-82** e **P-83**).

### 1.1 Camada de Dados
- **Banco de Dados Local:** SQLite persistente em `escala_p80.db` (tabelas `p80_membros` e `p80_escalas`).
- **Fontes Externas (Cronograma):** Planilhas Excel sincronizadas via Google Drive ou armazenamento local (`Planejamento Geral.xlsx`, `Estaleiro`, `Férias`).
- **Sincronização:** Os colaboradores cadastrados nas planilhas são automaticamente sincronizados com o banco de dados das plataformas para manter a equipe atualizada.

---

## 2. Controle de Acesso e Perfis

O sistema opera com separação estrita de papéis:

1. **Modo Visualizador (Padrão):**
   - Acesso somente leitura a todas as abas e matrizes mensais.
   - Navegação por meses e anos livre.
   - Filtros de pesquisa em tempo real por colaborador ou disciplina.
   - Sem possibilidade de alteração ou disparo de comandos de escrita.

2. **Modo Editor:**
   - Ativação mediante senha administrativa: `"admin"`.
   - Permite interação direta com a grade (clique em qualquer dia preenchido ou vazio para abrir o modal de edição).
   - Acesso ao painel administrativo expansível para lançamentos em lote, cadastro/desativação de membros e limpeza controlada de escalas.
   - Opção explícita de logout ("🔒 Sair da Edição") que retorna com segurança ao modo visualizador.

---

## 3. Estrutura das Abas e Apresentação da Equipe

O aplicativo organiza a gestão em abas dedicadas:
- **Cronograma:** Visão Gantt temporal interativa com destaque de disponibilidade e conflitos globais.
- **Escala P-80:** Matriz mensal de escalas da unidade P-80.
- **Escala P-82:** Matriz mensal de escalas da unidade P-82.
- **Escala P-83:** Matriz mensal de escalas da unidade P-83.
- **Conflitos:** Tabela analítica de sobreposições de atividades.
- **Sugestões:** Central colaborativa de sugestões com fluxo de aprovação por prioridade e acompanhamento de implementações.
- **Log de Edições:** Registro de auditoria cronológico e imutável de todas as inserções, edições, exclusões e permutas.
- **Equipe:** Cadastro geral com status instantâneo.
- **Exportar:** Exportação de relatórios em CSV.

### 3.1 Exibição de Linhas por Plataforma (Nativos x Externos)
- Cada aba de plataforma exibe, por padrão, **todos os membros ativos da sua respectiva equipe** (ex: os 17 membros alocados àquela unidade).
- As linhas são agrupadas pelas 4 disciplinas padrão:
  1. `Mecânica`
  2. `Instrumentação`
  3. `Elétrica`
  4. `Engenharia`
- **Membros Externos:** Se um colaborador pertencente a outra unidade ou cadastrado como "EXTERNO" for programado ou escalado naquela plataforma no mês visualizado, sua linha passa a ser exibida automaticamente na escala mensal da unidade com uma etiqueta identificadora `Externo (<Origem>)`.

### 3.2 Exclusão Estrita de Supervisores e Coordenadores
- Colaboradores com função de **Supervisor** ou **Coordenador** (liderança) **não são puxados para as escalas das plataformas (P-80, P-82 e P-83)**.
- Esses cargos permanecem no acompanhamento de disponibilidade macro e liderança (abas Cronograma e Equipe), sem ocupar linhas operacionais nas grades offshore.

---

## 4. Regra de Negócio: Proporcionalidade de Folga (1,5x)

No regime offshore implementado, cada dia de embarque confere direito a folga compensatória calculada pela proporção de 1,5 dias.

### 4.1 Projeção Automática de Embarque
- Ao cadastrar um período de embarque (`EMB`) de $N$ dias consecutivos:
  $$\text{Dias de Folga} = \lceil N \times 1.5 \rceil$$
- Os dias de folga (`FOL`) são gerados e gravados **automaticamente** no banco de dados na sequência imediata do embarque (a partir do dia $D_{\text{fim}} + 1$).
- *Exemplo:* Embarque de 10 dias (01/11 a 10/11) projeta automaticamente 15 dias de folga (11/11 a 25/11).

### 4.2 Ajuste Proporcional Dinâmico na Edição
- Caso um período de embarque seja alterado, o sistema ajusta proporcionalmente o saldo de folgas correspondente:
  - *Exemplo Prático:* Se um embarque de 10 dias (com 15 folgas projetadas) for reduzido para 6 dias (01/11 a 06/11):
    1. O novo período de folga passa a ser de $\lceil 6 \times 1.5 \rceil = 9$ dias.
    2. A nova folga inicia imediatamente após o novo término: de 07/11 a 15/11.
    3. **Limpeza Automática:** Todas as folgas anteriores excedentes (do dia 16/11 ao dia 25/11) são automaticamente removidas do banco de dados, evitando folgas órfãs ou duplicadas.

### 4.3 Alocação em Escala Cíclica (Regime 14x21)
- Ao lançar um novo embarque (`EMB`), o sistema disponibiliza a opção **"Alocar em escala contínua (regime 14x21)"**.
- O usuário define a quantidade de embarques completos a projetar (número inteiro $K$, ex: 7).
- A partir da data inicial selecionada ($D_0$), o sistema gera $K$ ciclos consecutivos:
  - **Embarque do Ciclo $i$:** 14 dias consecutivos com status `EMB`.
  - **Folga do Ciclo $i$:** 21 dias consecutivos com status `FOL` ($14 \times 1.5 = 21$).
  - **Início do Próximo Ciclo:** Exatamente $35\text{ dias}$ após o início do ciclo anterior ($14 + 21$).
- *Exemplo com 7 ciclos:* São criados 7 períodos de 14 dias de embarque (98 dias de `EMB`) intercalados por 7 períodos de 21 dias de folga (147 dias de `FOL`), cobrindo 245 dias contínuos de escala no banco de dados.

---

## 5. Regra de Negócio: Janela de "Auto-Permuta" de Folga

Quando um colaborador ou gestor clica sobre um período de folga (`FOL`) no modo edição, o sistema aciona uma janela dedicada de **Auto-Permuta**:

### 5.1 Parâmetros e Premissas
1. **Intervalo de Origem:** O usuário seleciona o intervalo de folga que deseja ceder/trocar (`[Data Início Origem, Data Fim Origem]`). O sistema contabiliza a quantidade de dias contínuos ($Q$).
2. **Nova Data de Início (Destino):** O usuário informa quando deseja iniciar o novo gozo da folga.
3. **Cálculo Automático do Término:** O sistema calcula automaticamente:
   $$\text{Data Fim Destino} = \text{Data Início Destino} + (Q - 1)\text{ dias}$$
   Garantindo rigorosamente a **mesma quantidade de dias contínuos**.
4. **Limite Temporal Estrito:**
   - A nova data de início deve ser posterior à data de origem.
   - A nova data de início não pode exceder **180 dias** de distância da data final de origem:
     $$\Delta t \le 180\text{ dias}$$
   - Se o usuário tentar permutar para além de 180 dias, a operação é bloqueada com alerta explicativo.
5. **Efetivação:** Ao confirmar a Auto-Permuta, as folgas originais são apagadas do banco e o novo bloco contínuo de folga é gravado nas datas de destino com observação detalhada da permuta.

---

## 6. Regra de Negócio: Integração com Cronograma & Conflitos (`CONF`)

A escala mensal matricial puxa os dados de eventos externos registrados no Cronograma Geral (`combined_df`), enriquecendo a matriz da plataforma com a rotina externa da equipe.

### 6.1 Mapeamento de Atividades Importadas
| Atividade no Cronograma | Status na Escala | Observação Automática | Regra de Folga 1,5x |
| :--- | :---: | :--- | :---: |
| **Estaleiro** | `BASE` | `Estaleiro (Cronograma)` | **Não calcula 1,5x** |
| **Férias** | `FÉR` | `Férias (Cronograma)` | **Não calcula 1,5x** |
| **Folga** | `FOL` | `Folga (Cronograma)` | **Não calcula 1,5x** |

> **Nota:** Eventos importados do cronograma não geram projeção automática de 1,5x de folga; eles são plotados pontualmente conforme lançado na origem.

### 6.2 Detecção de Conflitos e Sinalização (`CONF`)
- Se em um mesmo dia o colaborador tiver uma escala programada na plataforma (ex: `EMB`) e o cronograma registrar um evento divergente (ex: `Férias`, `Estaleiro` ou `Folga` externa):
  1. A célula é marcada com o status **`CONF`** (Conflito).
  2. A célula é renderizada em **vermelho vibrante** (`#dc2626`) com o texto em branco `CONF` e classe visual `.item-conf`.
  3. **Aviso no Hover (Tooltip):** Ao passar o cursor do mouse sobre qualquer célula `CONF`, o sistema exibe imediatamente a discriminação exata da divergência e os limites reais da escala do colaborador (ex: `⚠️ CONFLITO em 15/11: Escala P-80 (EMB) vs Cronograma (Férias) | Escala: 01/11 a 14/11 (EMB) (Clique para editar/excluir)`).
  4. A legenda superior exibe o chip obrigatório: **`CONF (Conflito)`**.
  5. O KPI do mês contabiliza o número de conflitos em um card dedicado.
  6. Um painel expansível logo abaixo dos KPIs lista analiticamente cada conflito detectado (nome, data, atividade na escala vs. atividade no cronograma) para rápida conciliação.
  7. **Preservação dos Limites Reais do Período no Modal:** Clicar sobre qualquer dia com `CONF` abre o modal de edição contendo o **período integral da escala original** (ex: início em 01/11 e término em 14/11) e condição `EMB`, acompanhado da caixa informativa de conflito em vermelho. Isso permite ao usuário gerenciar, ajustar ou excluir o embarque completo e suas folgas associadas sem reduzi-lo a um único dia.
  8. **Respeito a Extensões Manuais de Data:** Se o usuário alterar ou estender manualmente as datas de início ou término no formulário do modal e solicitar exclusão ou salvamento, o sistema respeita a extensão abrangente (`min` das datas iniciais e `max` das datas finais), garantindo a exclusão ou atualização de todo o intervalo pretendido.
  9. **Continuidade na Turma a Bordo:** Dias marcados com `CONF` cujo status da escala seja `EMB` continuam sendo computados na soma diária a bordo e preservam o balão contínuo do colaborador sem fragmentações.

### 6.3 Relatório Global de Conflitos (Filtragem Temporal e Divisão por Plataforma)
1. **Filtragem Estrita por Conflitos Futuros:**
   - O relatório de conflitos e a métrica de conflitos no topo do dashboard consideram **exclusivamente conflitos futuros**.
   - Critério temporal: a data de término da sobreposição entre os eventos divergentes deve ser estritamente maior que o dia atual:
     $$\text{Término Conflito} > \text{Data Atual (Hoje)}$$
   - Conflitos cujas sobreposições já se encerraram no passado são descartados das contagens e das tabelas analíticas do relatório.
2. **Métricas (KPIs):**
   - O card de métricas superior exibe `Conflitos (Futuros)`, computando unicamente o total de divergências que afetam o planejamento atual e futuro.
3. **Divisão por Plataforma (`P-80`, `P-82`, `P-83`):**
   - Na aba **Conflitos**, os registros futuros são organizados em sub-abas dedicadas por ativo:
     - **P-80 (contagem):** exibe os conflitos dos colaboradores alocados na P-80.
     - **P-82 (contagem):** exibe os conflitos dos colaboradores alocados na P-82.
     - **P-83 (contagem):** exibe os conflitos dos colaboradores alocados na P-83.
     - **Outros:** caso existam divergências associadas a outros projetos ou bases, uma sub-aba complementar é apresentada automaticamente.
   - Cada sub-aba conta com cartões de contagem e tabela analítica estruturada com as colunas: `Nome`, `Atividade A`, `Período A`, `Atividade B`, `Período B`, `Período Conflito` e `Sobreposição (dias)`.
4. **Exportação:**
   - Na aba **Exportar**, o botão **📥 Conflitos Futuros (CSV)** disponibiliza o download de todos os conflitos futuros detectados, contendo a coluna `Projeto` para auditoria externa.

---

## 7. Dicionário de Status e Identidade Visual

| Código | Nome | Cor Hex | Descrição |
| :---: | :--- | :---: | :--- |
| **EMB** | Embarcado | `#0284c7` (Azul) | Período de embarque ativo offshore (Gera folga 1,5x) |
| **FOL** | Folga | `#16a34a` (Verde) | Folga de escala / compensatória (Suporta Auto-Permuta) |
| **PMB** | PMB | `#0891b2` (Ciano) | Pré-embarque e atividades preparatórias |
| **FÉR** | Férias | `#d97706` (Laranja) | Férias regulamentares |
| **BASE** | Escritório / Base | `#7c3aed` (Roxo) | Trabalho em base, oficina ou estaleiro |
| **TREIN** | Treinamento | `#059669` (Esmeralda) | Capacitação e certificações |
| **CONF** | Conflito | `#dc2626` (Vermelho) | Sobreposição divergente entre escala e cronograma |

### 7.1 Divisão de Turma a Bordo (Soma Dia a Dia e Balões Contínuos)
Ao final das divisões de disciplinas da grade (Mecânica, Instrumentação, Elétrica, Engenharia), é apresentada a seção consolidada **Divisão de Turma a Bordo**:
1. **Linha de Soma Dia a Dia (`TOTAL A BORDO`):**
   - Apresenta no cabeçalho a soma dia a dia de quantos colaboradores estão embarcados (`EMB`) em cada dia do mês.
2. **Linhas de Colaboradores com Balão Contínuo:**
   - Cada colaborador com embarque no mês é representado por um **balão contínuo** que se estende por todo o período de embarque naquele mês, exibindo o nome e a duração.
3. **Código de Cores Obrigatório por Disciplina:**
   - **1 - Azul (`#0284c7`):** Mecânico
   - **2 - Laranja (`#ea580c`):** Instrumentação
   - **3 - Amarelo (`#eab308`):** Elétrica
   - **Outras/Engenharia (`#0d9488`):** Verde-azulado

---

## 8. Guia de Operações da Interface

1. **Lançamento de Escala Rápido:**
   - Em modo edição, clique em qualquer célula vazia (`—`) para abrir o modal de lançamento já pré-preenchido com a data e colaborador. É possível ativar a opção "Alocar em escala contínua (regime 14x21)" para gerar N ciclos automáticos.
2. **Edição Confiável de Períodos:**
   - Em modo edição, clique diretamente sobre o período colorido ou balão da turma a bordo para abrir o modal de edição.
   - O modal permite alterar data de início, data de término, status (`EMB`, `FOL`, `PMB`, `FÉR`, `BASE`, `TREIN`) e observações.
   - Ao clicar em **"💾 Salvar Alterações"**, o formulário é submetido de forma segura através do bridge entre o iframe matricial e a aplicação host via injeção/submissão dinâmica no contexto pai (`window.parent.document`), contornando restrições de sandbox. O banco SQLite recalcula automaticamente o saldo de folgas na proporção 1,5x e preserva o mês e ano visualizados.
3. **Exclusão de Períodos:**
   - Ao clicar em **"🗑️ Excluir"**, o sistema solicita confirmação via pop-up nativo.
   - Após a confirmação, o sistema identifica com precisão as datas de início e término do período e submete a requisição ao backend.
   - Se o período for um embarque (`EMB`), o sistema remove tanto os dias de embarque quanto as folgas projetadas subsequentes associadas (`status = 'FOL'`), atualizando o banco de dados e registrando a exclusão no Log de Auditoria.
   - Para períodos importados exclusivamente do cronograma geral (Excel), o sistema notifica informativamente que a alteração deve ocorrer na planilha fonte.
4. **Fechamento do Modal:**
   - Pode ser fechado clicando no `X`, pressionando a tecla `Esc` ou clicando em qualquer ponto fora do cartão (no fundo escuro/backdrop).
5. **Limpeza de Escalas:**
   - O painel do editor disponibiliza limpeza pontual por colaborador ou limpeza completa do mês mediante confirmação expressa por caixa de seleção.

---

## 9. Sistema de Sugestões de Implementação

A aba **Sugestões** funciona como canal de melhoria contínua, permitindo que a equipe submeta ideias de evolução do sistema com triagem e priorização estruturadas.

### 9.1 Submissão de Sugestões (Público)
- Aberta a todos os usuários (mesmo em modo consulta/visualizador).
- O autor pode se identificar nominalmente informando seu nome ou assinalar a caixa de seleção **"Anônimo"**.
- Ao assinalar **"Anônimo"**, a autoria é gravada no banco como `Anônimo` (`is_anonimo = 1`), preservando a privacidade.
- Campos obrigatórios: Título/Assunto da sugestão e Descrição detalhada da proposta/implementação.
- Status inicial gravado no banco: **`Pendente`**.

### 9.2 Triagem e Avaliação (Apenas Modo Edição)
- Sugestões com status `Pendente` **aparecem exclusivamente quando o usuário está em modo edição** (`is_editor`). Em modo visualizador, elas permanecem ocultas.
- O editor analisa o título, a descrição e o autor, e seleciona o nível de prioridade:
  - **`Alta`**
  - **`Média`**
  - **`Baixa`**
- O editor pode:
  - **Aprovar:** O status é alterado para `Aprovada` com a prioridade selecionada.
  - **Rejeitar:** O status é alterado para `Rejeitada`.
- Todas as decisões de aprovação ou rejeição são auditadas no Log de Edições.

### 9.3 Lista de Sugestões a Serem Implementadas (Público)
- Uma vez aprovada, a sugestão passa a constar publicamente na seção **"Sugestões a Serem Implementadas"**.
- **Ordenação Obrigatória por Prioridade:**
  $$\text{Prioridade Alta} \longrightarrow \text{Prioridade Média} \longrightarrow \text{Prioridade Baixa}$$
  Dentro da mesma prioridade, o desempate é feito pela data de criação.
- Cada cartão exibe o badge de prioridade estilizado, título, autor, data e descrição.

### 9.4 Sinalização de Conclusão e Bloco de Implementadas
- No modo edição, cada sugestão da lista a ser implementada disponibiliza a ação **"🚀 Marcar como Implementada"**.
- Ao acionar essa opção:
  1. O status é atualizado para **`Implementada`**.
  2. A data/hora exata da entrega é registrada (`implementada_at`).
  3. A sugestão é transferida automaticamente para a seção inferior: **"Sugestões Implementadas"**.
- O bloco de sugestões implementadas serve como histórico público de conquistas e melhorias entregues no sistema.

---

## 10. Log de Auditoria & Histórico de Edições

A aba **Log de Edições** oferece rastreabilidade total, segurança e transparência operacional, gravando de forma cronológica e imutável cada modificação realizada na base de dados.

### 10.1 Operações Auditadas
O sistema intercepta e audita automaticamente:
- **`INSERÇÃO`:** Lançamento de períodos de escala (`EMB`, `FOL`, `BASE`, etc.) e projeções de ciclos contínuos 14x21.
- **`EDIÇÃO`:** Modificação de datas de início/término ou condição de períodos existentes com recálculo automático de folgas.
- **`EXCLUSÃO`:** Remoção de períodos de escala e desativação das folgas subsequentes associadas.
- **`AUTO-PERMUTA`:** Troca de períodos de folga entre datas dentro da janela de até 180 dias.
- **`LIMPEZA`:** Limpeza de escalas pontuais por colaborador ou de plataformas inteiras no mês.
- **`EQUIPE`:** Cadastro, reativação ou desativação de colaboradores na equipe.
- **`SUGESTÃO`:** Submissão de novas sugestões, aprovação/rejeição com priorização e conclusão de implementações.

### 10.2 Atributos Gravados no Log
Cada registro na tabela `p80_audit_log` armazena:
1. **Data e Hora:** Horário local exato de Brasília no momento da operação (`AAAA-MM-DD HH:MM:SS`), exibido no padrão brasileiro `DD/MM/AAAA HH:MM:SS`.
2. **Usuário / Responsável:** Captura do usuário do sistema operacional (`os.environ.get("USERNAME")`), acrescido da identificação de papel ativo (ex: `viner (Editor)`).
3. **Tipo de Ação:** Categoria da operação com padronização em maiúsculas (`INSERÇÃO`, `EDIÇÃO`, `EXCLUSÃO`, etc.).
4. **Plataforma / Unidade:** Unidade afetada (`P-80`, `P-82`, `P-83` ou `GERAL`).
5. **Colaborador:** Nome do profissional diretamente impactado pela escala.
6. **Detalhes da Modificação:** Descrição completa contendo datas anteriores, novas datas, quantidade de dias de embarque, folgas geradas ou justificativas registradas.

### 10.3 Recursos de Visualização e Exportação
- **Cards de Métricas:** Total de ações registradas, total de inserções, total de edições/permutas e total de exclusões/limpezas.
- **Filtros Dinâmicos:** Filtro por plataforma/projeto, filtro por tipo de ação, busca textual livre por colaborador/termo e controle de volume de registros.
- **Exportação CSV:** Botão para download instantâneo de todo o log filtrado em formato CSV com codificação UTF-8.

---

## 11. Registro de Versões

- **v0.1 (Atual):**
  - Implementação das 3 abas individuais de plataforma (**P-80**, **P-82**, **P-83**).
  - Exclusão estrita de supervisores e coordenadores das escalas de plataforma.
  - Regra proporcional estrita de 1,5x na criação e redimensionamento de embarques.
  - Alocação em escala cíclica contínua no regime 14x21 (projeção de $K$ ciclos completos).
  - Janela e fluxo completo de **Auto-Permuta** de folga (mesma duração, $\le 180$ dias).
  - Seção consolidada **Divisão de Turma a Bordo** com soma diária e balões contínuos coloridos por disciplina (Azul = Mecânica, Laranja = Instrumentação, Amarelo = Elétrica).
  - Correção e robustecimento da ação de exclusão de períodos via modal pop-up.
  - Integração do cronograma externo com mapeamento de Estaleiro como `BASE`.
  - Sinalização em vermelho e legenda para status de conflito `CONF`.
  - **Aba de Sugestões:** Submissão nominal ou anônima, triagem exclusiva em modo edição com definição de prioridade (Alta, Média, Baixa), lista pública ordenada por prioridade e transferência para o bloco de implementadas.
  - **Aba Log de Edições:** Registro detalhado com horário local de Brasília, identificação do usuário editor, tipo de ação, plataforma, colaborador e exportação em CSV.
  - **Hover Informativo em Conflitos (CONF):** Exibição instantânea no hover do mouse da divergência exata entre a escala e o cronograma.
  - **Navegação Temporal Sincronizada:** Correção dos botões de Mês Anterior e Próximo Mês via callbacks, eliminando conflitos de estado com os seletores.
  - **Aba Exportar:** Ajuste na renderização condicional dos eventos no período.
  - Proteção por senha `"admin"` e modal embutido de alta performance.

