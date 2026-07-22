# Contexto: projeto de cotagem automática (pyRevit)

Cole este prompt no início de uma nova conversa para retomar o projeto de onde parei.

## O que é o projeto

Um plugin pyRevit (IronPython 2, API do Revit) que cria `Dimension`s
automaticamente numa planta/vista. Existem/existiram três variantes do
script, todas nomeadas `script.py`:

1. **"antigo" (v2.1, `Cotar Selecao`)** - versão mais antiga. Unidade
   central = PAREDE. Toda parede processada individualmente sempre
   recebe cota geral (+ subcotas quando há detalhe). Correntes/
   alinhamentos existem só para gerar cotas de conjunto
   (`alinhamento_total`/`geral`), nunca substituem a cota da parede.
   Roda a partir de **seleção manual** do usuário (`PickObjects`).

2. **"atual" (Fase 4.3)** - versão mais nova, bem mais enxuta (~700
   linhas vs ~1600). Roda **automaticamente** sobre todas as paredes da
   view, sem seleção. Cada parede é processada de forma totalmente
   isolada (não existe mais o conceito de corrente/alinhamento). Usa
   offsets fixos (2.0/1.0/0.4 pés) a partir da própria linha da parede,
   decide o lado "de fora" com um heurística geométrica (centroide
   geral do edifício via contorno do piso), e introduziu uma cota nova:
   distância entre paredes paralelas mais próximas (largura de
   cômodo/corredor) - por definição, uma cota interna.

3. **"Script 1 - Cotar(sel)"** - o que estamos evoluindo agora. É
   literalmente a base da v2.1 antiga (mesmo arquivo, 1601 linhas),
   focado em **seleção manual + perímetro**.

## Decisão de arquitetura (já tomada, não reabrir)

Vamos manter **dois scripts separados**, cada um evoluindo em fases:

- **Script 1 (perímetro, seleção manual)** - base = v2.1 antiga. É o que
  estamos mexendo agora.
- **Script 2 (interior, automático)** - ainda não começamos. Quando
  começar, a sugestão (ainda não confirmada pelo usuário) é partir da
  base "atual" (Fase 4.3), porque ela já tem a coleta rica (aberturas,
  cruzamentos, parede paralela mais próxima) que uma cotagem automática
  de interior precisa.

Não fundir os dois de volta num script só. Não voltar para a arquitetura
antiga de correntes globais na versão atual - aquilo já foi analisado e
descartado.

## Metodologia de trabalho (seguir sempre)

- Implementar em **fases pequenas**, uma de cada vez.
- Em cada fase: reaproveitar o máximo possível do código existente,
  mudar só o necessário para o objetivo da fase.
- Não implementar uma fase sem o usuário pedir explicitamente.
- Depois de cada mudança: validar sintaxe, e sempre que possível
  **testar contra o modelo real do Revit** (ver seção de conexão
  abaixo) antes de dar a mudança como resolvida - não confiar só em
  análise estática/álgebra manual.
- Entregar o arquivo atualizado (`present_files`) para o usuário testar
  no Revit de verdade e reportar o resultado.
- Se o usuário disser "ficou errado, vamos tentar de novo" - **não
  assumir que o diagnóstico anterior estava certo**. Voltar a
  investigar do zero, de preferência com dados reais do modelo (via
  MCP), em vez de insistir no mesmo raciocínio.

## Conexão com o Revit (MCP)

Há um conector MCP (`revit-pyrevit`) que permite, nesta mesma conversa://
- `get_revit_status` - checar se está conectado
- `get_current_view_info` - metadados da view ativa
- `get_current_view_elements` - elementos visíveis na view
- `execute_revit_code` - rodar IronPython arbitrário no contexto do
  documento aberto (`doc`, `uidoc`, `DB`, `revit`), sem transação
  automática (preciso abrir/commitar `DB.Transaction` manualmente se for
  alterar o modelo)
- Outras (list_revit_views, place_family, etc.) via `tool_search` se
  precisar.

Isso foi essencial para o achado do bug real (ver abaixo) - análise
puramente estática/algébrica levou a uma hipótese errada; só rodando
código de verdade contra o modelo aberto foi possível achar a causa
raiz.

## Estado atual do Script 1 (`Cotar(sel)`, em andamento)

**Sintoma relatado pelo usuário:** ao selecionar uma parede (ex.:
"embaixo" na planta), a cota às vezes nasce longe dela, "lá em cima".

**Primeira tentativa (v2.2) - PARCIALMENTE ERRADA, não repetir:**
Hipótese: `centro_perp_modelo` (usado em `processar_paredes_individualmente`
para decidir de que lado da parede a cota vai) era calculado com a
média de *todas* as paredes paralelas da view inteira, contaminado por
paredes distantes sem relação com a parede selecionada. Fix aplicado:
restringir essa média às faces de contexto cujo `pos_axis` sobrepõe o
próprio trecho da parede (usando a tolerância `cluster_tol` que já
existia). Esse fix é razoável e foi mantido, mas **não resolveu o
sintoma relatado** - havia uma causa mais profunda.

**Segunda tentativa (v2.3) - causa raiz real, confirmada ao vivo no
Revit:**
O bug não era "qual lado", era o **ponto de referência em si**. A
função `extrair_faces_referenciaveis` (única função do pipeline que
calcula `pos_axis`/`pos_perp` de qualquer face) usava
`face.Origin` da API do Revit. Descoberta: `PlanarFace.Origin` **não é
garantidamente um ponto dentro dos limites visíveis da face** - é só um
ponto usado para definir o plano/sistema de coordenadas da face. Em
faces de ponta/topo geradas por um encontro (miter/join) com outra
parede, isso pode devolver um ponto bem longe da face real - no caso
testado ao vivo, a face de ponta de uma parede em Y=31,33 devolvia
`Origin.Y = 42,49` (a posição exata de OUTRA parede do modelo).
Confirmado com dois métodos independentes (Triangulate + bbox-UV
Evaluate), ambos batendo em Y=31,33 (a posição real).

**Fix aplicado:** troquei `face.Origin` por um ponto avaliado de
verdade dentro da face (`face.GetBoundingBox()` em UV + `face.Evaluate`
no meio dessa bbox), com fallback pro `Origin` cru só se a avaliação
falhar (face degenerada). Testado ao vivo contra o modelo real via MCP:
o valor corrigido bate exatamente com a `LocationCurve` da parede.

**Aguardando confirmação do usuário** de que a cota agora nasce colada
na parede selecionada, testando no Revit de verdade.

## Próximos passos possíveis (não fazer sem pedir)

- Continuar refinando o Script 1 (perímetro/seleção) enquanto houver
  bugs/ajustes a pedido do usuário.
- Eventualmente começar o Script 2 (interior, automático), decidindo
  junto com o usuário a base de partida e as fases.
- Uma ideia já desenhada (mas explicitamente NÃO implementada, só
  documentada) para uma limpeza arquitetural futura, caso o usuário
  volte a pedir "unificar" ou reduzir a fragmentação parede-a-parede da
  versão "atual": criar uma camada intermediária `AlignmentBuilder` que
  agrupa paredes colineares antes do posicionamento das cotas, mantendo
  toda a coleta rica da versão atual. Isso foi cancelado a meio caminho
  da implementação (Fase 1 do AlignmentBuilder) a pedido do usuário -
  não retomar a menos que ele peça de novo.
