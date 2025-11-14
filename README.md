# Gestor de Tarefas (Google Apps Script)

Este projeto disponibiliza o código-fonte de um aplicativo de acompanhamento de tarefas construído em Google Apps Script, utilizando uma planilha do Google Sheets como base de dados.

## Funcionalidades principais

- Cadastro e edição de tarefas via barra lateral.
- Agrupamento e ordenação padrão por projeto, data FUP, prioridade e esforço.
- Campos de apoio para descrição, datas de FUP e deadline, projeto, status, prioridade, esforço e observações.
- Suporte a tarefas recorrentes com periodicidade diária, semanal (com escolha de dias) ou mensal.
- Filtros avançados por projeto, status, prioridade, esforço e intervalo de datas.
- Botão “Classificar (padrão)” que restaura a ordenação padrão e remove filtros, ocultando tarefas concluídas e canceladas.
- Rotina para atualização de tarefas recorrentes com cálculo automático da próxima ocorrência.

## Estrutura de arquivos

- `Code.gs`: script principal que cria o menu, manipula os dados da planilha e aplica regras de ordenação/recorrência.
- `TaskForm.html`: formulário exibido na barra lateral para cadastrar ou editar tarefas.
- `FilterSidebar.html`: barra lateral para aplicar filtros personalizados e restaurar a classificação padrão.

## Configuração da planilha

1. Crie uma planilha em branco no Google Sheets e renomeie para o nome desejado.
2. Abra **Extensões → Apps Script** e substitua o conteúdo do arquivo `Code.gs` pelo código fornecido neste repositório.
3. Adicione os arquivos HTML `TaskForm.html` e `FilterSidebar.html` no editor do Apps Script com o mesmo conteúdo deste projeto.
4. Salve o projeto e volte para a planilha.

Na primeira execução o script criará (ou atualizará) automaticamente uma aba chamada `Tarefas` com o cabeçalho:

```
ID | Título | Descrição | Projeto | Status | Prioridade | Esforço | Data FUP | Data Limite | Tipo Recorrência | Configuração Recorrência | Observações | Criado em | Atualizado em | Próxima Ocorrência
```

## Uso

1. **Adicionar tarefa**: abra o menu **Gestor de Tarefas → Adicionar tarefa** e preencha o formulário.
2. **Filtros personalizados**: use **Gestor de Tarefas → Aplicar filtros** para refinar a visualização.
3. **Classificar (padrão)**: utilize o botão no menu ou na barra de filtros para remover filtros e restaurar a visualização padrão (tarefas concluídas/canceladas ficam ocultas).
4. **Recorrências**: tarefas com status "Recorrente" exibem configurações extras. Execute periodicamente **Gestor de Tarefas → Atualizar recorrências** para avançar as datas conforme a periodicidade definida.

## Observações

- Os IDs das tarefas são gerados automaticamente e persistidos nas propriedades do documento.
- Ajuste os valores de prioridade e esforço conforme a escala desejada.
- Para edições diretas na planilha, certifique-se de manter o cabeçalho e a estrutura de colunas intactos.

## Licença

Distribuído sob a licença MIT. Consulte o arquivo `LICENSE` se disponível ou adapte conforme necessário.
