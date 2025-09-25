# 🏃‍♂️ Script de Sincronização de Treinos - Google Sheets + Calendar + Tasks

## 📋 Descrição

Este script automatiza a sincronização de treinos entre Google Sheets, Google Calendar e Google Tasks, oferecendo:

- ⏰ **Horários dinâmicos** configuráveis por treino
- 🎨 **Cores inteligentes** baseadas no tipo de treino  
- 🔔 **Notificações personalizadas**
- 🔄 **Sincronização bidirecional** de status
- 📝 **Templates customizáveis**

## 🚀 Configuração Inicial

### 1. Preparar o Google Apps Script

1. Acesse [Google Apps Script](https://script.google.com)
2. Crie um novo projeto
3. Cole o código do arquivo `script_calendar_community.js`
4. Ative as APIs necessárias:
   - Google Calendar API
   - Google Tasks API

### 2. Configurar as Credenciais

Edite a seção `CONFIG` no início do script:

```javascript
const CONFIG = {
  nomePlanilha: "SEU_NOME_DA_PLANILHA",     // Nome exato da sua planilha
  nomeAba: "SUA_ABA_DE_TREINOS",            // Nome da aba com os treinos
  taskListId: "SEU_TASK_LIST_ID",           // Descobrir executando obterTaskListId()
  // ... outras configurações
};
```

### 3. Descobrir o Task List ID

1. Execute a função `obterTaskListId()`
2. Vá em **Executions** → **View execution transcript**
3. Copie o ID da lista desejada
4. Cole em `CONFIG.taskListId`

## 📊 Estrutura da Planilha

Crie uma planilha com as seguintes colunas:

| A | B | C | D | E | F | G | H | I | J | K |
|---|---|---|---|---|---|---|---|---|---|---|
| Semana | Status | Action | Dia | Horário | Dia Semana | Treino | Zona Frequência | Descrição | Pós Treino | EventID |

### Colunas Obrigatórias:

- **C (Action)**: `Add`, `Update` ou `Remove`
- **D (Dia)**: Data do treino
- **E (Horário)**: Horário em formato 24h (`06:00`, `18:30`)
- **G (Treino)**: Nome/tipo do treino

### Colunas Opcionais:

- **A (Semana)**: Semana do programa
- **B (Status)**: Atualizado automaticamente
- **F (Dia Semana)**: Calculado automaticamente
- **H (Zona Frequência)**: Zona cardíaca alvo
- **I (Descrição)**: Detalhes do treino
- **J (Pós Treino)**: Observações após o treino
- **K (EventID)**: IDs internos (não editar)

## 🎯 Como Usar

### Adicionar Treino:
1. Preencha as colunas obrigatórias
2. Digite `Add` na coluna Action
3. Execute `sincronizarTreinosCompleto()`

### Atualizar Treino:
1. Modifique os dados desejados
2. Digite `Update` na coluna Action
3. Execute a sincronização

### Remover Treino:
1. Digite `Remove` na coluna Action
2. Execute a sincronização

## ⚙️ Personalização

### Formatos de Horário Aceitos:
- `06:00` (recomendado)
- `18h30`
- `07.45`
- `1930`

### Cores Automáticas por Tipo:
- 🔴 **Vermelho**: Intervalado, tiros, velocidade
- 🔵 **Azul**: Treinos longos, distância
- 🟢 **Verde**: Recuperação, regenerativo
- 🟠 **Laranja**: Outros treinos (padrão)

### Notificações Padrão:
- 24 horas antes
- 12 horas antes
- 6 horas antes
- 1 hora antes

## 🔧 Funções Disponíveis

### Principais:
- `sincronizarTreinosCompleto()` - Sincronização principal
- `verificarStatusCompleto()` - Atualiza status automaticamente

### Configuração:
- `obterTaskListId()` - Descobre IDs das listas de tarefas
- `debugEstruturaPlanilha()` - Verifica estrutura da planilha

### Menu na Planilha:
O script adiciona um menu personalizado com as opções principais.

## 📝 Templates Personalizáveis

### Google Tasks:
```
🎯 TREINO DO DIA: {TREINO}
📅 Data: {DATA_COMPLETA}
❤️ Zona de Frequência: {ZONA_FREQ}
📝 DESCRIÇÃO: {DESCRICAO}
```

### Google Calendar:
```
🏃‍♂️ PLANO DE TREINO
🎯 TREINO: {TREINO}
📅 Data: {DATA_COMPLETA}
⏰ Horário: {HORARIO}
❤️ Zona de Frequência: {ZONA_FREQ}
📝 INSTRUÇÕES: {DESCRICAO}
```

### Placeholders Disponíveis:
- `{TREINO}` - Nome do treino
- `{DATA_COMPLETA}` - Data por extenso
- `{HORARIO}` - Horário formatado
- `{ZONA_FREQ}` - Zona de frequência
- `{DESCRICAO}` - Descrição detalhada

## 🛠️ Solução de Problemas

### Erro de Horário:
- Verifique se o formato está correto (`HH:MM`)
- Certifique-se que a coluna E não está vazia

### Erro de Planilha:
- Confirme os nomes em `CONFIG.nomePlanilha` e `CONFIG.nomeAba`
- Execute `debugEstruturaPlanilha()` para verificar

### Erro de Task List:
- Execute `obterTaskListId()` novamente
- Verifique se o ID está correto em `CONFIG.taskListId`

### APIs não habilitadas:
1. Vá em **Services** no Apps Script
2. Adicione **Google Calendar API**
3. Adicione **Google Tasks API**

## 📈 Status Automáticos

O script atualiza automaticamente os status:

- **Pendente**: Treino agendado, não realizado
- **Concluído**: Tarefa marcada como concluída no Tasks
- **Expirado**: Passou do prazo + tolerância (padrão: 1 dia)

## 🎨 Personalização Avançada

### Modificar Duração Padrão:
```javascript
horario: {
  duracaoMinutos: 90  // Para treinos de 1h30
}
```

### Adicionar Notificações:
```javascript
notificacoes: [
  2 * 60,    // 2 horas antes
  30,        // 30 minutos antes
  // ...
]
```

### Customizar Cores:
```javascript
cores: {
  intervalado: CalendarApp.EventColor.PURPLE,
  // ...
}
```

## 🤝 Contribuições

Contribuições são bem-vindas! Para contribuir:

1. Fork o projeto
2. Crie uma branch para sua feature
3. Commit suas mudanças
4. Abra um Pull Request

## 📄 Licença

Este projeto está sob a licença MIT. Veja o arquivo `LICENSE` para mais detalhes.

## 🔗 Links Úteis

- [Google Apps Script](https://script.google.com)
- [Google Calendar API](https://developers.google.com/calendar)
- [Google Tasks API](https://developers.google.com/tasks)

## 📞 Suporte

Para dúvidas ou problemas:
1. Verifique a seção **Solução de Problemas**
2. Execute `debugEstruturaPlanilha()` para diagnóstico
3. Abra uma issue no GitHub

---

⭐ **Se este script foi útil, considere dar uma estrela no repositório!**
