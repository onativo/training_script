// ===== SCRIPT DE SINCRONIZAÇÃO DE TREINOS - VERSÃO COMUNIDADE ===== //
// 
// Este script sincroniza automaticamente treinos de uma planilha do Google Sheets
// com Google Calendar e Google Tasks, permitindo:
// - Horários dinâmicos configuráveis por treino
// - Notificações personalizadas
// - Cores inteligentes baseadas no tipo de treino
// - Sincronização bidirecional de status
//
// COMO CONFIGURAR:
// 1. Substitua os valores em CONFIG com suas informações
// 2. Configure a estrutura da planilha conforme comentários
// 3. Ative as APIs: Google Calendar, Google Tasks
// 4. Execute a função obterTaskListId() para encontrar seu taskListId
//
// ====================================================================== //

// ===== CONFIGURAÇÕES - EDITE ESTAS INFORMAÇÕES ===== //
const CONFIG = {
  nomePlanilha: "SEU_NOME_DA_PLANILHA",           // Nome exato da sua planilha
  nomeAba: "SUA_ABA_DE_TREINOS",                  // Nome da aba com os treinos
  taskListId: "SEU_TASK_LIST_ID",                 // Execute obterTaskListId() para descobrir
  
  horario: {
    duracaoMinutos: 60                            // Duração padrão dos treinos (em minutos)
  },
  
  // Notificações antes do treino (em minutos)
  notificacoes: [
    24 * 60, // 24 horas antes (1440 minutos)
    12 * 60, // 12 horas antes (720 minutos)
    6 * 60,  // 6 horas antes (360 minutos)
    1 * 60   // 1 hora antes (60 minutos)
  ],
  
  // Sistema de cores inteligente por tipo de treino
  cores: {
    evento: CalendarApp.EventColor.ORANGE,       // Cor principal dos treinos
    intervalado: CalendarApp.EventColor.RED,     // Treinos intensos (intervalado, tiros)
    longo: CalendarApp.EventColor.BLUE,          // Treinos longos (distância)
    recuperacao: CalendarApp.EventColor.GREEN,   // Treinos de recuperação
    default: CalendarApp.EventColor.ORANGE       // Cor padrão
  },
  
  // ESTRUTURA DA PLANILHA - Ajuste conforme sua planilha
  // Índices baseados em 0 (A=0, B=1, C=2, etc.)
  colunas: {
    SEMANA: 0,       // A - Semana do treino (opcional)
    STATUS: 1,       // B - Status: Pendente/Concluído/Expirado
    ACTION: 2,       // C - Ação: Add/Update/Remove
    DIA: 3,          // D - Data do treino
    HORARIO: 4,      // E - Horário do treino (formato 24h: "06:00", "18:30")
    DIA_SEMANA: 5,   // F - Dia da semana (opcional)
    TREINO: 6,       // G - Tipo/Nome do treino
    ZONA_FREQ: 7,    // H - Zona de frequência cardíaca
    DESCRICAO: 8,    // I - Descrição detalhada do treino
    POS_TREINO: 9,   // J - Observações pós-treino (opcional)
    EVENT_ID: 10     // K - IDs dos eventos (não editar manualmente)
  },
  
  // Configurações de status automático
  status: {
    pendente: "Pendente",
    concluido: "Concluído", 
    expirado: "Expirado",
    naoEncontrado: "ID not found",
    diasTolerancia: 1, // Dias após o treino para considerar expirado
    statusParaLimpar: ["Pendente", "Concluído", "Expirado"]
  },
  
  // Templates personalizáveis para Tasks e Events
  templates: {
    taskNotes: `════════════════════════
🎯 TREINO DO DIA: {TREINO}
════════════════════════

📅 Data: {DATA_COMPLETA}
❤️ Zona de Frequência: {ZONA_FREQ}

📝 DESCRIÇÃO:
┌───────────────────────────┐
│ {DESCRICAO}
└───────────────────────────┘

💪 Bom treino! Mantenha o foco e a determinação!
🏆 Cada treino te aproxima da sua meta!`,
    
    eventDescription: `════════════════════════
🏃‍♂️ PLANO DE TREINO
════════════════════════

🎯 TREINO: {TREINO}
📅 Data: {DATA_COMPLETA}
⏰ Horário: {HORARIO}
❤️ Zona de Frequência: {ZONA_FREQ}

📝 INSTRUÇÕES:
┌───────────────────────────┐
│ {DESCRICAO}
└───────────────────────────┘

💡 DICAS:
• Faça aquecimento antes de começar
• Mantenha-se hidratado durante o treino
• Respeite os limites do seu corpo
• Alongue-se ao final do exercício

🏆 Foco na meta! Cada treino conta!
💪 Você consegue!`
  }
};

// ===== FUNÇÕES AUXILIARES ===== //

function getSpreadsheetAndSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  if (spreadsheet.getName() !== CONFIG.nomePlanilha) {
    throw new Error(`Esta não é a planilha correta. Esperado: ${CONFIG.nomePlanilha}`);
  }
  
  const sheet = spreadsheet.getSheetByName(CONFIG.nomeAba);
  if (!sheet) {
    throw new Error(`Aba '${CONFIG.nomeAba}' não encontrada.`);
  }
  
  return { spreadsheet, sheet };
}

function formatarData(data) {
  const diasSemana = ['Dom', 'Seg', 'Ter', 'Qua', 'Qui', 'Sex', 'Sáb'];
  const meses = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'];
  
  const dia = data.getDate().toString().padStart(2, '0');
  const mes = meses[data.getMonth()];
  const diaSemana = diasSemana[data.getDay()];
  
  return `${diaSemana}, ${dia}/${mes}`;
}

function formatarDataCompleta(data) {
  const diasSemana = ['Domingo', 'Segunda-feira', 'Terça-feira', 'Quarta-feira', 'Quinta-feira', 'Sexta-feira', 'Sábado'];
  const meses = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'];
  
  const dia = data.getDate();
  const mes = meses[data.getMonth()];
  const ano = data.getFullYear();
  const diaSemana = diasSemana[data.getDay()];
  
  return `${diaSemana}, ${dia} de ${mes} de ${ano}`;
}

/**
 * Analisa e valida horário em formato 24h
 * Aceita formatos: "18:30", "18h30", "18.30", "1830"
 */
function analisarHorario(horarioTexto) {
  // Se for um objeto Date/Time, converte para string HH:MM
  if (horarioTexto instanceof Date) {
    const horas = horarioTexto.getHours().toString().padStart(2, '0');
    const minutos = horarioTexto.getMinutes().toString().padStart(2, '0');
    horarioTexto = `${horas}:${minutos}`;
  }
  
  // Converte para string se não for
  if (typeof horarioTexto !== 'string') {
    horarioTexto = String(horarioTexto);
  }
  
  if (!horarioTexto || horarioTexto.trim() === '') {
    Logger.log(`❌ ERRO: Horário é obrigatório na coluna E (Horário)`);
    throw new Error(`❌ HORÁRIO OBRIGATÓRIO: O horário deve ser preenchido na coluna E (Horário). Exemplo: "06:00", "18:30"`);
  }

  // Remove espaços e normaliza o texto
  const horarioLimpo = horarioTexto.trim();
  
  // Aceita formatos: "18:30", "18h30", "18.30", "1830"
  const regexHorario = /^(\d{1,2})[:h.]?(\d{2})?$/;
  const match = horarioLimpo.match(regexHorario);
  
  if (!match) {
    Logger.log(`❌ ERRO: Formato de horário inválido: "${horarioTexto}"`);
    throw new Error(`❌ FORMATO INVÁLIDO: O horário "${horarioTexto}" não é válido. Use formatos como: "06:00", "18:30", "07h15", "1945"`);
  }
  
  const horas = parseInt(match[1], 10);
  const minutos = match[2] ? parseInt(match[2], 10) : 0;
  
  // Valida horas (0-23) e minutos (0-59)
  if (horas < 0 || horas > 23 || minutos < 0 || minutos > 59) {
    Logger.log(`❌ ERRO: Horário fora dos limites: ${horas}:${minutos}`);
    throw new Error(`❌ HORÁRIO INVÁLIDO: ${horas}:${minutos} está fora dos limites válidos. Horas: 0-23, Minutos: 0-59`);
  }
  
  Logger.log(`✅ Horário analisado: ${horas}:${minutos.toString().padStart(2, '0')}`);
  return { horas, minutos, valido: true };
}

/**
 * Cria horários de início e fim para o evento baseado no horário da planilha
 */
function criarHorarioEvento(data, horarioTexto) {
  const horarioAnalisado = analisarHorario(horarioTexto);
  
  const inicio = new Date(data);
  inicio.setHours(horarioAnalisado.horas, horarioAnalisado.minutos, 0, 0);
  
  // Calcula fim baseado na duração configurada
  const fim = new Date(inicio);
  fim.setMinutes(fim.getMinutes() + CONFIG.horario.duracaoMinutos);
  
  const horarioFormatado = `${horarioAnalisado.horas}:${horarioAnalisado.minutos.toString().padStart(2, '0')}`;
  const fimFormatado = `${fim.getHours()}:${fim.getMinutes().toString().padStart(2, '0')}`;
  
  return { 
    inicio, 
    fim, 
    horarioTexto: `${horarioFormatado} - ${fimFormatado}`,
    horarioAnalisado
  };
}

/**
 * Processa templates substituindo placeholders pelos dados reais
 */
function processarTemplate(template, dados) {
  return template
    .replace('{TREINO}', dados.treino.toUpperCase())
    .replace('{DATA_COMPLETA}', formatarDataCompleta(dados.dia))
    .replace('{HORARIO}', dados.horarioTexto || 'Horário não especificado')
    .replace('{ZONA_FREQ}', dados.zonaFrequencia || 'Não especificada')
    .replace('{DESCRICAO}', dados.descricao || (template.includes('INSTRUÇÕES') ? 'Siga o plano de treino padrão' : 'Sem descrição adicional'));
}

function criarTitulos(treino, dia) {
  const dataFormatada = formatarData(dia);
  return {
    task: `🏃‍♂️ ${treino} | ${dataFormatada}`,
    event: `🏃‍♂️ ${treino} - ${dataFormatada}`
  };
}

/**
 * Configura notificações personalizadas para o evento
 */
function configurarNotificacoes(event) {
  event.removeAllReminders();
  CONFIG.notificacoes.forEach(minutos => {
    event.addPopupReminder(minutos);
  });
  
  const horas = CONFIG.notificacoes.map(min => `${min/60}h`).join(', ');
  Logger.log(`🔔 Notificações configuradas: ${horas} antes`);
}

/**
 * Sistema inteligente de cores baseado no tipo de treino
 */
function determinarCorEvento(treino) {
  const treinoLower = treino.toLowerCase();
  
  if (treinoLower.includes('interval') || treinoLower.includes('tiro') || treinoLower.includes('velocidade')) {
    return CONFIG.cores.intervalado; // Vermelho para treinos intensos
  } else if (treinoLower.includes('longo') || treinoLower.includes('longa') || treinoLower.includes('distancia')) {
    return CONFIG.cores.longo; // Azul para treinos longos
  } else if (treinoLower.includes('recupera') || treinoLower.includes('regenera') || treinoLower.includes('leve')) {
    return CONFIG.cores.recuperacao; // Verde para recuperação
  } else {
    return CONFIG.cores.default; // Cor padrão
  }
}

function configurarEventoCompleto(event, treino) {
  // Configura notificações
  configurarNotificacoes(event);
  
  // Configura cor baseada no tipo de treino
  const cor = determinarCorEvento(treino);
  event.setColor(cor);
  
  Logger.log(`🎨 Cor do evento configurada: ${getCorNome(cor)} para treino "${treino}"`);
}

function getCorNome(eventColor) {
  const cores = {
    [CalendarApp.EventColor.PALE_BLUE]: 'Azul Claro',
    [CalendarApp.EventColor.PALE_GREEN]: 'Verde Claro', 
    [CalendarApp.EventColor.MAUVE]: 'Roxo Claro',
    [CalendarApp.EventColor.PALE_RED]: 'Vermelho Claro',
    [CalendarApp.EventColor.YELLOW]: 'Amarelo',
    [CalendarApp.EventColor.ORANGE]: 'Laranja',
    [CalendarApp.EventColor.CYAN]: 'Ciano',
    [CalendarApp.EventColor.GRAY]: 'Cinza',
    [CalendarApp.EventColor.BLUE]: 'Azul',
    [CalendarApp.EventColor.GREEN]: 'Verde',
    [CalendarApp.EventColor.RED]: 'Vermelho'
  };
  return cores[eventColor] || 'Padrão';
}

/**
 * Cria menu personalizado na planilha
 */
function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu("🔁 Sincronizar Treinos")
    .addItem("📅 Sincronizar Treinos", "sincronizarTreinosCompleto")
    .addItem("🔄 Atualizar Status", "verificarStatusCompleto")
    .addToUi();
}

// ===== SINCRONIZAÇÃO PRINCIPAL ===== //

/**
 * Função principal: Sincroniza treinos da planilha com Calendar e Tasks
 * 
 * COMO USAR:
 * 1. Na coluna Action (C), digite:
 *    - "Add" para criar novo treino
 *    - "Update" para atualizar treino existente
 *    - "Remove" para remover treino
 * 2. Execute esta função ou use o menu personalizado
 */
function sincronizarTreinosCompleto() {
  try {
    const { sheet } = getSpreadsheetAndSheet();
    const data = sheet.getDataRange().getValues();
    const calendar = CalendarApp.getDefaultCalendar();
    let processedCount = 0;

    // Processa todas as linhas de uma vez para otimizar performance
    const updates = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const acao = row[CONFIG.colunas.ACTION];
      
      if (!acao) continue;

      const dadosTreino = {
        dia: new Date(row[CONFIG.colunas.DIA]),
        horario: row[CONFIG.colunas.HORARIO],
        treino: row[CONFIG.colunas.TREINO],
        zonaFrequencia: row[CONFIG.colunas.ZONA_FREQ],
        descricao: row[CONFIG.colunas.DESCRICAO],
        eventId: row[CONFIG.colunas.EVENT_ID]
      };

      // Cria informações de horário primeiro
      const horarioInfo = criarHorarioEvento(dadosTreino.dia, dadosTreino.horario);
      dadosTreino.horarioTexto = horarioInfo.horarioTexto;
      
      const titulos = criarTitulos(dadosTreino.treino, dadosTreino.dia);
      const taskNotes = processarTemplate(CONFIG.templates.taskNotes, dadosTreino);
      const eventDescription = processarTemplate(CONFIG.templates.eventDescription, dadosTreino);

      const resultado = processarAcao(acao, calendar, titulos, taskNotes, eventDescription, dadosTreino, i + 1);
      
      if (resultado.sucesso) {
        processedCount++;
        updates.push({ linha: i + 1, valor: "" }); // Limpar Action
      }
    }

    // Batch update para limpar as células de ação
    if (updates.length > 0) {
      updates.forEach(update => {
        sheet.getRange(update.linha, CONFIG.colunas.ACTION + 1).setValue(update.valor);
      });
    }

    mostrarResultado(processedCount);
    
  } catch (error) {
    Logger.log(`❌ Erro na sincronização: ${error.message}`);
    SpreadsheetApp.getUi().alert(`❌ Erro: ${error.message}`);
  }
}

function processarAcao(acao, calendar, titulos, taskNotes, eventDescription, dadosTreino, linha) {
  try {
    switch (acao) {
      case "Add":
        return adicionarTreinoCompleto(calendar, titulos, taskNotes, eventDescription, dadosTreino, linha);
      case "Update":
        return atualizarTreinoCompleto(calendar, dadosTreino.eventId, titulos, taskNotes, eventDescription, dadosTreino);
      case "Remove":
        return removerTreinoCompleto(calendar, dadosTreino.eventId, linha);
      default:
        Logger.log(`⚠️ Ação inválida: ${acao}`);
        return { sucesso: false };
    }
  } catch (error) {
    Logger.log(`❌ Erro ao processar ação ${acao}: ${error.message}`);
    return { sucesso: false };
  }
}

function mostrarResultado(processedCount) {
  if (processedCount > 0) {
    SpreadsheetApp.getUi().alert(`✅ Sincronização concluída!\n\n${processedCount} treino(s) processado(s) no Google Tasks e Calendar.`);
  } else {
    SpreadsheetApp.getUi().alert("ℹ️ Nenhuma ação encontrada para processar.\n\nAdicione 'Add', 'Update' ou 'Remove' na coluna Action.");
  }
}

function adicionarTreinoCompleto(calendar, titulos, taskNotes, eventDescription, dadosTreino, linha) {
  try {
    const { sheet } = getSpreadsheetAndSheet();
    
    // Adiciona ao Google Tasks
    const task = Tasks.Tasks.insert({
      title: titulos.task,
      notes: taskNotes,
      due: dadosTreino.dia.toISOString()
    }, CONFIG.taskListId);

    // Adiciona ao Google Calendar
    const { inicio, fim } = criarHorarioEvento(dadosTreino.dia, dadosTreino.horario);
    const event = calendar.createEvent(titulos.event, inicio, fim, {
      description: eventDescription,
      location: "🏃‍♂️ Local de Treino"
    });
    
    // Configura evento completo (notificações + cor)
    configurarEventoCompleto(event, dadosTreino.treino);

    // Salva IDs combinados
    const combinedId = `${task.id}|${event.getId()}`;
    sheet.getRange(linha, CONFIG.colunas.EVENT_ID + 1).setValue(combinedId);
    
    Logger.log(`✅ Treino criado: ${titulos.task} | Task: ${task.id} | Event: ${event.getId()}`);
    return { sucesso: true };
    
  } catch (error) {
    Logger.log(`❌ Erro ao adicionar treino: ${error.message}`);
    return { sucesso: false };
  }
}

function atualizarTreinoCompleto(calendar, combinedId, titulos, taskNotes, eventDescription, dadosTreino) {
  if (!combinedId) {
    Logger.log("⚠️ IDs não encontrados.");
    return { sucesso: false };
  }

  try {
    const [taskId, eventId] = combinedId.split("|");

    // Atualiza Tasks
    if (taskId) {
      const task = Tasks.Tasks.get(CONFIG.taskListId, taskId);
      if (task) {
        task.title = titulos.task;
        task.notes = taskNotes;
        task.due = dadosTreino.dia.toISOString();
        Tasks.Tasks.update(task, CONFIG.taskListId, taskId);
      }
    }

    // Atualiza Calendar
    if (eventId) {
      const event = calendar.getEventById(eventId);
      if (event) {
        const { inicio, fim } = criarHorarioEvento(dadosTreino.dia, dadosTreino.horario);
        event.setTitle(titulos.event);
        event.setDescription(eventDescription);
        event.setTime(inicio, fim);
        
        // Atualiza cor baseada no novo tipo de treino
        const cor = determinarCorEvento(dadosTreino.treino);
        event.setColor(cor);
        Logger.log(`🎨 Cor atualizada: ${getCorNome(cor)} para "${dadosTreino.treino}"`);
      }
    }
    
    Logger.log(`✏️ Treino atualizado: ${titulos.task}`);
    return { sucesso: true };
    
  } catch (error) {
    Logger.log(`⚠️ Erro ao atualizar treino: ${error.message}`);
    return { sucesso: false };
  }
}

function removerTreinoCompleto(calendar, combinedId, linha) {
  if (!combinedId) {
    Logger.log("⚠️ IDs não encontrados.");
    return { sucesso: false };
  }

  try {
    const { sheet } = getSpreadsheetAndSheet();
    const [taskId, eventId] = combinedId.split("|");

    // Remove do Google Tasks
    if (taskId) {
      Tasks.Tasks.remove(CONFIG.taskListId, taskId);
      Logger.log(`✅ Task removida: ${taskId}`);
    }
    
    // Remove do Google Calendar
    if (eventId) {
      const event = calendar.getEventById(eventId);
      if (event) {
        event.deleteEvent();
        Logger.log(`✅ Evento removido: ${eventId}`);
      }
    }
    
    // Limpa os IDs e status da planilha
    sheet.getRange(linha, CONFIG.colunas.EVENT_ID + 1).setValue(""); // Limpa EventID
    
    // Verifica se o status atual deve ser limpo também
    const statusAtual = sheet.getRange(linha, CONFIG.colunas.STATUS + 1).getValue();
    if (CONFIG.status.statusParaLimpar.includes(statusAtual)) {
      sheet.getRange(linha, CONFIG.colunas.STATUS + 1).setValue(""); // Limpa Status
      Logger.log(`🧹 Status '${statusAtual}' limpo junto com a remoção do treino`);
    }
    
    Logger.log("✅ Treino completamente removido");
    return { sucesso: true };
    
  } catch (error) {
    Logger.log(`⚠️ Erro ao remover treino: ${error.message}`);
    return { sucesso: false };
  }
}

// ===== VERIFICAÇÃO AUTOMÁTICA DE STATUS ===== //

/**
 * Verifica e atualiza automaticamente os status dos treinos
 * com base no status das tarefas no Google Tasks
 */
function verificarStatusCompleto() {
  try {
    const { sheet } = getSpreadsheetAndSheet();
    const data = sheet.getDataRange().getValues();
    let updatedCount = 0;
    const updates = [];

    for (let i = 1; i < data.length; i++) {
      const combinedId = data[i][CONFIG.colunas.EVENT_ID];
      const statusAtual = data[i][CONFIG.colunas.STATUS];
      const dataEvento = data[i][CONFIG.colunas.DIA]; // Data do treino

      if (combinedId && dataEvento) {
        const novoStatus = verificarStatusTarefa(combinedId, statusAtual, dataEvento);
        // Verifica se há mudança (incluindo limpeza - string vazia)
        if (novoStatus !== null && novoStatus !== statusAtual) {
          updates.push({ linha: i + 1, status: novoStatus });
          updatedCount++;
        }
      }
    }

    // Batch update para otimizar performance
    if (updates.length > 0) {
      updates.forEach(update => {
        sheet.getRange(update.linha, CONFIG.colunas.STATUS + 1).setValue(update.status);
      });
    }

    mostrarResultadoStatus(updatedCount);
    
  } catch (error) {
    Logger.log(`❌ Erro na verificação de status: ${error.message}`);
    SpreadsheetApp.getUi().alert(`❌ Erro: ${error.message}`);
  }
}

function verificarStatusTarefa(combinedId, statusAtual, dataEvento) {
  try {
    const [taskId] = combinedId.split("|");
    
    if (!taskId) return null;
    
    const task = Tasks.Tasks.get(CONFIG.taskListId, taskId);
    const agora = new Date();
    const dataEventoObj = new Date(dataEvento);
    
    // Verifica se a tarefa foi concluída
    if (task.status === "completed" && statusAtual !== CONFIG.status.concluido) {
      Logger.log(`✅ Status atualizado para '${CONFIG.status.concluido}': ${task.title}`);
      return CONFIG.status.concluido;
    } 
    
    // Se não foi concluída, verifica se já passou do prazo
    if (task.status !== "completed") {
      // Adiciona tolerância configurável após o dia do treino
      const dataLimite = new Date(dataEventoObj);
      dataLimite.setDate(dataLimite.getDate() + CONFIG.status.diasTolerancia);
      dataLimite.setHours(23, 59, 59, 999); // Final do dia limite
      
      const statusProtegidos = [CONFIG.status.expirado, CONFIG.status.concluido];
      
      if (agora > dataLimite && !statusProtegidos.includes(statusAtual)) {
        Logger.log(`⏰ Tarefa expirada: ${task.title} | Data limite: ${formatarDataCompleta(dataLimite)}`);
        return CONFIG.status.expirado;
      } else if (agora <= dataLimite && ![CONFIG.status.pendente, ...statusProtegidos].includes(statusAtual)) {
        Logger.log(`📅 Status atualizado para '${CONFIG.status.pendente}': ${task.title}`);
        return CONFIG.status.pendente;
      }
    }
    
    return null; // Sem mudança
    
  } catch (error) {
    // Verifica se o erro indica que a tarefa foi removida/não existe
    const erroRemocao = error.message.includes('notFound') || 
                       error.message.includes('404') || 
                       error.message.includes('Not Found') ||
                       error.message.includes('does not exist');
                       
    if (erroRemocao) {
      // Se a tarefa não existe e o status atual está na lista para limpar
      if (CONFIG.status.statusParaLimpar.includes(statusAtual)) {
        Logger.log(`🗑️ Tarefa removida - limpando status '${statusAtual}' | ID: ${combinedId}`);
        return ""; // Retorna string vazia para limpar a célula
      }
      // Se o status atual já é vazio ou não está na lista para limpar, não faz nada
      return null;
    }
    
    // Para outros tipos de erro, mantém o comportamento anterior
    Logger.log(`⚠️ Erro ao buscar tarefa ${combinedId}: ${error.message}`);
    return statusAtual !== CONFIG.status.naoEncontrado ? CONFIG.status.naoEncontrado : null;
  }
}

function mostrarResultadoStatus(updatedCount) {
  if (updatedCount > 0) {
    SpreadsheetApp.getUi().alert(`✅ Status atualizado!\n\n${updatedCount} status(es) foram atualizados.`);
  } else {
    SpreadsheetApp.getUi().alert("ℹ️ Todos os status já estão atualizados.");
  }
}

// ===== FUNÇÕES DE CONFIGURAÇÃO INICIAL ===== //

/**
 * FUNÇÃO DE CONFIGURAÇÃO INICIAL
 * Execute esta função UMA VEZ para descobrir o ID da sua lista de tarefas
 * Copie o ID que aparecer no log e cole em CONFIG.taskListId
 */
function obterTaskListId() {
  const taskLists = Tasks.Tasklists.list();
  if (taskLists.items && taskLists.items.length > 0) {
    Logger.log("🔍 === SUAS LISTAS DE TAREFAS ===");
    taskLists.items.forEach(taskList => {
      Logger.log(`📋 Nome: "${taskList.title}" | 🆔 ID: "${taskList.id}"`);
    });
    Logger.log("📝 Copie o ID da lista que deseja usar e cole em CONFIG.taskListId");
  } else {
    Logger.log("❌ Nenhuma lista de tarefas encontrada. Crie uma lista no Google Tasks primeiro.");
  }
}

/**
 * FUNÇÃO DE DEBUG
 * Execute para verificar se a estrutura da planilha está correta
 */
function debugEstruturaPlanilha() {
  try {
    const { sheet } = getSpreadsheetAndSheet();
    const data = sheet.getDataRange().getValues();
    
    Logger.log("🔍 === DEBUG ESTRUTURA DA PLANILHA ===");
    Logger.log(`📊 Total de linhas: ${data.length}`);
    Logger.log(`📊 Total de colunas: ${data[0] ? data[0].length : 0}`);
    
    // Mostra o cabeçalho (linha 1)
    if (data.length > 0) {
      Logger.log("📋 CABEÇALHO (Linha 1):");
      data[0].forEach((header, index) => {
        const coluna = String.fromCharCode(65 + index); // A=0, B=1, etc.
        Logger.log(`   ${coluna} (${index}): "${header}"`);
      });
    }
    
    // Mostra algumas linhas de dados
    Logger.log("\n📋 PRIMEIRAS LINHAS DE DADOS:");
    for (let i = 1; i < Math.min(data.length, 4); i++) {
      Logger.log(`\n📍 Linha ${i + 1}:`);
      data[i].forEach((cell, index) => {
        const coluna = String.fromCharCode(65 + index);
        Logger.log(`   ${coluna} (${index}): "${cell}" (${typeof cell})`);
      });
    }
    
    // Mostra configuração atual das colunas
    Logger.log("\n⚙️ CONFIGURAÇÃO ATUAL DAS COLUNAS:");
    Object.entries(CONFIG.colunas).forEach(([nome, indice]) => {
      const coluna = String.fromCharCode(65 + indice);
      Logger.log(`   ${nome}: ${coluna} (índice ${indice})`);
    });
    
    Logger.log("🔍 === FIM DEBUG ===");
    
  } catch (error) {
    Logger.log(`❌ Erro no debug: ${error.message}`);
  }
}

// ===== INSTRUÇÕES DE USO ===== //
/**
 * 📚 COMO USAR ESTE SCRIPT:
 * 
 * 1. CONFIGURAÇÃO INICIAL:
 *    - Edite as configurações em CONFIG
 *    - Execute obterTaskListId() para descobrir seu taskListId
 *    - Ative as APIs necessárias no Google Apps Script
 * 
 * 2. ESTRUTURA DA PLANILHA:
 *    A | B      | C      | D   | E       | F           | G      | H              | I          | J          | K
 *    --+--------+--------+-----+---------+-------------+--------+----------------+------------+------------+-------
 *    A | Status | Action | Dia | Horário | Dia Semana | Treino | Zona Frequência| Descrição  | Pós Treino | EventID
 * 
 * 3. USANDO O SCRIPT:
 *    - Na coluna Action (C), digite: "Add", "Update" ou "Remove"
 *    - Na coluna Horário (E), use formato 24h: "06:00", "18:30", etc.
 *    - Execute sincronizarTreinosCompleto() ou use o menu
 * 
 * 4. RECURSOS:
 *    ✅ Horários dinâmicos por treino
 *    ✅ Cores inteligentes por tipo de treino
 *    ✅ Notificações personalizáveis
 *    ✅ Sincronização bidirecional de status
 *    ✅ Templates personalizáveis
 * 
 * 🔗 Mais informações: https://github.com/seu-usuario/seu-repo
 */
