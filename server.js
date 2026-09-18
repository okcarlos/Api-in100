const express = require("express");
const multer = require("multer");
const XLSX = require("xlsx");
const axios = require("axios");
const path = require("path");
const cors = require("cors");
const crypto = require("crypto");

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());

// ============================================================
// CONFIGURAÇÃO
// ============================================================

const API_URL =
  "https://integration.ajin.io/v3/query-inss-balances/finder/await";

const API_KEY = process.env.API_KEY || "";

// Intervalo entre consultas
const INTERVALO = 1000;

const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 20 * 1024 * 1024
  }
});

app.use(express.static(path.join(__dirname, "public")));

// ============================================================
// TAREFAS EM MEMÓRIA
// ============================================================

const tarefas = new Map();

// ============================================================
// UTILITÁRIOS
// ============================================================

function somenteNumeros(valor) {
  return String(valor ?? "").replace(/\D/g, "");
}

function formatarCPF(valor) {
  let numero = somenteNumeros(valor);

  if (!numero) return "";

  numero = numero.padStart(11, "0");

  if (numero.length > 11) {
    numero = numero.slice(-11);
  }

  return `${numero.slice(0, 3)}.${numero.slice(3, 6)}.${numero.slice(6, 9)}-${numero.slice(9, 11)}`;
}

function formatarSegundaColuna(valor) {
  let numero = somenteNumeros(valor);

  if (!numero) return "";

  numero = numero.padStart(10, "0");

  if (numero.length > 10) {
    numero = numero.slice(-10);
  }

  return numero;
}

function esperar(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

// ============================================================
// ENVIAR EVENTO SSE
// ============================================================

function enviarEvento(tarefa, tipo, dados) {
  tarefa.ultimoEvento = {
    tipo,
    dados
  };

  for (const cliente of tarefa.clientes) {
    cliente.res.write(`event: ${tipo}\n`);
    cliente.res.write(`data: ${JSON.stringify(dados)}\n\n`);
  }
}

// ============================================================
// CONSULTAR IN100
// ============================================================

async function consultarIN100(cpf, beneficio) {
  if (!API_KEY) {
    return {
      sucesso: false,
      statusHTTP: null,
      blockType: null,
      margem: null,
      mensagem: "API_KEY não configurada."
    };
  }

  try {
    const response = await axios.post(
      API_URL,
      {
        identity: cpf,
        benefitNumber: beneficio,
        lastHours: 1,
        timeout: 120
      },
      {
        headers: {
          "Content-Type": "application/json",
          "Accept": "application/json",
          "apikey": API_KEY
        },
        timeout: 180000
      }
    );

    const margem = response.data?.consignedCreditBalance;

    return {
      sucesso: true,
      statusHTTP: response.status,
      blockType: response.data?.blockType || null,
      margem:
        margem !== undefined && margem !== null
          ? Number(margem)
          : null,
      mensagem: null
    };
  } catch (error) {
    if (error.response) {
      const mensagem =
        error.response?.data?.messages?.[0]?.text || null;

      return {
        sucesso: false,
        statusHTTP: error.response.status,
        blockType: null,
        margem: null,
        mensagem: mensagem || `HTTP ${error.response.status}`
      };
    }

    return {
      sucesso: false,
      statusHTTP: null,
      blockType: null,
      margem: null,
      mensagem: error.code || error.message
    };
  }
}

// ============================================================
// DEFINIR STATUS
// ============================================================

function definirStatus(resultado) {
  if (resultado.blockType === "not_blocked") {
    return "DESBLOQUEADO";
  }

  if (resultado.blockType) {
    return "BLOQUEADO";
  }

  if (resultado.sucesso) {
    return "DESBLOQUEADO";
  }

  if (resultado.mensagem) {
    return resultado.mensagem;
  }

  return "ERRO";
}

// ============================================================
// GERAR EXCEL
// ============================================================

function gerarExcel(resultados) {
  const worksheet = XLSX.utils.json_to_sheet(resultados);

  worksheet["!cols"] = [
    { wch: 18 },
    { wch: 18 },
    { wch: 28 },
    { wch: 18 }
  ];

  const range = XLSX.utils.decode_range(worksheet["!ref"]);

  // Coluna D = margem
  // Aceita valores positivos e negativos
  for (let linha = 1; linha <= range.e.r; linha++) {
    const celula = worksheet[
      XLSX.utils.encode_cell({
        r: linha,
        c: 3
      })
    ];

    if (celula && typeof celula.v === "number") {
      celula.z = 'R$ #,##0.00;[Red]-R$ #,##0.00';
    }
  }

  const novoWorkbook = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(
    novoWorkbook,
    worksheet,
    "Resultado"
  );

  return XLSX.write(novoWorkbook, {
    bookType: "xlsx",
    type: "buffer"
  });
}

// ============================================================
// EXECUTAR PROCESSAMENTO
// ============================================================

async function executarTarefa(tarefa, dados) {
  try {
    const resultados = [];

    // --------------------------------------------------------
    // PRIMEIRO: descobrir clientes válidos
    // --------------------------------------------------------

    const clientes = [];

    for (let indice = 1; indice < dados.length; indice++) {
      const linha = dados[indice] || [];

      const cpfFormatado = formatarCPF(linha[0]);
      const beneficioFormatado = formatarSegundaColuna(linha[1]);

      const cpf = somenteNumeros(cpfFormatado);
      const beneficio = somenteNumeros(beneficioFormatado);

      if (!cpf || !beneficio) {
        continue;
      }

      clientes.push({
        cpf,
        beneficio,
        cpfFormatado,
        beneficioFormatado
      });
    }

    const total = clientes.length;

    tarefa.total = total;

    if (!total) {
      tarefa.erro =
        "Nenhum cliente válido foi encontrado. Verifique se a planilha possui CPF na primeira coluna e benefício na segunda.";

      enviarEvento(tarefa, "erro", {
        mensagem: tarefa.erro
      });

      return;
    }

    // --------------------------------------------------------
    // INFORMAR TOTAL
    // --------------------------------------------------------

    enviarEvento(tarefa, "inicio", {
      total
    });

    // --------------------------------------------------------
    // CONSULTAR CLIENTES
    // --------------------------------------------------------

    for (let indice = 0; indice < clientes.length; indice++) {
      const cliente = clientes[indice];

      const numeroAtual = indice + 1;

      tarefa.processados = numeroAtual;

      console.log(
        `[${numeroAtual}/${total}] Consultando CPF ${cliente.cpf} | benefício ${cliente.beneficio}`
      );

      // Atualiza a interface antes da consulta
      enviarEvento(tarefa, "progresso", {
        atual: numeroAtual,
        total,
        percentual: Math.round((numeroAtual / total) * 100)
      });

      const resultado = await consultarIN100(
        cliente.cpf,
        cliente.beneficio
      );

      const status = definirStatus(resultado);

      const margemValida =
        resultado.margem !== null &&
        Number.isFinite(resultado.margem);

      resultados.push({
        CPF: cliente.cpfFormatado,
        BENEFICIO: cliente.beneficioFormatado,
        "STATUS DO BENEFICIO": status,
        MARGEM: margemValida
          ? resultado.margem
          : ""
      });

      console.log(
        `   ${status} | margem: ${
          margemValida
            ? resultado.margem
            : "N/A"
        }`
      );

      // Espera entre consultas
      if (indice < clientes.length - 1) {
        await esperar(INTERVALO);
      }
    }

    // --------------------------------------------------------
    // GERAR EXCEL
    // --------------------------------------------------------

    enviarEvento(tarefa, "finalizando", {
      atual: total,
      total
    });

    const arquivoSaida = gerarExcel(resultados);

    tarefa.arquivo = arquivoSaida;
    tarefa.concluida = true;

    // --------------------------------------------------------
    // CONCLUÍDO
    // --------------------------------------------------------

    enviarEvento(tarefa, "concluido", {
      atual: total,
      total,
      percentual: 100,
      download: `/api/download/${tarefa.id}`
    });

    console.log(
      `Tarefa ${tarefa.id} concluída: ${total} clientes.`
    );

  } catch (erro) {
    console.error(
      `Erro na tarefa ${tarefa.id}:`,
      erro
    );

    tarefa.erro =
      erro.message ||
      "Não foi possível processar a planilha.";

    enviarEvento(tarefa, "erro", {
      mensagem: tarefa.erro
    });
  }
}

// ============================================================
// CRIAR PROCESSAMENTO
// ============================================================

app.post(
  "/api/processar",
  upload.single("arquivo"),
  async (req, res) => {
    try {
      if (!req.file) {
        return res.status(400).json({
          erro: "Envie uma planilha."
        });
      }

      if (!API_KEY) {
        return res.status(500).json({
          erro:
            "API_KEY não configurada. Configure a variável de ambiente API_KEY antes de iniciar o programa."
        });
      }

      const workbook = XLSX.read(req.file.buffer, {
        type: "buffer",
        cellDates: false,
        raw: true
      });

      const nomeAba = workbook.SheetNames[0];

      const sheet = workbook.Sheets[nomeAba];

      const dados = XLSX.utils.sheet_to_json(
        sheet,
        {
          header: 1,
          defval: "",
          raw: true
        }
      );

      if (!dados.length) {
        return res.status(400).json({
          erro: "A planilha está vazia."
        });
      }

      // ------------------------------------------------------
      // CRIAR ID DA TAREFA
      // ------------------------------------------------------

      const id = crypto.randomUUID();

      const tarefa = {
        id,
        total: 0,
        processados: 0,
        concluida: false,
        erro: null,
        arquivo: null,
        clientes: [],
        ultimoEvento: null,
        criadaEm: Date.now()
      };

      tarefas.set(id, tarefa);

      // ------------------------------------------------------
      // RESPONDER IMEDIATAMENTE
      // ------------------------------------------------------

      res.json({
        sucesso: true,
        tarefaId: id
      });

      // ------------------------------------------------------
      // PROCESSAR EM SEGUNDO PLANO
      // ------------------------------------------------------

      executarTarefa(tarefa, dados);

    } catch (erro) {
      console.error(erro);

      return res.status(500).json({
        erro:
          erro.message ||
          "Não foi possível iniciar o processamento."
      });
    }
  }
);

// ============================================================
// SSE - PROGRESSO EM TEMPO REAL
// ============================================================

app.get("/api/progresso/:id", (req, res) => {
  const tarefa = tarefas.get(req.params.id);

  if (!tarefa) {
    return res.status(404).json({
      erro: "Tarefa não encontrada."
    });
  }

  res.setHeader(
    "Content-Type",
    "text/event-stream"
  );

  res.setHeader(
    "Cache-Control",
    "no-cache"
  );

  res.setHeader(
    "Connection",
    "keep-alive"
  );

  res.setHeader(
    "X-Accel-Buffering",
    "no"
  );

  res.flushHeaders();

  const cliente = {
    res
  };

  tarefa.clientes.push(cliente);

  // ----------------------------------------------------------
  // Mandar último evento para quem acabou de conectar
  // ----------------------------------------------------------

  if (tarefa.ultimoEvento) {
    res.write(
      `event: ${tarefa.ultimoEvento.tipo}\n`
    );

    res.write(
      `data: ${JSON.stringify(
        tarefa.ultimoEvento.dados
      )}\n\n`
    );
  }

  // ----------------------------------------------------------
  // Se já terminou antes do SSE conectar
  // ----------------------------------------------------------

  if (tarefa.concluida) {
    res.write(
      `event: concluido\n`
    );

    res.write(
      `data: ${JSON.stringify({
        atual: tarefa.total,
        total: tarefa.total,
        percentual: 100,
        download: `/api/download/${tarefa.id}`
      })}\n\n`
    );
  }

  if (tarefa.erro) {
    res.write(
      `event: erro\n`
    );

    res.write(
      `data: ${JSON.stringify({
        mensagem: tarefa.erro
      })}\n\n`
    );
  }

  // ----------------------------------------------------------
  // Heartbeat
  // ----------------------------------------------------------

  const intervaloHeartbeat = setInterval(() => {
    try {
      res.write(": heartbeat\n\n");
    } catch (_) {}
  }, 15000);

  // ----------------------------------------------------------
  // Desconexão
  // ----------------------------------------------------------

  req.on("close", () => {
    clearInterval(intervaloHeartbeat);

    const indice = tarefa.clientes.indexOf(cliente);

    if (indice !== -1) {
      tarefa.clientes.splice(indice, 1);
    }
  });
});

// ============================================================
// DOWNLOAD
// ============================================================

app.get("/api/download/:id", (req, res) => {
  const tarefa = tarefas.get(req.params.id);

  if (!tarefa) {
    return res.status(404).json({
      erro: "Tarefa não encontrada."
    });
  }

  if (!tarefa.concluida || !tarefa.arquivo) {
    return res.status(400).json({
      erro: "O processamento ainda não terminou."
    });
  }

  res.setHeader(
    "Content-Disposition",
    'attachment; filename="resultado_in100.xlsx"'
  );

  res.setHeader(
    "Content-Type",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  );

  res.send(tarefa.arquivo);
});

// ============================================================
// STATUS DA API
// ============================================================

app.get("/api/status", (req, res) => {
  res.json({
    online: true,
    versao: "Beta 3 - Formatador + IN100 + Progresso",
    apiConfigurada: Boolean(API_KEY)
  });
});

// ============================================================
// LIMPEZA DE TAREFAS ANTIGAS
// ============================================================

setInterval(() => {
  const agora = Date.now();

  for (const [id, tarefa] of tarefas.entries()) {
    // Apaga tarefas com mais de 1 hora
    if (agora - tarefa.criadaEm > 60 * 60 * 1000) {
      tarefas.delete(id);
    }
  }
}, 10 * 60 * 1000);

// ============================================================
// INICIAR SERVIDOR
// ============================================================

app.listen(PORT, () => {
  console.log(
    `Beta 3 - Formatador + IN100 + Progresso rodando na porta ${PORT}`
  );

  if (!API_KEY) {
    console.log("⚠️ API_KEY não configurada.");
  } else {
    console.log("✅ API_KEY configurada.");
  }
});
