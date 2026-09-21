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
app.use(express.json());

// ============================================================
// CONFIGURAÇÃO
// ============================================================

const API_URL =
  "https://integration.ajin.io/v3/query-inss-balances/finder/await";

const API_KEY = process.env.API_KEY || "";

// Intervalo entre consultas
const INTERVALO = 1000;

// Tempo máximo de uma tarefa na memória
const TEMPO_TAREFA = 60 * 60 * 1000;

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
    dados,
    enviadoEm: Date.now()
  };

  for (const cliente of [...tarefa.clientes]) {
    try {
      cliente.res.write(`event: ${tipo}\n`);
      cliente.res.write(
        `data: ${JSON.stringify(dados)}\n\n`
      );
    } catch (erro) {
      console.error(
        `Erro enviando SSE da tarefa ${tarefa.id}:`,
        erro.message
      );

      removerCliente(tarefa, cliente);
    }
  }
}

// ============================================================
// REMOVER CLIENTE SSE
// ============================================================

function removerCliente(tarefa, cliente) {
  const indice =
    tarefa.clientes.indexOf(cliente);

  if (indice !== -1) {
    tarefa.clientes.splice(indice, 1);
  }

  if (cliente.heartbeat) {
    clearInterval(cliente.heartbeat);
    cliente.heartbeat = null;
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

    const margem =
      response.data?.consignedCreditBalance;

    return {
      sucesso: true,
      statusHTTP: response.status,
      blockType:
        response.data?.blockType || null,
      margem:
        margem !== undefined &&
        margem !== null
          ? Number(margem)
          : null,
      mensagem: null
    };

  } catch (error) {

    if (error.response) {

      const mensagem =
        error.response?.data?.messages?.[0]?.text ||
        null;

      return {
        sucesso: false,
        statusHTTP: error.response.status,
        blockType: null,
        margem: null,
        mensagem:
          mensagem ||
          `HTTP ${error.response.status}`
      };
    }

    return {
      sucesso: false,
      statusHTTP: null,
      blockType: null,
      margem: null,
      mensagem:
        error.code ||
        error.message ||
        "Erro desconhecido"
    };
  }
}

// ============================================================
// DEFINIR STATUS
// ============================================================

function definirStatus(resultado) {

  if (
    resultado.blockType ===
    "not_blocked"
  ) {
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

  const worksheet =
    XLSX.utils.json_to_sheet(resultados);

  worksheet["!cols"] = [
    { wch: 18 },
    { wch: 18 },
    { wch: 28 },
    { wch: 18 }
  ];

  if (worksheet["!ref"]) {

    const range =
      XLSX.utils.decode_range(
        worksheet["!ref"]
      );

    // Coluna D = margem
    // Aceita valores positivos e negativos

    for (
      let linha = 1;
      linha <= range.e.r;
      linha++
    ) {

      const celula =
        worksheet[
          XLSX.utils.encode_cell({
            r: linha,
            c: 3
          })
        ];

      if (
        celula &&
        typeof celula.v === "number"
      ) {

        celula.z =
          'R$ #,##0.00;[Red]-R$ #,##0.00';
      }
    }
  }

  const novoWorkbook =
    XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(
    novoWorkbook,
    worksheet,
    "Resultado"
  );

  return XLSX.write(
    novoWorkbook,
    {
      bookType: "xlsx",
      type: "buffer"
    }
  );
}

// ============================================================
// EXECUTAR PROCESSAMENTO
// ============================================================

async function executarTarefa(
  tarefa,
  dados
) {

  try {

    const resultados = [];

    // --------------------------------------------------------
    // DESCOBRIR CLIENTES VÁLIDOS
    // --------------------------------------------------------

    const clientes = [];

    for (
      let indice = 1;
      indice < dados.length;
      indice++
    ) {

      const linha =
        dados[indice] || [];

      const cpfFormatado =
        formatarCPF(linha[0]);

      const beneficioFormatado =
        formatarSegundaColuna(linha[1]);

      const cpf =
        somenteNumeros(cpfFormatado);

      const beneficio =
        somenteNumeros(
          beneficioFormatado
        );

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

    const total =
      clientes.length;

    tarefa.total = total;

    // --------------------------------------------------------
    // NENHUM CLIENTE
    // --------------------------------------------------------

    if (!total) {

      tarefa.erro =
        "Nenhum cliente válido foi encontrado. Verifique se a planilha possui CPF na primeira coluna e benefício na segunda.";

      enviarEvento(
        tarefa,
        "erro",
        {
          mensagem: tarefa.erro
        }
      );

      return;
    }

    // --------------------------------------------------------
    // INFORMAR TOTAL
    // --------------------------------------------------------

    enviarEvento(
      tarefa,
      "inicio",
      {
        total
      }
    );

    // --------------------------------------------------------
    // CONSULTAR CLIENTES
    // --------------------------------------------------------

    for (
      let indice = 0;
      indice < clientes.length;
      indice++
    ) {

      const cliente =
        clientes[indice];

      const numeroAtual =
        indice + 1;

      console.log(
        `[${numeroAtual}/${total}] Consultando CPF ${cliente.cpf} | benefício ${cliente.beneficio}`
      );

      // ------------------------------------------------------
      // AVISAR QUE ESTÁ CONSULTANDO
      // ------------------------------------------------------

      enviarEvento(
        tarefa,
        "consultando",
        {
          atual: tarefa.processados,
          consultando: numeroAtual,
          total,
          percentual:
            Math.round(
              (tarefa.processados / total) *
              100
            )
        }
      );

      // ------------------------------------------------------
      // CONSULTA IN100
      // ------------------------------------------------------

      const resultado =
        await consultarIN100(
          cliente.cpf,
          cliente.beneficio
        );

      const status =
        definirStatus(resultado);

      const margemValida =
        resultado.margem !== null &&
        Number.isFinite(
          resultado.margem
        );

      resultados.push({
        CPF: cliente.cpfFormatado,

        BENEFICIO:
          cliente.beneficioFormatado,

        "STATUS DO BENEFICIO":
          status,

        MARGEM:
          margemValida
            ? resultado.margem
            : ""
      });

      // ------------------------------------------------------
      // CONSULTA CONCLUÍDA
      // ------------------------------------------------------

      tarefa.processados =
        numeroAtual;

      const percentual =
        Math.round(
          (tarefa.processados / total) *
          100
        );

      console.log(
        `   ${status} | margem: ${
          margemValida
            ? resultado.margem
            : "N/A"
        }`
      );

      enviarEvento(
        tarefa,
        "progresso",
        {
          atual:
            tarefa.processados,

          total,

          percentual
        }
      );

      // ------------------------------------------------------
      // INTERVALO
      // ------------------------------------------------------

      if (
        indice <
        clientes.length - 1
      ) {
        await esperar(
          INTERVALO
        );
      }
    }

    // --------------------------------------------------------
    // GERAR EXCEL
    // --------------------------------------------------------

    enviarEvento(
      tarefa,
      "finalizando",
      {
        atual: total,
        total
      }
    );

    const arquivoSaida =
      gerarExcel(resultados);

    tarefa.arquivo =
      arquivoSaida;

    tarefa.concluida =
      true;

    tarefa.processados =
      total;

    // --------------------------------------------------------
    // CONCLUÍDO
    // --------------------------------------------------------

    enviarEvento(
      tarefa,
      "concluido",
      {
        atual: total,
        total,
        percentual: 100,
        download:
          `/api/download/${tarefa.id}`
      }
    );

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

    enviarEvento(
      tarefa,
      "erro",
      {
        mensagem:
          tarefa.erro
      }
    );
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
          erro:
            "Envie uma planilha."
        });
      }

      if (!API_KEY) {

        return res.status(500).json({
          erro:
            "API_KEY não configurada. Configure a variável de ambiente API_KEY antes de iniciar o programa."
        });
      }

      const workbook =
        XLSX.read(
          req.file.buffer,
          {
            type: "buffer",
            cellDates: false,
            raw: true
          }
        );

      const nomeAba =
        workbook.SheetNames[0];

      const sheet =
        workbook.Sheets[nomeAba];

      const dados =
        XLSX.utils.sheet_to_json(
          sheet,
          {
            header: 1,
            defval: "",
            raw: true
          }
        );

      if (!dados.length) {

        return res.status(400).json({
          erro:
            "A planilha está vazia."
        });
      }

      // ------------------------------------------------------
      // CRIAR ID
      // ------------------------------------------------------

      const id =
        crypto.randomUUID();

      const tarefa = {

        id,

        total: 0,

        processados: 0,

        concluida: false,

        erro: null,

        arquivo: null,

        clientes: [],

        ultimoEvento: null,

        criadaEm:
          Date.now(),

        ultimaAtualizacao:
          Date.now()
      };

      tarefas.set(
        id,
        tarefa
      );

      // ------------------------------------------------------
      // RESPONDER IMEDIATAMENTE
      // ------------------------------------------------------

      res.json({
        sucesso: true,
        tarefaId: id
      });

      // ------------------------------------------------------
      // PROCESSAMENTO EM SEGUNDO PLANO
      // ------------------------------------------------------

      executarTarefa(
        tarefa,
        dados
      ).catch(erro => {

        console.error(
          `Erro não tratado na tarefa ${id}:`,
          erro
        );

      });

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
// SSE - PROGRESSO
// ============================================================

app.get(
  "/api/progresso/:id",
  (req, res) => {

    const tarefa =
      tarefas.get(
        req.params.id
      );

    if (!tarefa) {

      return res.status(404).json({
        erro:
          "Tarefa não encontrada."
      });
    }

    res.setHeader(
      "Content-Type",
      "text/event-stream"
    );

    res.setHeader(
      "Cache-Control",
      "no-cache, no-transform"
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
      res,
      heartbeat: null
    };

    tarefa.clientes.push(
      cliente
    );

    // --------------------------------------------------------
    // ÚLTIMO EVENTO
    // --------------------------------------------------------

    if (
      tarefa.ultimoEvento
    ) {

      try {

        res.write(
          `event: ${tarefa.ultimoEvento.tipo}\n`
        );

        res.write(
          `data: ${JSON.stringify(
            tarefa.ultimoEvento.dados
          )}\n\n`
        );

      } catch (_) {}
    }

    // --------------------------------------------------------
    // SE JÁ CONCLUIU
    // --------------------------------------------------------

    if (
      tarefa.concluida
    ) {

      try {

        res.write(
          `event: concluido\n`
        );

        res.write(
          `data: ${JSON.stringify({
            atual:
              tarefa.total,

            total:
              tarefa.total,

            percentual: 100,

            download:
              `/api/download/${tarefa.id}`
          })}\n\n`
        );

      } catch (_) {}
    }

    // --------------------------------------------------------
    // SE DEU ERRO
    // --------------------------------------------------------

    if (tarefa.erro) {

      try {

        res.write(
          `event: erro\n`
        );

        res.write(
          `data: ${JSON.stringify({
            mensagem:
              tarefa.erro
          })}\n\n`
        );

      } catch (_) {}
    }

    // --------------------------------------------------------
    // HEARTBEAT
    // --------------------------------------------------------

    cliente.heartbeat =
      setInterval(() => {

        try {

          res.write(
            ": heartbeat\n\n"
          );

        } catch (_) {

          removerCliente(
            tarefa,
            cliente
          );
        }

      }, 15000);

    // --------------------------------------------------------
    // DESCONEXÃO
    // --------------------------------------------------------

    req.on(
      "close",
      () => {

        removerCliente(
          tarefa,
          cliente
        );

      }
    );
  }
);

// ============================================================
// STATUS DA TAREFA
// ============================================================

app.get(
  "/api/status-tarefa/:id",
  (req, res) => {

    const tarefa =
      tarefas.get(
        req.params.id
      );

    if (!tarefa) {

      return res.status(404).json({
        encontrada: false,
        erro:
          "Tarefa não encontrada."
      });
    }

    res.json({
      encontrada: true,

      id:
        tarefa.id,

      total:
        tarefa.total,

      processados:
        tarefa.processados,

      percentual:
        tarefa.total > 0
          ? Math.round(
              (tarefa.processados /
                tarefa.total) *
                100
            )
          : 0,

      concluida:
        tarefa.concluida,

      erro:
        tarefa.erro,

      criadaEm:
        tarefa.criadaEm,

      ultimaAtualizacao:
        tarefa.ultimaAtualizacao,

      download:
        tarefa.concluida
          ? `/api/download/${tarefa.id}`
          : null
    });
  }
);

// ============================================================
// DOWNLOAD
// ============================================================

app.get(
  "/api/download/:id",
  (req, res) => {

    const tarefa =
      tarefas.get(
        req.params.id
      );

    if (!tarefa) {

      return res.status(404).json({
        erro:
          "Tarefa não encontrada."
      });
    }

    if (
      !tarefa.concluida ||
      !tarefa.arquivo
    ) {

      return res.status(400).json({
        erro:
          "O processamento ainda não terminou."
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

    res.send(
      tarefa.arquivo
    );
  }
);

// ============================================================
// STATUS DA API
// ============================================================

app.get(
  "/api/status",
  (req, res) => {

    res.json({
      online: true,

      versao:
        "Beta 3 - Formatador + IN100 + Progresso",

      apiConfigurada:
        Boolean(API_KEY),

      tarefasAtivas:
        tarefas.size
    });
  }
);

// ============================================================
// LIMPEZA DE TAREFAS ANTIGAS
// ============================================================

setInterval(
  () => {

    const agora =
      Date.now();

    for (
      const [id, tarefa]
      of tarefas.entries()
    ) {

      if (
        agora -
          tarefa.criadaEm >
        TEMPO_TAREFA
      ) {

        console.log(
          `Removendo tarefa antiga: ${id}`
        );

        // Fecha possíveis conexões SSE
        for (
          const cliente
          of [...tarefa.clientes]
        ) {

          try {
            cliente.res.end();
          } catch (_) {}

          removerCliente(
            tarefa,
            cliente
          );
        }

        tarefas.delete(id);
      }
    }

  },
  10 * 60 * 1000
);

// ============================================================
// INICIAR SERVIDOR
// ============================================================

app.listen(
  PORT,
  () => {

    console.log(
      `Beta 3 - Formatador + IN100 + Progresso rodando na porta ${PORT}`
    );

    if (!API_KEY) {

      console.log(
        "⚠️ API_KEY não configurada."
      );

    } else {

      console.log(
        "✅ API_KEY configurada."
      );
    }
  }
);
