const express = require("express");
const multer = require("multer");
const XLSX = require("xlsx");
const axios = require("axios");
const path = require("path");

const app = express();
const PORT = process.env.PORT || 3000;

// ============================================================
// CONFIGURAÇÃO
// ============================================================

const API_URL =
  "https://integration.ajin.io/v3/query-inss-balances/finder/await";

// Por segurança, a API Key deve ser configurada como variável de ambiente.
// PowerShell:
//   $env:API_KEY="SUA_CHAVE_AQUI"
// const API_KEY = process.env.API_KEY || "";
const API_KEY = "XNgTlkeYbqjEf4A07bqMoqp5FaFiGEmu8WDgeIF4Oa10rLQKP6u18Y/mio552RH/";

// Intervalo entre consultas.
const INTERVALO = 1000;

const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 20 * 1024 * 1024
  }
});

app.use(express.static(path.join(__dirname, "public")));

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
// PROCESSAR PLANILHA + CONSULTAR IN100
// ============================================================

app.post("/api/processar", upload.single("arquivo"), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({
        erro: "Envie uma planilha."
      });
    }

    if (!API_KEY) {
      return res.status(500).json({
        erro: "API_KEY não configurada. Configure a variável de ambiente API_KEY antes de iniciar o programa."
      });
    }

    const workbook = XLSX.read(req.file.buffer, {
      type: "buffer",
      cellDates: false,
      raw: true
    });

    const nomeAba = workbook.SheetNames[0];
    const sheet = workbook.Sheets[nomeAba];

    const dados = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      defval: "",
      raw: true
    });

    if (!dados.length) {
      return res.status(400).json({
        erro: "A planilha está vazia."
      });
    }

    const resultados = [];
    let processados = 0;

    // A primeira linha é o cabeçalho original e não é consultada.
    for (let indice = 1; indice < dados.length; indice++) {
      const linha = dados[indice] || [];

      const cpfFormatado = formatarCPF(linha[0]);
      const beneficioFormatado = formatarSegundaColuna(linha[1]);

      const cpf = somenteNumeros(cpfFormatado);
      const beneficio = somenteNumeros(beneficioFormatado);

      if (!cpf || !beneficio) {
        continue;
      }

      processados++;

      console.log(
        `[${processados}] Consultando CPF ${cpf} | benefício ${beneficio}`
      );

      const resultado = await consultarIN100(cpf, beneficio);
      const status = definirStatus(resultado);

      const margemValida =
        resultado.margem !== null &&
        Number.isFinite(resultado.margem);

      resultados.push({
        CPF: cpfFormatado,
        BENEFICIO: beneficioFormatado,
        "STATUS DO BENEFICIO": status,
        MARGEM: margemValida ? resultado.margem : ""
      });

      console.log(
        `   ${status} | margem: ${margemValida ? resultado.margem : "N/A"}`
      );

      if (indice < dados.length - 1) {
        await esperar(INTERVALO);
      }
    }

    if (!resultados.length) {
      return res.status(400).json({
        erro: "Nenhum cliente válido foi encontrado. Verifique se a planilha possui CPF na primeira coluna e benefício na segunda."
      });
    }

    // ========================================================
    // CRIAR EXCEL FINAL
    // ========================================================

    const worksheet = XLSX.utils.json_to_sheet(resultados);

    worksheet["!cols"] = [
      { wch: 18 },
      { wch: 18 },
      { wch: 28 },
      { wch: 18 }
    ];

    const range = XLSX.utils.decode_range(worksheet["!ref"]);

    // Coluna D = margem. Aceita positivos e negativos.
    for (let linha = 1; linha <= range.e.r; linha++) {
      const celula = worksheet[
        XLSX.utils.encode_cell({ r: linha, c: 3 })
      ];

      if (celula && typeof celula.v === "number") {
        celula.z = 'R$ #,##0.00;[Red]-R$ #,##0.00';
      }
    }

    const novoWorkbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(novoWorkbook, worksheet, "Resultado");

    const arquivoSaida = XLSX.write(novoWorkbook, {
      bookType: "xlsx",
      type: "buffer"
    });

    res.setHeader(
      "Content-Disposition",
      'attachment; filename="resultado_in100.xlsx"'
    );
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );

    return res.send(arquivoSaida);
  } catch (erro) {
    console.error(erro);

    return res.status(500).json({
      erro: erro.message || "Não foi possível processar a planilha."
    });
  }
});

app.get("/api/status", (req, res) => {
  res.json({
    online: true,
    versao: "Beta 1 - Formatador + IN100",
    apiConfigurada: Boolean(API_KEY)
  });
});

app.listen(PORT, () => {
  console.log(`Beta 1 - Formatador + IN100 rodando na porta ${PORT}`);

  if (!API_KEY) {
    console.log("⚠️ API_KEY não configurada.");
  } else {
    console.log("✅ API_KEY configurada.");
  }
});
