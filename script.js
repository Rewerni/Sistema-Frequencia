const url = "https://script.google.com/macros/s/AKfycbzHlC67n9PFDs7ThnBrQbURD3-Pg7mvtSSUC13AoQ9xjObXoCQ6YWxWzx8lbpkDLUpi/exec";
const totalMonitores = 14;
let registrosUltimaFrequencia = [];
let dataUltimaFrequencia = null;
let registrosFrequenciaAnterior = [];
let dataFrequenciaAnterior = null;

// ── Utilitários ──────────────────────────────────────────────

function parseIsoDate(value) {
  if (!value) return null;
  const d = new Date(value);
  return isNaN(d.getTime()) ? null : d;
}

function formatDateBR(dateObj) {
  if (!dateObj) return "--";
  return dateObj.toLocaleDateString("pt-BR", { timeZone: "UTC" });
}

function formatWeekdayBR(dateObj) {
  if (!dateObj) return "";
  return dateObj
    .toLocaleDateString("pt-BR", { weekday: "long", timeZone: "UTC" })
    .split("-")
    .map(part => part.charAt(0).toUpperCase() + part.slice(1))
    .join("-");
}

function formatDateTimeBR(dateObj) {
  if (!dateObj) return "--";
  return (
    dateObj.toLocaleDateString("pt-BR", { timeZone: "UTC" }) +
    " às " +
    dateObj.toLocaleTimeString("pt-BR", { hour: "2-digit", minute: "2-digit", timeZone: "UTC" })
  );
}

function formatShortDateTimeBR(dateObj) {
  if (!dateObj) return "--";
  return (
    dateObj.toLocaleDateString("pt-BR", { day: "2-digit", month: "2-digit", timeZone: "UTC" }) +
    " às " +
    dateObj.toLocaleTimeString("pt-BR", { hour: "2-digit", minute: "2-digit", timeZone: "UTC" })
  );
}

function primeiroSegundoNome(nome) {
  const partes = String(nome).trim().split(/\s+/).filter(Boolean);
  return partes.slice(0, 2).join(" ");
}

function esc(text) {
  return String(text)
    .replaceAll("&", "&amp;").replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;").replaceAll('"', "&quot;");
}

function normalizarNomeArquivo(text) {
  return String(text)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-zA-Z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .toLowerCase();
}

function atualizarBotaoDownload(habilitado) {
  const botao = document.getElementById("btn-download-pdf");
  if (!botao) return;
  botao.disabled = !habilitado;
}

function atualizarLinkDownloadAnterior(habilitado) {
  const link = document.getElementById("btn-download-pdf-anterior");
  if (!link) return;
  link.classList.toggle("desabilitado", !habilitado);
  link.setAttribute("aria-disabled", String(!habilitado));

  if (habilitado) {
    link.removeAttribute("tabindex");
  } else {
    link.setAttribute("tabindex", "-1");
  }
}

function animarRosca(valorFinal) {
  const rosca = document.getElementById("rosca");
  const percentualTexto = document.getElementById("percentual");
  const inicio = performance.now();
  const duracao = 1100;
  const limite = Math.max(0, Math.min(100, valorFinal));

  function easeOutCubic(t) {
    return 1 - Math.pow(1 - t, 3);
  }

  function frame(agora) {
    const progresso = Math.min((agora - inicio) / duracao, 1);
    const valor = Math.round(limite * easeOutCubic(progresso));

    percentualTexto.innerHTML = `${valor}<span class="simbolo-percentual">%</span>`;
    rosca.style.setProperty("--pct", valor);

    if (progresso < 1) {
      requestAnimationFrame(frame);
    }
  }

  percentualTexto.innerHTML = `0<span class="simbolo-percentual">%</span>`;
  rosca.style.setProperty("--pct", 0);
  requestAnimationFrame(frame);
}

function baixarPdfFrequencia(registros, dataFrequencia, textoData, nomeArquivoBase) {
  const jsPDF = window.jspdf?.jsPDF;
  if (!jsPDF) {
    alert("Nao foi possivel carregar a biblioteca de PDF. Verifique a conexao e tente novamente.");
    return;
  }

  const doc = new jsPDF({ orientation: "landscape", unit: "pt", format: "a4" });
  if (typeof doc.autoTable !== "function") {
    alert("Nao foi possivel preparar a tabela do PDF. Tente recarregar a pagina.");
    return;
  }

  const dataFormatada = formatDateBR(dataFrequencia);
  const linhas = registros.map(r => [
    r.nome,
    formatDateBR(r.dataObj),
    r.tipo || "--",
    r.email,
  ]);

  doc.setFont("helvetica", "bold");
  doc.setFontSize(18);
  doc.setTextColor(45, 51, 56);
  doc.text("Registro de Frequência", 40, 44);

  doc.setFont("helvetica", "normal");
  doc.setFontSize(10);
  doc.setTextColor(89, 96, 101);
  doc.text(`${textoData}: ${dataFormatada}`, 40, 64);
  doc.text(`Total de assinaturas: ${registros.length}`, 40, 80);

  doc.autoTable({
    startY: 108,
    head: [["Nome", "Frequência", "Tipo", "E-mail institucional"]],
    body: linhas,
    margin: { left: 40, right: 40 },
    tableWidth: "auto",
    styles: {
      font: "helvetica",
      fontSize: 9,
      cellPadding: 8,
      overflow: "linebreak",
      valign: "middle",
      textColor: [89, 96, 101],
      lineColor: [235, 238, 242],
      lineWidth: 0.4,
    },
    headStyles: {
      fillColor: [249, 249, 251],
      textColor: [45, 51, 56],
      fontStyle: "bold",
      lineWidth: 0,
    },
    alternateRowStyles: {
      fillColor: [242, 244, 246],
    },
    columnStyles: {
      0: { cellWidth: 170, textColor: [45, 51, 56] },
      1: { cellWidth: 86 },
      2: { cellWidth: 170 },
      3: { cellWidth: "auto" },
    },
  });

  doc.save(`${nomeArquivoBase}-${normalizarNomeArquivo(dataFormatada)}.pdf`);
}

function baixarPdfUltimaFrequencia() {
  if (!registrosUltimaFrequencia.length || !dataUltimaFrequencia) {
    alert("Aguarde os dados da frequencia carregarem.");
    return;
  }

  baixarPdfFrequencia(
    registrosUltimaFrequencia,
    dataUltimaFrequencia,
    "Última frequência",
    "frequencia"
  );
}

function baixarPdfFrequenciaAnterior(event) {
  event.preventDefault();

  if (!registrosFrequenciaAnterior.length || !dataFrequenciaAnterior) {
    alert("Ainda nao ha uma frequencia anterior para exportar.");
    return;
  }

  baixarPdfFrequencia(
    registrosFrequenciaAnterior,
    dataFrequenciaAnterior,
    "Frequência anterior",
    "frequencia-anterior"
  );
}

// ── Render ───────────────────────────────────────────────────

function renderDados(data) {
  if (!Array.isArray(data) || data.length === 0) throw new Error("Sem dados.");

  // Monta registros — chaves já vêm sem espaço graças ao .trim() no Apps Script
  const registros = data.map(item => ({
  nome:           String(item["Nome completo"] ?? "").trim(),
  email:          String(item["Email Institucional"] ?? "").trim(),
  tipo:           String(item["Tipo"] ?? "").trim(),
  dataFrequencia: item["Data da Frequência"] || item["Data da Frequencia"] || "",
  carimboObj:     parseIsoDate(item["Carimbo de data/hora"]),
  dataObj:        parseIsoDate(item["Data da Frequência"] || item["Data da Frequencia"]),
})).filter(r => r.nome && r.email && r.dataObj);
    
  if (registros.length === 0) throw new Error("Nenhum registro válido.");

  // Datas únicas ordenadas (mais recente primeiro)
  const datasUnicas = [...new Set(registros.map(r => r.dataObj.getTime()))]
    .sort((a, b) => b - a);

  const [ultimaDataTs, segundaDataTs] = datasUnicas;
  const duasUltimas = new Set([ultimaDataTs, segundaDataTs].filter(Boolean));

  const registrosUltimaData = registros.filter(r => r.dataObj.getTime() === ultimaDataTs);
  registrosUltimaFrequencia = registrosUltimaData
    .slice()
    .sort((a, b) => (b.carimboObj?.getTime() ?? 0) - (a.carimboObj?.getTime() ?? 0));
  dataUltimaFrequencia = registrosUltimaData[0]?.dataObj ?? null;
  atualizarBotaoDownload(registrosUltimaFrequencia.length > 0);

  const registrosDataAnterior = segundaDataTs
    ? registros.filter(r => r.dataObj.getTime() === segundaDataTs)
    : [];
  registrosFrequenciaAnterior = registrosDataAnterior
    .slice()
    .sort((a, b) => (b.carimboObj?.getTime() ?? 0) - (a.carimboObj?.getTime() ?? 0));
  dataFrequenciaAnterior = registrosDataAnterior[0]?.dataObj ?? null;
  atualizarLinkDownloadAnterior(registrosFrequenciaAnterior.length > 0);

  const registrosTabela = registros
    .filter(r => duasUltimas.has(r.dataObj.getTime()))
    .sort((a, b) => {
      const dDiff = b.dataObj - a.dataObj;
      if (dDiff !== 0) return dDiff;
      return (b.carimboObj?.getTime() ?? 0) - (a.carimboObj?.getTime() ?? 0);
    });

  const totalUltima   = registrosUltimaData.length;
  const percentual    = Math.min(100, Math.round((totalUltima / totalMonitores) * 100));

  // ✅ CORRIGIDO: usava .carimbo (string) em vez de .carimboObj (Date)
  const ultimoCarimbo = registrosUltimaData
    .filter(r => r.carimboObj)
    .sort((a, b) => b.carimboObj - a.carimboObj)[0];

  // ── Preenche DOM ─────────────────────────────────────────────

  document.getElementById("data").innerHTML =
    `<span class="data-card-data">${formatDateBR(registrosUltimaData[0].dataObj)}</span>` +
    `<span class="data-card-dia">${formatWeekdayBR(registrosUltimaData[0].dataObj)}</span>`;

  document.getElementById("contador-assinaturas").textContent =
    `${totalUltima} ${totalUltima === 1 ? "assinatura" : "assinaturas"} até o momento`;

  document.getElementById("ultimo-tipo").textContent =
    ultimoCarimbo?.tipo || registrosUltimaData[0]?.tipo || "--";

  document.getElementById("carimbo-recente").textContent =
    formatShortDateTimeBR(ultimoCarimbo?.carimboObj ?? null);

  animarRosca(percentual);

  document.getElementById("corpo-tabela").innerHTML = registrosTabela.map(r => `
     <tr>
       <td>${esc(r.nome)}</td>
       <td>${esc(formatDateBR(r.dataObj))}</td>
       <td>${esc(r.tipo || "--")}</td>
       <td>${esc(r.email)}</td>
      </tr>
    `).join("");
}

function mostrarErro(err) {
  console.error("Erro:", err);
  registrosUltimaFrequencia = [];
  dataUltimaFrequencia = null;
  registrosFrequenciaAnterior = [];
  dataFrequenciaAnterior = null;
  atualizarBotaoDownload(false);
  atualizarLinkDownloadAnterior(false);
  document.getElementById("data").textContent = "--";
  document.getElementById("contador-assinaturas").textContent = "0 assinaturas até o momento";
  document.getElementById("ultimo-tipo").textContent = "--";
  document.getElementById("carimbo-recente").textContent = "--";
  document.getElementById("percentual").innerHTML = `0<span class="simbolo-percentual">%</span>`;
  document.getElementById("rosca").style.setProperty("--pct", 0);
  document.getElementById("corpo-tabela").innerHTML =
    `<tr><td colspan="4">Nenhum registro encontrado.</td></tr>`;
}

// ── Fetch ────────────────────────────────────────────────────

fetch(url)
  .then(res => {
    if (!res.ok) throw new Error("HTTP " + res.status);
    return res.json();
  })
  .then(renderDados)
  .catch(mostrarErro);

atualizarBotaoDownload(false);
atualizarLinkDownloadAnterior(false);
document.getElementById("btn-download-pdf")?.addEventListener("click", baixarPdfUltimaFrequencia);
document.getElementById("btn-download-pdf-anterior")?.addEventListener("click", baixarPdfFrequenciaAnterior);
