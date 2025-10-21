/*
 * script.js
 *
 * Este arquivo implementa a lógica principal do editor de documentos.
 * Lida com a leitura de arquivos DOCX e PDF, a conversão para HTML,
 * a substituição inteligente dos campos de sublinhados por elementos
 * interativos, e a exportação do resultado para PDF ou Word.
 */

// Configurar pdf.js. Definimos a URL do worker dinamicamente a partir do CDN.
pdfjsLib.GlobalWorkerOptions.workerSrc =
  'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.14.305/pdf.worker.min.js';

// Aguarda que o DOM esteja pronto
window.addEventListener('DOMContentLoaded', () => {
  const container = document.getElementById('documentContainer');
  const fileInput = document.getElementById('fileInput');
  const downloadPdfBtn = document.getElementById('downloadPdf');
  const downloadDocBtn = document.getElementById('downloadDoc');

  // Carregar modelo padrão (gerado pelo pandoc a partir do DOCX fornecido)
  fetch('template.html')
    .then((res) => res.text())
    .then((html) => {
      const processed = processHtmlToInputs(html);
      container.innerHTML = processed;
      // Ativa botões de download para o template inicial
      downloadPdfBtn.disabled = false;
      downloadDocBtn.disabled = false;
    })
    .catch((err) => {
      console.error('Erro ao carregar o modelo padrão:', err);
    });

  // Lidar com uploads de arquivo
  fileInput.addEventListener('change', (event) => {
    const file = event.target.files[0];
    if (!file) return;
    const ext = file.name.split('.').pop().toLowerCase();
    // Limpar conteúdo existente
    container.innerHTML = '';
    downloadPdfBtn.disabled = true;
    downloadDocBtn.disabled = true;
    if (ext === 'docx') {
      handleDocxUpload(file, container, downloadPdfBtn, downloadDocBtn);
    } else if (ext === 'pdf') {
      handlePdfUpload(file, container, downloadPdfBtn, downloadDocBtn);
    } else {
      alert('Formato não suportado. Por favor, selecione um arquivo DOCX ou PDF.');
    }
  });

  // Download em PDF
  downloadPdfBtn.addEventListener('click', async () => {
    const filledHtml = getFilledHtml(container);
    if (!filledHtml) return;
    // Configurações para o html2pdf
    const opt = {
      margin: [10, 10, 10, 10],
      filename: 'documento-final.pdf',
      html2canvas: { scale: 2, useCORS: true },
      jsPDF: { unit: 'pt', format: 'a4', orientation: 'portrait' },
    };
    // Cria um elemento temporário com o HTML preenchido
    const tmp = document.createElement('div');
    tmp.style.padding = '20px';
    tmp.innerHTML = filledHtml;
    // Usa a biblioteca html2pdf para gerar o PDF
    await html2pdf().set(opt).from(tmp).save();
  });

  // Download em Word
  downloadDocBtn.addEventListener('click', async () => {
    const filledHtml = getFilledHtml(container);
    if (!filledHtml) return;
    // Envolve o HTML em uma estrutura completa de documento para o Word
    const fullHtml = buildFullHtml(filledHtml);
    const blob = window.htmlDocx.asBlob(fullHtml);
    saveAs(blob, 'documento-final.docx');
  });
});

/**
 * Converte um HTML em formato string para um HTML com campos interativos no lugar
 * dos sublinhados (__, ___, etc.). A função percorre os nós de texto e
 * substitui sequências de dois ou mais caracteres '_' por elementos <input>.
 *
 * @param {string} htmlString
 * @returns {string} HTML processado pronto para exibição
 */
function processHtmlToInputs(htmlString) {
  // Usamos DOMParser para transformar a string em um documento DOM
  const parser = new DOMParser();
  const doc = parser.parseFromString(htmlString, 'text/html');
  const body = doc.body;

  /**
   * Função recursiva que percorre os nós e faz a substituição nos nós de texto.
   * @param {Node} node
   */
  function traverse(node) {
    node.childNodes.forEach((child) => {
      if (child.nodeType === Node.TEXT_NODE) {
        const text = child.textContent;
        if (/__{1,}/.test(text)) {
          // Encontrou uma ou mais sequências de sublinhados
          const fragment = document.createDocumentFragment();
          let lastIndex = 0;
          const regex = /(_{2,})/g;
          let match;
          while ((match = regex.exec(text)) !== null) {
            // Texto antes dos sublinhados
            const before = text.substring(lastIndex, match.index);
            if (before) {
              fragment.appendChild(document.createTextNode(before));
            }
            const underscores = match[0];
            const input = document.createElement('input');
            input.type = 'text';
            input.className = 'editable-field';
            // A largura inicial baseia-se na quantidade de underscores
            const length = underscores.length;
            input.setAttribute('data-placeholder-length', length);
            input.style.width = `${length * 8}px`;
            // Ajustar largura conforme digitação
            input.addEventListener('input', function () {
              const len = this.value.length || this.getAttribute('data-placeholder-length');
              // Ajusta largura com limite mínimo
              const newWidth = Math.max(len, this.getAttribute('data-placeholder-length')) * 8 + 20;
              this.style.width = `${newWidth}px`;
            });
            fragment.appendChild(input);
            lastIndex = match.index + underscores.length;
          }
          // Texto após a última sequência
          const after = text.substring(lastIndex);
          if (after) {
            fragment.appendChild(document.createTextNode(after));
          }
          // Substitui o nó de texto original pelo fragmento
          child.parentNode.replaceChild(fragment, child);
        }
      } else if (child.nodeType === Node.ELEMENT_NODE) {
        traverse(child);
      }
    });
  }
  traverse(body);
  // Retorna o conteúdo interno do body
  return body.innerHTML;
}

/**
 * Lida com upload de arquivo DOCX. Utiliza a biblioteca Mammoth para converter
 * o conteúdo do DOCX em HTML e, em seguida, processa o HTML para inserir
 * campos interativos.
 *
 * @param {File} file O arquivo DOCX selecionado pelo usuário
 * @param {HTMLElement} container O contêiner onde o conteúdo será inserido
 * @param {HTMLElement} pdfBtn Botão de download do PDF
 * @param {HTMLElement} docBtn Botão de download do Word
 */
function handleDocxUpload(file, container, pdfBtn, docBtn) {
  const reader = new FileReader();
  reader.onload = async function (e) {
    const arrayBuffer = e.target.result;
    try {
      const result = await mammoth.convertToHtml({ arrayBuffer });
      const html = result.value; // HTML simples gerado pelo Mammoth
      const processed = processHtmlToInputs(html);
      container.innerHTML = processed;
      pdfBtn.disabled = false;
      docBtn.disabled = false;
    } catch (err) {
      console.error('Erro ao converter DOCX:', err);
      alert('Erro ao ler o documento DOCX.');
    }
  };
  reader.readAsArrayBuffer(file);
}

/**
 * Lida com upload de arquivo PDF. Usa pdf.js para extrair o texto e
 * reconstruir parágrafos mantendo quebras de linha. O texto extraído é
 * então processado para substituir os sublinhados por campos interativos.
 *
 * @param {File} file O arquivo PDF selecionado pelo usuário
 * @param {HTMLElement} container O contêiner onde o conteúdo será inserido
 * @param {HTMLElement} pdfBtn Botão de download do PDF
 * @param {HTMLElement} docBtn Botão de download do Word
 */
function handlePdfUpload(file, container, pdfBtn, docBtn) {
  const reader = new FileReader();
  reader.onload = async function (e) {
    const arrayBuffer = e.target.result;
    try {
      const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
      let fullHtml = '';
      for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
        const page = await pdf.getPage(pageNum);
        const content = await page.getTextContent();
        const strings = content.items.map((item) => item.str);
        // Agrupa strings em um único parágrafo; quebra de página vira nova linha
        const pageText = strings.join(' ');
        fullHtml += `<p>${escapeHtml(pageText)}</p>`;
      }
      const processed = processHtmlToInputs(fullHtml);
      container.innerHTML = processed;
      pdfBtn.disabled = false;
      docBtn.disabled = false;
    } catch (err) {
      console.error('Erro ao processar PDF:', err);
      alert('Erro ao ler o documento PDF.');
    }
  };
  reader.readAsArrayBuffer(file);
}

/**
 * Constrói uma string HTML completa (com DOCTYPE, html, head e body) a partir
 * do conteúdo interno fornecido. Inclui estilos simples para preservar
 * formatação básica e sublinhados nos campos preenchidos.
 *
 * @param {string} bodyContent O HTML que será colocado dentro do <body>
 * @returns {string} HTML completo
 */
function buildFullHtml(bodyContent) {
  return `<!DOCTYPE html>
  <html lang="pt-BR">
    <head>
      <meta charset="UTF-8" />
      <style>
        body { font-family: Georgia, serif; line-height: 1.5; }
        .filled-field { border-bottom: 1px solid #000; padding: 0 2px; }
      </style>
    </head>
    <body>
      ${bodyContent}
    </body>
  </html>`;
}

/**
 * Itera pelo contêiner do documento e constrói uma string HTML onde cada
 * elemento <input> é substituído pelo texto digitado pelo usuário. A classe
 * `.filled-field` é aplicada para manter a linha de sublinhado no
 * documento final.
 *
 * @param {HTMLElement} container O contêiner que contém o documento editável
 * @returns {string} HTML preenchido
 */
function getFilledHtml(container) {
  // Clona o contêiner para não modificar o DOM original
  const clone = container.cloneNode(true);
  const inputs = clone.querySelectorAll('input.editable-field');
  inputs.forEach((input) => {
    const span = document.createElement('span');
    span.className = 'filled-field';
    const value = input.value || '';
    const placeholderLength = parseInt(input.getAttribute('data-placeholder-length'), 10) || 3;
    // Se valor vazio, usa espaços não separáveis para manter a largura
    const displayText = value ? escapeHtml(value) : '&nbsp;'.repeat(placeholderLength);
    span.innerHTML = displayText;
    input.parentNode.replaceChild(span, input);
  });
  return clone.innerHTML;
}

/**
 * Escapa caracteres HTML especiais para evitar quebras inesperadas ao
 * inserir conteúdo em HTML dinâmico.
 *
 * @param {string} str
 * @returns {string}
 */
function escapeHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}
