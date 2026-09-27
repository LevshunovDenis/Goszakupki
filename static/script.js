const urlInput = document.getElementById('urlInput');
const parseBtn = document.getElementById('parseBtn');
const generateBtn = document.getElementById('generateBtn');
const manualAmountInput = document.getElementById('manualAmountInput');
const statusEl = document.getElementById('status');
const resultEl = document.getElementById('result');

const customerEl = document.getElementById('customer');
const tenderNumberEl = document.getElementById('tenderNumber');
const purchaseTypeEl = document.getElementById('purchaseType');
const subjectEl = document.getElementById('subject');
const quantityEl = document.getElementById('quantity');
const amountEl = document.getElementById('amount');
const deliveryEl = document.getElementById('delivery');
let currentData = null;

function showStatus(message, type) {
  statusEl.textContent = message;
  statusEl.className = 'status ' + type;
  statusEl.classList.remove('hidden');
}

function hideStatus() {
  statusEl.classList.add('hidden');
}

function setResult(data) {
  currentData = data;
  purchaseTypeEl.textContent = data.purchase_type || '—';
  tenderNumberEl.textContent = data.tender_number || '—';
  customerEl.textContent = data.customer || '—';
  subjectEl.textContent = data.subject || '—';
  quantityEl.textContent = data.quantity || '—';
  amountEl.textContent = data.amount || '—';
  deliveryEl.textContent = data.delivery || '—';
  generateBtn.disabled = false;
  resultEl.classList.remove('hidden');
}

async function parsePurchase() {
  const value = urlInput.value.trim();

  if (!value) {
    showStatus('Введите ссылку или ID закупки', 'error');
    return;
  }

  hideStatus();
  currentData = null;
  generateBtn.disabled = true;
  resultEl.classList.add('hidden');
  parseBtn.disabled = true;
  parseBtn.textContent = 'Загрузка...';

  try {
    const response = await fetch('/api/extract', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ url: value })
    });

    const result = await response.json();

    if (!response.ok || !result.success) {
      showStatus(result.error || 'Ошибка получения данных', 'error');
      resultEl.classList.add('hidden');
      return;
    }

    setResult(result.data);
    showStatus('Данные успешно загружены', 'success');
  } catch (error) {
    showStatus('Не удалось выполнить запрос: ' + error.message, 'error');
    resultEl.classList.add('hidden');
  } finally {
    parseBtn.disabled = false;
    parseBtn.textContent = 'Получить данные';
  }
}

async function generateDocument() {
  const manualAmount = manualAmountInput.value.trim();
  if (!currentData) {
    showStatus('Сначала получите данные закупки', 'error');
    return;
  }
  if (!manualAmount) {
    showStatus('Введите сумму для документа', 'error');
    manualAmountInput.focus();
    return;
  }

  hideStatus();
  generateBtn.disabled = true;
  generateBtn.textContent = 'Формирование...';

  try {
    const response = await fetch('/api/generate', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ data: currentData, manual_amount: manualAmount })
    });

    if (!response.ok) {
      const result = await response.json();
      showStatus(result.error || 'Не удалось сформировать документ', 'error');
      return;
    }

    const fileUrl = URL.createObjectURL(await response.blob());
    const downloadLink = document.createElement('a');
    downloadLink.href = fileUrl;
    const tenderNumber = (currentData.tender_number || '').replace(/[^A-Za-z0-9_-]/g, '');
    downloadLink.download = tenderNumber ? `${tenderNumber}.docx` : 'Закупка.docx';
    downloadLink.click();
    URL.revokeObjectURL(fileUrl);
    showStatus('Word-документ сформирован', 'success');
  } catch (error) {
    showStatus('Не удалось сформировать документ: ' + error.message, 'error');
  } finally {
    generateBtn.disabled = false;
    generateBtn.textContent = 'Сформировать Word';
  }
}

parseBtn.addEventListener('click', parsePurchase);
generateBtn.addEventListener('click', generateDocument);
urlInput.addEventListener('keydown', (event) => {
  if (event.key === 'Enter') {
    parsePurchase();
  }
});

