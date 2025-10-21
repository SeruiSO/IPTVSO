// Ініціалізація змінних та констант
let touchStartX = 0;
let touchEndX = 0;
let touchStartTarget = null;
let dataToSubmit = {};
let isSubmitting = false;
let currencyMessages = [];
let expensesMessages = [];
let historyMessages = [];
let currentThemeIndex = 0;
const themes = ['dark', 'grey', 'white'];

// Список валют
const currencies = [
  'USD', 'USDNOV', 'EUR', 'EURMELKI', 'UAH', 'PLN', 'GBP', 
  'RON', 'MDL', 'CHF', 'CZK', 'CAD'
];

// Валюти для відображення в бігущій строці (НБУ)
const displayCurrencies = ['USD', 'EUR', 'PLN', 'GBP', 'RON', 'CHF'];

// Валюти для залишків
const balanceCurrencies = [
  'USD', 'USDNOV', 'EUR', 'EURMELKI', 'UAH', 'PLN', 
  'GBP', 'RON', 'MDL', 'CHF', 'CZK', 'UAH'
];

// DOM-елементи
const domElements = {
  currencyTableBody: document.getElementById('currencyTableBody'),
  historyContent: document.getElementById('historyContent'),
  historyMessageList: document.getElementById('historyMessageList'),
  expensesTableBody: document.getElementById('expensesTableBody'),
  expensesMessageList: document.getElementById('expensesMessageList'),
  balancesTableBody: document.getElementById('balancesTableBody'),
  operation: document.getElementById('operation'),
  dateInput: document.getElementById('dateInput'),
  saveButton: document.getElementById('saveButton'),
  totalGrn: document.getElementById('total_grn'),
  confirmModal: document.getElementById('confirmModal'),
  modalTableBody: document.getElementById('modalTableBody'),
  modalTotalGrn: document.getElementById('modalTotalGrn'),
  modalPreviewHeader: document.getElementById('modalPreviewHeader'),
  confirmButton: document.getElementById('confirmButton'),
  undoModal: document.getElementById('undoModal'),
  undoTableBody: document.getElementById('undoTableBody'),
  undoPreviewHeader: document.getElementById('undoPreviewHeader'),
  undoConfirmButton: document.getElementById('undoConfirmButton'),
  undoCancelButton: document.getElementById('undoCancelButton'),
  saveExpensesButton: document.getElementById('saveExpensesButton'),
  loader: document.getElementById('loader'),
  historyLoader: document.getElementById('historyLoader'),
  expensesLoader: document.getElementById('expensesLoader'),
  balancesLoader: document.getElementById('balancesLoader'),
  messageList: document.getElementById('messageList'),
  fetchButton: document.getElementById('fetchButton'),
  pointModal: document.getElementById('pointModal'),
  pointSelect: document.getElementById('pointSelect'),
  tickerContent: document.getElementById('tickerContent')
};

// Список вкладок
const tabs = ['tab-currency', 'tab-history', 'tab-expenses', 'tab-balances'];

// Категорії витрат
const expenseCategories = [
  "Поліна заняття, додаткові, розваги",
  "Еда: ресторан, кафе",
  "Подарки",
  "Магазин/базар",
  "Транспортні (бензин, парковка) расх. на авто",
  "Інтернет, телефон",
  "Кофе",
  "Поповнення карточки",
  "Тварини",
  "Лікування",
  "За старі долари",
  "РОЗМІН",
  "ТЕСЛА",
  "Абонплата",
  "Розмін",
  "Категорія 16",
  "Категорія 17",
  "Категорія 18",
  "Категорія 19",
  "Категорія 20",
  "Недостача/лишек",
  "Другое"
];

// Генерація UUID
function generateUUID() {
  return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, c => {
    const r = Math.random() * 16 | 0;
    const v = c === 'x' ? r : (r & 0x3 | 0x8);
    return v.toString(16);
  });
}

// Дебаунсинг для обробки подій
function debounce(func, wait) {
  let timeout;
  return function executedFunction(...args) {
    const later = () => {
      clearTimeout(timeout);
      func(...args);
    };
    clearTimeout(timeout);
    timeout = setTimeout(later, wait);
  };
}

// Вібрація та звуковий сигнал для помилок
function triggerErrorFeedback(message) {
  if (message.includes('❌') || message.includes('суми згруповані')) {
    if ('vibrate' in navigator) {
      navigator.vibrate([900, 300, 900]);
    }
    try {
      const audioContext = new (window.AudioContext || window.webkitAudioContext)();
      const oscillator = audioContext.createOscillator();
      oscillator.type = 'square';
      oscillator.frequency.setValueAtTime(440, audioContext.currentTime);
      oscillator.connect(audioContext.destination);
      oscillator.start();
      oscillator.stop(audioContext.currentTime + 0.5);
    } catch (e) {
      console.error('Помилка відтворення звуку:', e);
    }
  }
}

// Зміна теми
function cycleTheme() {
  currentThemeIndex = (currentThemeIndex + 1) % themes.length;
  const theme = themes[currentThemeIndex];
  document.body.className = theme === 'dark' ? '' : `${theme}-theme`;
  localStorage.setItem('selectedTheme', theme);
}

// Ініціалізація при завантаженні сторінки
document.addEventListener('DOMContentLoaded', () => {
  // Відновлення збереженої теми
  const savedTheme = localStorage.getItem('selectedTheme') || 'dark';
  currentThemeIndex = themes.indexOf(savedTheme);
  document.body.className = savedTheme === 'dark' ? '' : `${savedTheme}-theme`;

  // Встановлення поточної дати
  const today = new Date();
  const day = String(today.getDate()).padStart(2, '0');
  domElements.dateInput.value = day;

  // Ініціалізація форм
  initForm();
  initExpensesForm();

  // Відновлення активної вкладки
  const savedTab = sessionStorage.getItem('activeTab') || 'tab-currency';
  switchTab(savedTab);

  // Завантаження курсів НБУ для бігущої строки
  loadNbuRates();

  // Обробники подій для вкладок
  document.querySelectorAll('.tab').forEach(tab => {
    tab.addEventListener('click', function() {
      const tabId = this.getAttribute('data-tab');
      switchTab(tabId);
    });
  });

  // Обробка свайпів
  document.addEventListener('touchstart', function(e) {
    touchStartX = e.changedTouches[0].screenX;
    touchStartTarget = e.target;
  });

  document.addEventListener('touchend', function(e) {
    touchEndX = e.changedTouches[0].screenX;
    if (touchStartTarget && touchStartTarget.closest('.history-table')) {
      return;
    }
    handleSwipe();
  });

  // Обробка зміни дати
  domElements.dateInput.addEventListener('change', function() {
    loadExpenses();
    loadBalances();
  });

  // Обробники для форми валют
  domElements.saveButton.addEventListener('click', prepareSubmission);
  document.getElementById('cancelButton').addEventListener('click', closeModal);
  domElements.confirmButton.addEventListener('click', function() {
    domElements.loader.style.display = "block";
    submitConfirmed();
  });
  domElements.undoCancelButton.addEventListener('click', closeUndoModal);
  domElements.undoConfirmButton.addEventListener('click', confirmUndo);
  domElements.operation.addEventListener('change', function() {
    this.style.backgroundColor = this.value === 'buy' ? '#4CAF50' : '#f44336';
  });
  domElements.operation.style.backgroundColor = 
    domElements.operation.value === 'buy' ? '#4CAF50' : '#f44336';

  // Обробники для форми витрат
  domElements.saveExpensesButton.addEventListener('click', saveExpenses);

  // Нові обробники для підтягування даних
  domElements.fetchButton.addEventListener('click', showPointModal);
});

// Обробка свайпів між вкладками
function handleSwipe() {
  const swipeThreshold = 70;
  const swipeDistance = touchStartX - touchEndX;

  if (Math.abs(swipeDistance) < swipeThreshold) return;

  const currentTab = tabs.find(tab => document.getElementById(tab).classList.contains('active'));
  const currentIndex = tabs.indexOf(currentTab);

  if (swipeDistance > 0 && currentIndex < tabs.length - 1) {
    switchTab(tabs[currentIndex + 1]);
  } else if (swipeDistance < 0 && currentIndex > 0) {
    switchTab(tabs[currentIndex - 1]);
  }
}

// Перемикання вкладок
function switchTab(tabId) {
  document.querySelectorAll('.tab-content').forEach(tab => {
    tab.classList.remove('active');
  });
  document.querySelectorAll('.tab').forEach(tab => {
    tab.classList.remove('active');
  });
  document.getElementById(tabId).classList.add('active');
  document.querySelector(`.tab[data-tab="${tabId}"]`).classList.add('active');
  switch (tabId) {
    case 'tab-history':
      loadHistory();
      break;
    case 'tab-expenses':
      loadExpenses();
      break;
    case 'tab-balances':
      loadBalances();
      break;
  }
  sessionStorage.setItem('activeTab', tabId);
}

// Завантаження курсів НБУ для бігущої строки
function loadNbuRates() {
  const cachedRates = sessionStorage.getItem('nbuRates');
  if (cachedRates) {
    displayNbuTicker(JSON.parse(cachedRates));
    return;
  }
  fetch(WEB_APP_URL, {
    method: 'POST',
    body: JSON.stringify({ method: 'getNbuRates' }),
    headers: { 'Content-Type': 'application/json' }
  })
    .then(response => response.json())
    .then(rates => {
      sessionStorage.setItem('nbuRates', JSON.stringify(rates));
      displayNbuTicker(rates);
    })
    .catch(error => {
      domElements.tickerContent.innerHTML = `Помилка завантаження курсів: ${error.message}`;
    });
}

// Відображення бігущої строки з курсами НБУ
function displayNbuTicker(rates) {
  let html = '';
  displayCurrencies.forEach(currency => {
    html += `<span class="ticker-item"><span class="currency-span" data-currency="${currency}">${currency}</span>: ${formatNumber(rates[currency], 4)} UAH</span>`;
  });
  // Подвоюємо контент для безшовного циклу
  domElements.tickerContent.innerHTML = html + html;
}

// Ініціалізація форми валют
function initForm() {
  let html = '';
  currencies.forEach(currency => {
    html += `
      <tr>
        <td class="currency-cell" data-currency="${currency}">${currency}</td>
        <td><input type="text" id="amount_${currency}" class="currency-input amount-input" inputmode="decimal"></td>
        <td><input type="text" id="rate_${currency}" class="currency-input rate-input" inputmode="decimal"></td>
        <td><input type="text" id="grn_${currency}" readonly></td>
      </tr>
    `;
  });
  domElements.currencyTableBody.innerHTML = html;

  const debouncedCalculate = debounce((currency) => calculateTotal(currency), 100);
  document.querySelectorAll('.amount-input').forEach(input => {
    input.addEventListener('input', function() {
      this.value = formatInput(this.value, 2);
      debouncedCalculate(this.id.split('_')[1]);
    });
  });
  document.querySelectorAll('.rate-input').forEach(input => {
    input.addEventListener('input', function() {
      this.value = formatInput(this.value, 4);
      debouncedCalculate(this.id.split('_')[1]);
    });
  });
}

// Ініціалізація форми витрат
function initExpensesForm() {
  let html = '';
  expenseCategories.forEach((category, index) => {
    html += `
      <tr>
        <td>${category}</td>
        <td><input type="text" id="expense_${index}" class="expense-input" inputmode="decimal"></td>
      </tr>
    `;
  });
  domElements.expensesTableBody.innerHTML = html;

  document.querySelectorAll('.expense-input').forEach(input => {
    input.addEventListener('input', function() {
      this.value = formatInput(this.value, 2);
    });
  });
}

// Завантаження історії
async function loadHistory() {
  domElements.historyLoader.style.display = 'block';
  const spreadsheetId = document.getElementById('spreadsheetSelect').value;
  try {
    const response = await fetch(WEB_APP_URL, {
      method: 'POST',
      body: JSON.stringify({
        method: 'getLastHistoryEntries',
        spreadsheetId,
        params: { limit: 10 }
      }),
      headers: { 'Content-Type': 'application/json' }
    });
    const history = await response.json();
    displayHistory(history);
    domElements.historyLoader.style.display = 'none';
  } catch (error) {
    domElements.historyContent.innerHTML = 
      `<p style="color: red">Помилка завантаження історії: ${error.message}</p>`;
    domElements.historyLoader.style.display = 'none';
  }
}

// Завантаження витрат
async function loadExpenses() {
  domElements.expensesLoader.style.display = 'block';
  const sheetName = domElements.dateInput.value;
  const spreadsheetId = document.getElementById('spreadsheetSelect').value;
  try {
    const response = await fetch(WEB_APP_URL, {
      method: 'POST',
      body: JSON.stringify({
        method: 'getExpenses',
        spreadsheetId,
        params: { sheetName }
      }),
      headers: { 'Content-Type': 'application/json' }
    });
    const expenses = await response.json();
    displayExpenses(expenses);
    domElements.expensesLoader.style.display = 'none';
  } catch (error) {
    showExpensesMessage(`❌ Помилка завантаження витрат: ${error.message}`, 'error');
    domElements.expensesLoader.style.display = 'none';
  }
}

// Завантаження залишків
async function loadBalances() {
  domElements.balancesLoader.style.display = 'block';
  const sheetName = domElements.dateInput.value;
  const spreadsheetId = document.getElementById('spreadsheetSelect').value;
  try {
    const response = await fetch(WEB_APP_URL, {
      method: 'POST',
      body: JSON.stringify({
        method: 'getBalances',
        spreadsheetId,
        params: { sheetName }
      }),
      headers: { 'Content-Type': 'application/json' }
    });
    const balances = await response.json();
    displayBalances(balances);
    domElements.balancesLoader.style.display = 'none';
  } catch (error) {
    showCurrencyMessage(`❌ Помилка завантаження залишків: ${error.message}`, 'error');
    domElements.balancesLoader.style.display = 'none';
  }
}

// Відображення історії
function displayHistory(historyData) {
  if (!historyData || historyData.length === 0) {
    domElements.historyContent.innerHTML = "<p>Немає операцій в історії</p>";
    return;
  }
  
  let html = `
    <table class="history-table">
      <thead>
        <tr>
          <th>Час</th>
          <th>Тип</th>
          <th>Лист</th>
          <th>Валюта</th>
          <th>Сума</th>
          <th>Курс</th>
          <th>ГРН</th>
          <th>Дія</th>
        </tr>
      </thead>
      <tbody>
  `;
  historyData.forEach(item => {
    html += `
      <tr class="${item.type === 'Купівля' ? 'buy-row' : 'sell-row'}">
        <td>${item.time}</td>
        <td>${item.type}</td>
        <td>${item.sheet}</td>
        <td class="history-currency-cell" data-currency="${item.currency}">${item.currency}</td>
        <td>${item.amount}</td>
        <td>${item.rate}</td>
        <td>${item.grn}</td>
        <td><button class="delete-btn" data-transaction-id="${item.transactionId}" data-currency="${item.currency}" data-amount="${item.amount}" data-rate="${item.rate}" data-sheet="${item.sheet}">Видалити</button></td>
      </tr>
    `;
  });
  html += `</tbody></table>`;
  domElements.historyContent.innerHTML = html;

  document.querySelectorAll('.delete-btn').forEach(button => {
    button.addEventListener('click', function() {
      const data = {
        transactionId: this.getAttribute('data-transaction-id'),
        currency: this.getAttribute('data-currency'),
        amount: this.getAttribute('data-amount'),
        rate: this.getAttribute('data-rate'),
        sheet: this.getAttribute('data-sheet'),
        time: this.parentElement.parentElement.children[0].textContent,
        type: this.parentElement.parentElement.children[1].textContent,
        grn: this.parentElement.parentElement.children[6].textContent
      };
      prepareUndoModal(data);
    });
  });
}

// Відображення витрат
function displayExpenses(expenses) {
  if (!expenses || expenses.length === 0) {
    expenses = Array(expenseCategories.length).fill("");
  }
  
  expenseCategories.forEach((category, index) => {
    const input = document.getElementById(`expense_${index}`);
    if (input) {
      input.value = expenses[index] ? expenses[index].replace('.', ',') : '';
    }
  });
}

// Відображення залишків
function displayBalances(balances) {
  let html = '';
  balanceCurrencies.forEach((currency, index) => {
    let value = '';
    if (balances[index]) {
      const numValue = parseFloat(balances[index].replace(',', '.'));
      value = isNaN(numValue) ? '' : numValue.toFixed(0);
    }
    const currencyClass = currency === "UAH" && index === 11 ? "UAH2" : currency;
    html += `
      <tr>
        <td class="currency-cell" data-currency="${currencyClass}">${currency}${index === 11 ? " (2)" : ""}</td>
        <td><input type="text" id="balance_${index}" value="${value}" readonly></td>
      </tr>
    `;
  });
  domElements.balancesTableBody.innerHTML = html;
}

// Форматування чисел
function formatNumber(value, decimals) {
  if (!value || isNaN(value)) return '—';
  return parseFloat(value).toFixed(decimals).replace('.', ',');
}

// Форматування введення
function formatInput(value, decimals) {
  if (!value) return '';
  value = value.replace(',', '.').replace(/[^\d.]/g, '');
  const parts = value.split('.');
  if (parts.length > 1) {
    parts[1] = parts[1].slice(0, decimals);
    value = parts.join('.');
  }
  return value.replace('.', ',');
}

// Обчислення суми в гривнях
function calculateTotal(currency) {
  const amountInput = document.getElementById(`amount_${currency}`);
  const rateInput = document.getElementById(`rate_${currency}`);
  const amount = parseFloat(amountInput.value.replace(',', '.')) || 0;
  const rate = parseFloat(rateInput.value.replace(',', '.')) || 0;
  const grn = amount * rate;
  document.getElementById(`grn_${currency}`).value = grn.toFixed(2).replace('.', ',');
  updateTotalGrn();
}

// Оновлення загальної суми
function updateTotalGrn() {
  let total = 0;
  currencies.forEach(currency => {
    const grnValue = document.getElementById(`grn_${currency}`).value.replace(',', '.');
    total += parseFloat(grnValue) || 0;
  });
  const totalStr = total.toFixed(2).replace('.', ',');
  domElements.totalGrn.value = totalStr;
  domElements.modalTotalGrn.value = totalStr;
}

// Підготовка до відправки даних
function prepareSubmission() {
  const pendingTransaction = sessionStorage.getItem('pendingTransaction');
  if (pendingTransaction) {
    showCurrencyMessage('⏳ Попередній запит ще обробляється. Зачекайте!', 'warning');
    return;
  }

  dataToSubmit = { transactionId: generateUUID() };
  sessionStorage.setItem('pendingTransaction', dataToSubmit.transactionId);
  let hasData = false;
  domElements.modalTableBody.innerHTML = "";
  let totalGrn = 0;

  let html = '';
  currencies.forEach(currency => {
    const amount = document.getElementById(`amount_${currency}`).value.trim();
    const rate = document.getElementById(`rate_${currency}`).value.trim();
    if (amount && rate) {
      dataToSubmit[currency] = {
        displayAmount: amount,
        displayRate: rate,
        amount: amount.replace(',', '.'),
        rate: rate.replace(',', '.')
      };
      const amountNum = parseFloat(dataToSubmit[currency].amount);
      const rateNum = parseFloat(dataToSubmit[currency].rate);
      const grn = (amountNum * rateNum);
      totalGrn += grn;
      html += `
        <tr>
          <td class="currency-cell" data-currency="${currency}">${currency}</td>
          <td>${dataToSubmit[currency].displayAmount}</td>
          <td>${dataToSubmit[currency].displayRate}</td>
          <td>${grn.toFixed(2).replace('.', ',')}</td>
        </tr>
      `;
      hasData = true;
    }
  });

  if (!hasData) {
    showCurrencyMessage('⚠ Введіть хоча б одне значення!', 'warning');
    sessionStorage.removeItem('pendingTransaction');
    return;
  }

  domElements.modalTableBody.innerHTML = html;
  const type = domElements.operation.value;
  const day = domElements.dateInput.value;
  
  domElements.modalPreviewHeader.innerHTML = `<div class="modal-operation-header ${type}">
    <div><b>Операція:</b> ${type === 'buy' ? 'Купівля' : 'Продаж'}</div>
    <div><b>Дата:</b> ${day}</div>
  </div>`;
  
  domElements.confirmModal.style.display = "block";
}

// Збереження витрат
async function saveExpenses() {
  const expenses = [];
  expenseCategories.forEach((category, index) => {
    const value = document.getElementById(`expense_${index}`).value.trim();
    expenses.push(value ? value.replace(',', '.') : "");
  });

  const sheetName = domElements.dateInput.value;
  const spreadsheetId = document.getElementById('spreadsheetSelect').value;
  domElements.expensesLoader.style.display = 'block';
  domElements.saveExpensesButton.disabled = true;

  try {
    const response = await fetch(WEB_APP_URL, {
      method: 'POST',
      body: JSON.stringify({
        method: 'saveExpenses',
        spreadsheetId,
        params: { expenses, sheetName }
      }),
      headers: { 'Content-Type': 'application/json' }
    });
    const result = await response.json();
    showExpensesMessage(result.message, 'success');
    domElements.expensesLoader.style.display = 'none';
    domElements.saveExpensesButton.disabled = false;
  } catch (error) {
    showExpensesMessage(`❌ Помилка збереження витрат: ${error.message}`, 'error');
    domElements.expensesLoader.style.display = 'none';
    domElements.saveExpensesButton.disabled = false;
  }
}

// Відправка даних з повторними спробами
async function submitWithRetry(data, type, sheetName, retries = 3, delay = 2000) {
  if (retries <= 0) {
    domElements.loader.style.display = "none";
    showCurrencyMessage("❌ Не вдалося виконати запит. Перевірте з'єднання.", 'error');
    sessionStorage.removeItem('pendingTransaction');
    domElements.saveButton.disabled = false;
    domElements.confirmButton.disabled = false;
    isSubmitting = false;
    return;
  }

  showCurrencyMessage(`⏳ Спроба ${4 - retries}/3...`, 'warning');

  const spreadsheetId = document.getElementById('spreadsheetSelect').value;
  try {
    const response = await fetch(WEB_APP_URL, {
      method: 'POST',
      body: JSON.stringify({
        method: 'saveCurrencyData',
        spreadsheetId,
        params: { data, type, sheetName }
      }),
      headers: { 'Content-Type': 'application/json' }
    });
    const result = await response.json();
    domElements.loader.style.display = "none";
    if (result.message) {
      const messageLines = result.message.split('\n');
      const savedCurrencies = result.savedCurrencies || [];
      const failedCurrencies = result.failedCurrencies || [];
      messageLines.forEach(line => {
        let currency = null;
        const msgType = line.includes('✅') ? 'success' : line.includes('❌') ? 'error' : 'warning';
        if (msgType === 'success') {
          currency = savedCurrencies.find(curr => line.includes(curr));
        } else if (msgType === 'error') {
          currency = failedCurrencies.find(curr => line.includes(curr));
        }
        if (!currency) {
          const normalizedLine = line.replace(/\s+/g, ' ').trim().toUpperCase();
          currencies.forEach(curr => {
            if (normalizedLine.includes(curr)) {
              currency = curr;
            }
          });
        }
        showCurrencyMessage(line, msgType, currency);
      });

      if (result.success && failedCurrencies.length === 0) {
        // Звуковий сигнал успіху
        try {
          const audioContext = new (window.AudioContext || window.webkitAudioContext)();
          const oscillator1 = audioContext.createOscillator();
          oscillator1.type = 'sine';
          oscillator1.frequency.setValueAtTime(1000, audioContext.currentTime);
          oscillator1.connect(audioContext.destination);
          oscillator1.start();
          oscillator1.stop(audioContext.currentTime + 0.1);

          const oscillator2 = audioContext.createOscillator();
          oscillator2.type = 'sine';
          oscillator2.frequency.setValueAtTime(1200, audioContext.currentTime + 0.07);
          oscillator2.connect(audioContext.destination);
          oscillator2.start(audioContext.currentTime + 0.1);
          oscillator2.stop(audioContext.currentTime + 0.22);
        } catch (e) {
          console.error('Помилка відтворення звуку успіху:', e);
        }
      }
    } else {
      showCurrencyMessage(result.message || '❌ Невідома помилка', result.success ? 'success' : 'error');
    }
    if (result.success) {
      clearFields();
      loadHistory();
    }
    sessionStorage.removeItem('pendingTransaction');
    domElements.saveButton.disabled = false;
    domElements.confirmButton.disabled = false;
    isSubmitting = false;
  } catch (error) {
    setTimeout(() => {
      submitWithRetry(data, type, sheetName, retries - 1, delay);
    }, delay);
  }
}

// Підтвердження відправки
function submitConfirmed() {
  if (isSubmitting) {
    showCurrencyMessage('⏳ Зачекайте, операція вже виконується!', 'warning');
    return;
  }

  isSubmitting = true;
  currencyMessages = [];
  displayCurrencyMessages();
  closeModal();
  domElements.saveButton.disabled = true;
  domElements.confirmButton.disabled = true;

  submitWithRetry(
    dataToSubmit,
    domElements.operation.value,
    domElements.dateInput.value
  );
}

// Відображення повідомлень для валют
function showCurrencyMessage(message, type, currency = null) {
  currencyMessages.push({ message, type, currency });
  triggerErrorFeedback(message);
  displayCurrencyMessages();
}

// Відображення повідомлень для витрат
function showExpensesMessage(message, type, currency = null) {
  expensesMessages.push({ message, type, currency });
  triggerErrorFeedback(message);
  displayExpensesMessages();
}

// Відображення повідомлень для історії
function showHistoryMessage(message, type, currency = null) {
  historyMessages.push({ message, type, currency });
  triggerErrorFeedback(message);
  displayHistoryMessages();
}

// Відображення повідомлень валют
function displayCurrencyMessages() {
  let html = '<ul>';
  currencyMessages.forEach(msg => {
    let displayMessage = msg.message;
    if (msg.currency) {
      displayMessage = displayMessage.replace(
        msg.currency,
        `<span class="currency-span" data-currency="${msg.currency}">${msg.currency}</span>`
      );
    }
    html += `<li class="${msg.type}">${displayMessage}</li>`;
  });
  html += '</ul>';
  domElements.messageList.innerHTML = html;
  domElements.messageList.scrollTop = domElements.messageList.scrollHeight;
}

// Відображення повідомлень витрат
function displayExpensesMessages() {
  let html = '<ul>';
  expensesMessages.forEach(msg => {
    let displayMessage = msg.message;
    if (msg.currency) {
      displayMessage = displayMessage.replace(
        msg.currency,
        `<span class="currency-span" data-currency="${msg.currency}">${msg.currency}</span>`
      );
    }
    html += `<li class="${msg.type}">${displayMessage}</li>`;
  });
  html += '</ul>';
  domElements.expensesMessageList.innerHTML = html;
  domElements.expensesMessageList.scrollTop = domElements.expensesMessageList.scrollHeight;
}

// Відображення повідомлень історії
function displayHistoryMessages() {
  let html = '<ul>';
  historyMessages.forEach(msg => {
    let displayMessage = msg.message;
    if (msg.currency) {
      displayMessage = displayMessage.replace(
        msg.currency,
        `<span class="currency-span" data-currency="${msg.currency}">${msg.currency}</span>`
      );
    }
    html += `<li class="${msg.type}">${displayMessage}</li>`;
  });
  html += '</ul>';
  domElements.historyMessageList.innerHTML = html;
  domElements.historyMessageList.scrollTop = domElements.historyMessageList.scrollHeight;
}

// Закриття модального вікна підтвердження
function closeModal() {
  domElements.confirmModal.style.display = "none";
  sessionStorage.removeItem('pendingTransaction');
}

// Закриття модального вікна скасування
function closeUndoModal() {
  domElements.undoModal.style.display = "none";
}

// Підготовка модального вікна скасування
function prepareUndoModal(data) {
  domElements.undoTableBody.innerHTML = `
    <tr>
      <td class="currency-cell" data-currency="${data.currency}">${data.currency}</td>
      <td>${data.amount}</td>
      <td>${data.rate}</td>
      <td>${data.time}</td>
      <td>${data.grn}</td>
    </tr>
  `;
  
  domElements.undoPreviewHeader.innerHTML = `<div class="modal-operation-header ${data.type === 'Купівля' ? 'buy' : 'sell'}">
    <div><b>Операція:</b> ${data.type}</div>
    <div><b>Лист:</b> ${data.sheet}</div>
  </div>`;
  
  domElements.undoModal.style.display = "block";
  domElements.undoConfirmButton.setAttribute('data-transaction-id', data.transactionId);
  domElements.undoConfirmButton.setAttribute('data-currency', data.currency);
  domElements.undoConfirmButton.setAttribute('data-amount', data.amount);
  domElements.undoConfirmButton.setAttribute('data-rate', data.rate);
  domElements.undoConfirmButton.setAttribute('data-sheet', data.sheet);
}

// Підтвердження скасування
async function confirmUndo() {
  const params = {
    transactionId: domElements.undoConfirmButton.getAttribute('data-transaction-id'),
    currency: domElements.undoConfirmButton.getAttribute('data-currency'),
    amount: domElements.undoConfirmButton.getAttribute('data-amount'),
    rate: domElements.undoConfirmButton.getAttribute('data-rate'),
    sheet: domElements.undoConfirmButton.getAttribute('data-sheet')
  };
  
  domElements.loader.style.display = "block";
  closeUndoModal();
  
  const spreadsheetId = document.getElementById('spreadsheetSelect').value;
  try {
    const response = await fetch(WEB_APP_URL, {
      method: 'POST',
      body: JSON.stringify({
        method: 'undoEntry',
        spreadsheetId,
        params
      }),
      headers: { 'Content-Type': 'application/json' }
    });
    const result = await response.json();
    domElements.loader.style.display = "none";
    showHistoryMessage(result.message, result.success ? 'success' : 'error', params.currency);
    if (result.success) {
      loadHistory();
    }
  } catch (error) {
    domElements.loader.style.display = "none";
    showHistoryMessage(`❌ Помилка скасування: ${error.message}`, 'error');
  }
}

// Очищення полів форми
function clearFields() {
  currencies.forEach(currency => {
    document.getElementById(`amount_${currency}`).value = '';
    document.getElementById(`rate_${currency}`).value = '';
    document.getElementById(`grn_${currency}`).value = '';
  });
  domElements.totalGrn.value = '';
}

// Показ модального вікна вибору точки
function showPointModal() {
  domElements.pointModal.style.display = "block";
}

// Закриття модального вікна вибору точки
function closePointModal() {
  domElements.pointModal.style.display = "none";
}

// Підтягування даних з точки
async function fetchPointData() {
  const pointId = domElements.pointSelect.value;
  const pointName = domElements.pointSelect.options[domElements.pointSelect.selectedIndex].text;
  const fetchButton = document.getElementById('fetchDataButton');
  
  if (!pointId) {
    showCurrencyMessage('❌ Виберіть точку', 'error');
    return;
  }
  
  // Вимикаємо кнопку
  fetchButton.disabled = true;
  
  try {
    const response = await fetch(WEB_APP_URL, {
      method: 'POST',
      body: JSON.stringify({
        method: 'fetchPointData',
        params: { pointId }
      }),
      headers: { 'Content-Type': 'application/json' }
    });
    const result = await response.json();
    fetchButton.disabled = false;
    if (result.success) {
      const data = result.data;
      Object.keys(data).forEach(currency => {
        document.getElementById(`amount_${currency}`).value = data[currency].amount;
        document.getElementById(`rate_${currency}`).value = data[currency].rate;
        calculateTotal(currency);
      });
      showCurrencyMessage(`Дані успішно отримано з точки ${pointName}`, 'success');
      closePointModal();
    } else {
      showCurrencyMessage(result.message, 'error');
    }
  } catch (error) {
    fetchButton.disabled = false;
    showCurrencyMessage(`❌ Помилка: ${error.message}`, 'error');
  }
}