const expressionEl = document.getElementById('expression');
const resultEl = document.getElementById('result');
const buttons = document.querySelectorAll('.btn');

let currentExpression = '';
let lastAnswer = '0';

function updateDisplay() {
  expressionEl.textContent = currentExpression || '0';
  resultEl.textContent = lastAnswer;
}

function isOperator(value) {
  return ['+', '-', '*', '/'].includes(value);
}

function sanitizeExpression(raw) {
  return raw
    .replace(/×/g, '*')
    .replace(/÷/g, '/')
    .replace(/−/g, '-')
    .replace(/--/g, '-')
    .replace(/\s+/g, '');
}

function evaluateExpression(rawExpression) {
  const safeExpression = sanitizeExpression(rawExpression);

  if (!safeExpression) {
    return '0';
  }

  if (!/^[0-9+\-*/.()]+$/.test(safeExpression)) {
    throw new Error('Некорректное выражение');
  }

  const result = Function(`"use strict"; return (${safeExpression});`)();

  if (!Number.isFinite(result)) {
    throw new Error('Ошибка вычисления');
  }

  return Number(result.toFixed(10)).toString();
}

function appendValue(value) {
  if (currentExpression === '0' && !isOperator(value) && value !== '.') {
    currentExpression = value;
  } else if (currentExpression === '' && isOperator(value)) {
    return;
  } else {
    currentExpression += value;
  }

  const lastChar = currentExpression[currentExpression.length - 1];
  if (value === '.' && currentExpression.match(/\d+\.\d*$/) === null && !/[+\-*/.]$/.test(currentExpression.slice(0, -1))) {
    // allow decimal input only once per number
  }

  const lastNumber = currentExpression.split(/[+\-*/]/).at(-1);
  if (value === '.' && lastNumber.includes('.')) {
    currentExpression = currentExpression.slice(0, -1);
  }

  if (isOperator(lastChar) && currentExpression.length > 1 && isOperator(currentExpression[currentExpression.length - 2])) {
    currentExpression = currentExpression.slice(0, -2) + lastChar;
  }

  lastAnswer = currentExpression || '0';
  updateDisplay();
}

function clearAll() {
  currentExpression = '';
  lastAnswer = '0';
  updateDisplay();
}

function deleteLast() {
  currentExpression = currentExpression.slice(0, -1);
  lastAnswer = currentExpression || '0';
  updateDisplay();
}

function calculate() {
  if (!currentExpression) {
    return;
  }

  if (isOperator(currentExpression[currentExpression.length - 1])) {
    currentExpression = currentExpression.slice(0, -1);
  }

  try {
    const result = evaluateExpression(currentExpression);
    lastAnswer = result;
    currentExpression = result;
  } catch (error) {
    lastAnswer = 'Ошибка';
    currentExpression = '';
  }

  updateDisplay();
}

buttons.forEach((button) => {
  button.addEventListener('click', () => {
    const value = button.dataset.value;
    const action = button.dataset.action;

    if (action === 'clear') {
      clearAll();
      return;
    }

    if (action === 'delete') {
      deleteLast();
      return;
    }

    if (action === 'equals') {
      calculate();
      return;
    }

    if (value) {
      appendValue(value);
    }
  });
});

window.addEventListener('keydown', (event) => {
  const { key } = event;
  const allowed = /^[0-9+\-*/.=]$/;

  if (allowed.test(key)) {
    event.preventDefault();
    if (key === '=') {
      calculate();
      return;
    }

    if (key === 'Enter') {
      calculate();
      return;
    }

    if (key === 'Backspace') {
      deleteLast();
      return;
    }

    if (key === 'Escape') {
      clearAll();
      return;
    }

    appendValue(key);
  }
});

updateDisplay();
