/******************* CONFIGURATION *******************/
const CONFIG = {
  TARGET_LEVEL: 10,         // Конечный уровень вещи (состояния 0...10; E(10)=0)
  INITIAL_COST: 1000,       // Начальное приближение для E(state) для не терминальных состояний
  MAX_POLICY_ITER: 5000,    // Максимальное число итераций по политике
  TOLERANCE: 1e-12,          // Критерий сходимости
  
  // Параметры динамики улучшения:
  // При успехе: прирост +1 с 90%, +2 с 7%, +3 с 3%
  SUCCESS_DISTRIBUTION: [   
    { delta: 1, prob: 0.90 },
    { delta: 2, prob: 0.07 },
    { delta: 3, prob: 0.03 }
  ],
  // При неудаче: если состояние равно 0, остаёмся на 0; иначе – снижаем на 1
  failureDistribution: function(state) {
    if (state === 0) {
      return [{ delta: 0, prob: 1.0 }];
    } else {
      return [{ delta: -1, prob: 1.0 }];
    }
  },
  
  // Лист, где находятся данные (и куда будут выводиться результаты)
  sheetStones: "Камни",
  
  // Настройки вывода результатов: начальная ячейка для таблицы результатов в листе Камни
  RESULT_OUTPUT: {
    startRow: 2,    // строка, с которой начинается вывод результатов
    startCol: 5     // столбец, с которого начинается вывод (E = 5-й столбец)
  }
};

/******************* INPUT DATA *******************/
// Данные для камней – они находятся в листе "Камни" в диапазоне:
// Колонка A: Уровень камня (66...113)
// Колонка B: Цена камня
// Колонка C: Шанс успеха (например, 0.017 ... 0.85)
// Здесь в скрипте данные будут считываться из листа

/******************* MAIN FUNCTION *******************/
function calculateOptimalStrategy() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetStones = ss.getSheetByName(CONFIG.sheetStones);
  
  // Чтение входных данных с листа "Камни":
  // Предполагается, что заголовки находятся в первой строке, данные начинаются со строки 2, столбцы A:C.
  const lastRow = sheetStones.getLastRow();
  const stonesData = sheetStones.getRange(2, 1, lastRow - 1, 3).getValues();
  
  // Формируем массив объектов: { stone_level, price, p_success }
  const stones = stonesData.map(function(row) {
    return {
      stone_level: row[0],
      price: row[1],
      p_success: row[2]
    };
  });
  
  const numStates = CONFIG.TARGET_LEVEL; // состояния 0..TARGET_LEVEL-1 (для TARGET_LEVEL E=0)
  let E_vals = new Array(numStates).fill(CONFIG.INITIAL_COST);
  let policy = new Array(numStates).fill(0);
  
  // Инициализируем начальную политику: выбираем для всех состояний камень с наибольшей вероятностью успеха
  let bestIndex = stones.reduce((maxIndex, stone, index, arr) => 
                      stone.p_success > arr[maxIndex].p_success ? index : maxIndex, 0);
  policy.fill(bestIndex);
  
  // Функция расчёта Q-значения для состояния state при использовании заданного камня
  function Q_value(state, stone, E_vals) {
    const cost = stone.price;
    const p = stone.p_success;
    
    let successCost = 0;
    CONFIG.SUCCESS_DISTRIBUTION.forEach(function(item) {
      let nextState = Math.min(state + item.delta, CONFIG.TARGET_LEVEL);
      if (nextState < CONFIG.TARGET_LEVEL) {
        successCost += item.prob * E_vals[nextState];
      }
      // Если nextState == TARGET_LEVEL, считается E=0.
    });
    
    let failCost = 0;
    CONFIG.failureDistribution(state).forEach(function(item) {
      let nextState = state + item.delta;
      nextState = Math.max(0, Math.min(nextState, CONFIG.TARGET_LEVEL));
      if (nextState < CONFIG.TARGET_LEVEL) {
        failCost += item.prob * E_vals[nextState];
      }
    });
    
    return cost + p * successCost + (1 - p) * failCost;
  }
  
  // Policy Iteration
  let policyStable = false;
  let iteration = 0;
  while (!policyStable && iteration < CONFIG.MAX_POLICY_ITER) {
    iteration++;
    // ----- Policy Evaluation -----
    const A = [];
    const b = [];
    for (let state = 0; state < numStates; state++) {
      const stone = stones[policy[state]];
      const cost = stone.price;
      const p = stone.p_success;
      let row = new Array(numStates).fill(0);
      row[state] = 1;  // коэффициент при E(state)
      
      // Учет успеха: для каждого delta из SUCCESS_DISTRIBUTION
      CONFIG.SUCCESS_DISTRIBUTION.forEach(function(item) {
        let nextState = Math.min(state + item.delta, CONFIG.TARGET_LEVEL);
        if (nextState < CONFIG.TARGET_LEVEL) {
          row[nextState] -= p * item.prob;
        }
      });
      
      // Учет неудачи: если state == 0 – остаёмся на 0, иначе берём E(state-1)
      if (state === 0) {
        row[0] -= (1 - p);
      } else {
        row[state - 1] -= (1 - p);
      }
      
      A.push(row);
      b.push(cost);
    }
    
    let E_new;
    try {
      E_new = gaussSolve(A, b);
    } catch (error) {
      Logger.log("Ошибка при решении системы уравнений: " + error);
      return;
    }
    
    let maxDiff = Math.max(...E_new.map((val, i) => Math.abs(val - E_vals[i])));
    E_vals = E_new.slice();
    
    // ----- Policy Improvement -----
    policyStable = true;
    for (let state = 0; state < numStates; state++) {
      let currentAction = policy[state];
      let currentQ = Q_value(state, stones[currentAction], E_vals);
      let bestAction = currentAction;
      for (let i = 0; i < stones.length; i++) {
        let q_val = Q_value(state, stones[i], E_vals);
        if (q_val < currentQ) {
          currentQ = q_val;
          bestAction = i;
        }
      }
      if (bestAction !== policy[state]) {
        policy[state] = bestAction;
        policyStable = false;
      }
    }
    
    if (maxDiff < CONFIG.TOLERANCE) break;
  }
  
  /************ OUTPUT RESULTS ************/
  // Формируем результирующую таблицу:
  // Заголовок: ["Уровень вещи", "Камень (уровень)", "Цена", "Шанс успеха", "E(state)"]
  let output = [["Уровень вещи", "Камень lvl", "Цена", "Шанс успеха", "E(state)"]];
  for (let state = 0; state < numStates; state++) {
    let stone = stones[policy[state]];
    output.push([state, stone.stone_level, stone.price, stone.p_success, E_vals[state]]);
  }
  // Добавляем терминальное состояние
  // output.push([CONFIG.TARGET_LEVEL, "-", "-", "-", 0]);
  
  // Записываем результаты в лист "Камни", начиная со второй строки, 5-го столбца
  // Сначала очищаем область вывода (например, диапазон 2:100, столбцы E:И)
  const outputStartRow = CONFIG.RESULT_OUTPUT.startRow;
  const outputStartCol = CONFIG.RESULT_OUTPUT.startCol;
  const numRows = output.length;
  const numCols = output[0].length;
  
  // Очищаем область вывода (можно настроить диапазон по необходимости)
  sheetStones.getRange(outputStartRow, outputStartCol, sheetStones.getMaxRows()-outputStartRow+1, numCols).clearContent();
  
  sheetStones.getRange(outputStartRow, outputStartCol, numRows, numCols).setValues(output);
  SpreadsheetApp.flush();
  Logger.log("Оптимальная стратегия пересчитана за " + iteration + " итераций.");
}

/******************* TRIGGER SETUP *******************/
/**
 * Функция, которая устанавливает onEdit триггер для листа "Камни".
 * При изменении данных в листе "Камни" автоматически запускается пересчёт.
 */
function createOnEditTrigger() {
  // Удаляем все предыдущие триггеры для onEdit, чтобы не было дублирования
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === "onEditTrigger") {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  ScriptApp.newTrigger("onEditTrigger")
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();
}

/**
 * Функция-триггер, которая срабатывает при редактировании.
 * Если редактируется лист "Камни", запускается пересчёт оптимальной стратегии.
 */
function onEditTrigger(e) {
  const editedSheet = e.range.getSheet();
  if (editedSheet.getName() === CONFIG.sheetStones) {
    // Можно добавить проверку, что изменение произошло в нужном диапазоне (например, столбцы A:C)
    calculateOptimalStrategy();
  }
}

/******************* GAUSS SOLVER *******************/
/**
 * Функция решения системы линейных уравнений методом Гаусса.
 * Принимает A – массив массивов коэффициентов, b – массив свободных членов.
 */
function gaussSolve(A, b) {
  const n = A.length;
  let M = A.map(function(row) { return row.slice(); });
  let B = b.slice();
  
  // Прямой ход
  for (let k = 0; k < n; k++) {
    let maxRow = k;
    for (let i = k + 1; i < n; i++) {
      if (Math.abs(M[i][k]) > Math.abs(M[maxRow][k])) {
        maxRow = i;
      }
    }
    [M[k], M[maxRow]] = [M[maxRow], M[k]];
    [B[k], B[maxRow]] = [B[maxRow], B[k]];
    
    if (Math.abs(M[k][k]) < 1e-12) {
      throw new Error("Система вырождена");
    }
    
    for (let i = k + 1; i < n; i++) {
      let factor = M[i][k] / M[k][k];
      for (let j = k; j < n; j++) {
        M[i][j] -= factor * M[k][j];
      }
      B[i] -= factor * B[k];
    }
  }
  
  // Обратный ход
  let x = new Array(n);
  for (let i = n - 1; i >= 0; i--) {
    let sum = 0;
    for (let j = i + 1; j < n; j++) {
      sum += M[i][j] * x[j];
    }
    x[i] = (B[i] - sum) / M[i][i];
  }
  return x;
}
