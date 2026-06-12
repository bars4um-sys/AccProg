// Хранилище данных
let data = {
    table1: null,
    table2: null,
    file1Name: '',
    file2Name: ''
};

// DOM элементы
const file1Input = document.getElementById('file1');
const file2Input = document.getElementById('file2');
const fileName1Display = document.getElementById('fileName1');
const fileName2Display = document.getElementById('fileName2');
const compareBtn = document.getElementById('compareBtn');
const resetBtn = document.getElementById('resetBtn');
const resultsSection = document.getElementById('resultsSection');
const resultContent = document.getElementById('resultContent');
const resultsCount = document.getElementById('resultsCount');
const errorMessage = document.getElementById('errorMessage');

// Обработчики файлов
file1Input.addEventListener('change', (e) => handleFileUpload(e, 1));
file2Input.addEventListener('change', (e) => handleFileUpload(e, 2));
compareBtn.addEventListener('click', performComparison);
resetBtn.addEventListener('click', resetApp);

/**
 * Очистка данных: удаление пустых строк и итоговых данных
 */
function cleanData(jsonData) {
    return jsonData.filter(row => {
        // Пропускаем пустые строки
        if (!row || Object.keys(row).length === 0) {
            return false;
        }

        // Получаем первое текстовое значение строки (обычно название компании)
        let firstTextValue = '';
        for (const key in row) {
            if (typeof row[key] === 'string') {
                firstTextValue = row[key].trim().toLowerCase();
                break;
            }
        }

        // Пропускаем строки с итогами, заголовками и пустыми значениями
        const skipPatterns = [
            'итого',
            'всего',
            'total',
            'sum',
            'сумма',
            'grand total',
            'конец',
            'end',
            'footer',
            'подитог',
            'subtotal',
            'наименование', // заголовок
            'название',      // заголовок
            'компания'        // может быть заголовок
        ];

        for (const pattern of skipPatterns) {
            if (firstTextValue.includes(pattern)) {
                return false;
            }
        }

        // Пропускаем строки, где нет числовых значений
        let hasNumber = false;
        for (const key in row) {
            if (typeof row[key] === 'number' && !isNaN(row[key])) {
                hasNumber = true;
                break;
            }
        }

        return hasNumber && firstTextValue !== '';
    });
}

/**
 * Парсинг CSV файла
 */
function parseCSV(csvText) {
    const lines = csvText.trim().split('\n');
    if (lines.length < 2) {
        throw new Error('Таблица должна содержать как минимум заголовок и одну строку данных');
    }

    // Парсим заголовок
    const headers = lines[0].split(',').map(h => h.trim());
    
    if (headers.length < 2) {
        throw new Error('Таблица должна содержать как минимум 2 колонки');
    }

    // Парсим данные
    let data = [];
    for (let i = 1; i < lines.length; i++) {
        const line = lines[i].trim();
        if (!line) continue; // Пропускаем пустые строки
        
        const values = line.split(',').map(v => v.trim());
        const row = {};
        
        headers.forEach((header, index) => {
            let value = values[index] || '';
            
            // Пытаемся преобразовать в число
            if (!isNaN(value) && value !== '') {
                value = parseFloat(value);
            }
            
            row[header] = value;
        });
        
        data.push(row);
    }

    // Очищаем данные от итоговых строк
    data = cleanData(data);

    return data;
}

/**
 * Парсинг Excel/CSV файла
 */
function handleFileUpload(event, tableNumber) {
    const file = event.target.files[0];
    if (!file) return;

    clearError();
    const reader = new FileReader();
    
    reader.onload = (e) => {
        try {
            let jsonData;

            // Определяем тип файла по расширению
            const fileName = file.name.toLowerCase();
            const isExcel = fileName.endsWith('.xlsx') || fileName.endsWith('.xls');
            const isCSV = fileName.endsWith('.csv');

            if (isExcel) {
                // Парсим Excel файл
                const data_array = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data_array, { type: 'array' });
                
                // Берем первый лист
                const sheetName = workbook.SheetNames[0];
                const sheet = workbook.Sheets[sheetName];
                
                // Читаем данные с заголовками
                jsonData = XLSX.utils.sheet_to_json(sheet);
                
                // Очищаем данные от итоговых строк
                jsonData = cleanData(jsonData);
            } else if (isCSV) {
                // Парсим CSV
                const csvText = e.target.result;
                jsonData = parseCSV(csvText);
            } else {
                // По умолчанию пытаемся как CSV
                const csvText = e.target.result;
                jsonData = parseCSV(csvText);
            }
            
            if (jsonData.length === 0) {
                showError('Таблица пуста или не содержит данных');
                return;
            }

            // Сохраняем данные и имя файла
            if (tableNumber === 1) {
                data.table1 = jsonData;
                data.file1Name = file.name;
                fileName1Display.textContent = `✓ ${file.name}`;
                fileName1Display.classList.add('loaded');
            } else {
                data.table2 = jsonData;
                data.file2Name = file.name;
                fileName2Display.textContent = `✓ ${file.name}`;
                fileName2Display.classList.add('loaded');
            }

            // Активируем кнопку сравнения если оба файла загружены
            updateCompareButtonState();
            
        } catch (error) {
            showError(`Ошибка при чтении файла: ${error.message}`);
            console.error(error);
        }
    };

    reader.onerror = () => {
        showError('Ошибка при чтении файла');
    };

    // Для CSV читаем как текст, для Excel как ArrayBuffer
    const fileName = file.name.toLowerCase();
    if (fileName.endsWith('.csv')) {
        reader.readAsText(file, 'utf-8');
    } else {
        reader.readAsArrayBuffer(file);
    }
}

/**
 * Обновление состояния кнопки сравнения
 */
function updateCompareButtonState() {
    compareBtn.disabled = !data.table1 || !data.table2;
}

/**
 * Агрегирование данных: суммирование по компаниям
 */
function aggregateData(tableData) {
    const aggregated = {};

    for (const row of tableData) {
        // Находим столбцы с наименованием и суммой
        // Ищем колонки, которые содержат названия компаний (обычно первая текстовая колонка)
        // и суммы (обычно числовая колонка)
        
        let companyName = null;
        let amount = null;

        // Ищем наименование (первый текстовый столбец)
        for (const key in row) {
            if (typeof row[key] === 'string' && !companyName) {
                companyName = row[key].trim();
            }
        }

        // Ищем сумму (первый числовой столбец или столбец с "Сумма" в названии)
        const keys = Object.keys(row);
        for (const key of keys) {
            const value = row[key];
            if (typeof value === 'number' && amount === null) {
                amount = value;
                break;
            } else if (key.toLowerCase().includes('сумма') || key.toLowerCase().includes('sum')) {
                amount = parseFloat(value) || 0;
                break;
            }
        }

        if (companyName && amount !== null) {
            if (!aggregated[companyName]) {
                aggregated[companyName] = 0;
            }
            aggregated[companyName] += amount;
        }
    }

    return aggregated;
}

/**
 * Сравнение двух агрегированных таблиц
 */
function compareAggregatedData(agg1, agg2) {
    const differences = [];

    // Получаем все уникальные компании из обеих таблиц
    const allCompanies = new Set([
        ...Object.keys(agg1),
        ...Object.keys(agg2)
    ]);

    // Сравниваем значения
    for (const company of allCompanies) {
        const amount1 = agg1[company] || 0;
        const amount2 = agg2[company] || 0;

        if (amount1 !== amount2) {
            differences.push({
                company: company,
                amount1: amount1,
                amount2: amount2,
                difference: Math.abs(amount1 - amount2)
            });
        }
    }

    // Сортируем по названию компании
    differences.sort((a, b) => a.company.localeCompare(b.company, 'ru'));

    return differences;
}

/**
 * Выполнение сравнения
 */
function performComparison() {
    if (!data.table1 || !data.table2) {
        showError('Загрузите обе таблицы');
        return;
    }

    clearError();

    try {
        // Агрегируем данные из обеих таблиц
        const agg1 = aggregateData(data.table1);
        const agg2 = aggregateData(data.table2);

        // Сравниваем агрегированные данные
        const differences = compareAggregatedData(agg1, agg2);

        // Отображаем результаты
        displayResults(differences);

    } catch (error) {
        showError(`Ошибка при сравнении: ${error.message}`);
        console.error(error);
    }
}

/**
 * Отображение результатов
 */
function displayResults(differences) {
    resultsSection.classList.add('show');

    if (differences.length === 0) {
        resultsCount.textContent = 'Расхождений не найдено ✓';
        resultContent.innerHTML = '<div class="no-differences">Обе таблицы содержат одинаковые суммы по всем компаниям</div>';
        return;
    }

    resultsCount.textContent = `Обнаружено ${differences.length} компании(й) с расхождениями:`;

    const list = document.createElement('ul');
    list.className = 'differences-list';

    for (const diff of differences) {
        const item = document.createElement('li');
        item.className = 'difference-item';

        const companyDiv = document.createElement('div');
        companyDiv.className = 'company-name';
        companyDiv.textContent = diff.company;

        const amountsDiv = document.createElement('div');
        amountsDiv.className = 'amounts';

        const amount1Div = document.createElement('div');
        amount1Div.className = 'amount-item';
        amount1Div.innerHTML = `
            <span class="amount-label">Таблица 1:</span>
            <span class="amount-value">${formatNumber(diff.amount1)}</span>
        `;

        const amount2Div = document.createElement('div');
        amount2Div.className = 'amount-item';
        amount2Div.innerHTML = `
            <span class="amount-label">Таблица 2:</span>
            <span class="amount-value">${formatNumber(diff.amount2)}</span>
        `;

        const differenceDiv = document.createElement('div');
        differenceDiv.className = 'amount-item';
        differenceDiv.innerHTML = `
            <span class="amount-label">Разница:</span>
            <span class="amount-value" style="color: #dc3545;">${formatNumber(diff.difference)}</span>
        `;

        amountsDiv.appendChild(amount1Div);
        amountsDiv.appendChild(amount2Div);
        amountsDiv.appendChild(differenceDiv);

        item.appendChild(companyDiv);
        item.appendChild(amountsDiv);

        list.appendChild(item);
    }

    resultContent.innerHTML = '';
    resultContent.appendChild(list);
}

/**
 * Форматирование числа
 */
function formatNumber(num) {
    if (num === null || num === undefined) return '0';
    return num.toLocaleString('ru-RU', {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
    });
}

/**
 * Очистка приложения
 */
function resetApp() {
    data = {
        table1: null,
        table2: null,
        file1Name: '',
        file2Name: ''
    };

    file1Input.value = '';
    file2Input.value = '';
    fileName1Display.textContent = '';
    fileName1Display.classList.remove('loaded');
    fileName2Display.textContent = '';
    fileName2Display.classList.remove('loaded');
    
    resultsSection.classList.remove('show');
    resultContent.innerHTML = '';
    resultsCount.textContent = '';
    
    updateCompareButtonState();
    clearError();
}

/**
 * Показ ошибки
 */
function showError(message) {
    errorMessage.textContent = message;
    errorMessage.classList.add('show');
}

/**
 * Очистка ошибки
 */
function clearError() {
    errorMessage.classList.remove('show');
    errorMessage.textContent = '';
}
