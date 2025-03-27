document.addEventListener('DOMContentLoaded', function() {
    // Создаем скрытые input'ы для файлов
    const file1Input = document.createElement('input');
    file1Input.type = 'file';
    file1Input.accept = '.xlsx,.xls';
    file1Input.style.display = 'none';

    const file2Input = document.createElement('input');
    file2Input.type = 'file';
    file2Input.accept = '.xlsx,.xls';
    file2Input.style.display = 'none';

    const file3Input = document.createElement('input');
    file3Input.type = 'file';
    file3Input.accept = '.xlsx,.xls';
    file3Input.style.display = 'none';

    document.body.appendChild(file1Input);
    document.body.appendChild(file2Input);
    document.body.appendChild(file3Input);

    // Переменные для хранения данных
    let df1 = [], df2 = [], df3 = [];
    let filesReady = 0;
    const totalFiles = 3;

    // Обновление состояния кнопки
    function updateMergeButton() {
        document.getElementById('mergeBtn').disabled = filesReady !== totalFiles;
    }

    // Обработчики кнопок выбора файлов
    document.getElementById('selectFile1').addEventListener('click', () => file1Input.click());
    document.getElementById('selectFile2').addEventListener('click', () => file2Input.click());
    document.getElementById('selectFile3').addEventListener('click', () => file3Input.click());

    // Общая функция обработки файлов
    function handleFileSelect(input, labelId, dataArray) {
        return function(e) {
            if (e.target.files.length > 0) {
                const file = e.target.files[0];
                const label = document.getElementById(labelId);
                label.textContent = `Selected: ${file.name}`;
                label.classList.add('loaded');

                const reader = new FileReader();
                reader.onload = function(e) {
                    try {
                        const data = new Uint8Array(e.target.result);
                        const workbook = XLSX.read(data, {type: 'array'});
                        const sheet = workbook.Sheets[workbook.SheetNames[0]];
                        dataArray.length = 0;
                        dataArray.push(...XLSX.utils.sheet_to_json(sheet, {header: 1}));
                        filesReady++;
                        updateMergeButton();
                    } catch (error) {
                        label.textContent = `Error: ${file.name}`;
                        label.classList.add('error');
                    }
                };
                reader.readAsArrayBuffer(file);
            }
        };
    }

    // Назначаем обработчики
    file1Input.addEventListener('change', handleFileSelect(file1Input, 'file1Label', df1));
    file2Input.addEventListener('change', handleFileSelect(file2Input, 'file2Label', df2));
    file3Input.addEventListener('change', handleFileSelect(file3Input, 'file3Label', df3));

    // Основная функция обработки
    document.getElementById('mergeBtn').addEventListener('click', async () => {
        if (filesReady !== totalFiles) return;

        try {
            // Обработка данных (оригинальная логика)
            df1 = df1.filter(row => row.some(cell => cell !== null && cell !== ''));
            df2 = df2.filter(row => row.some(cell => cell !== null && cell !== ''));
            df3 = df3.filter(row => row.some(cell => cell !== null && cell !== ''));

            df1 = df1.slice(1);
            df2 = df2.slice(2);
            df3 = df3.slice(2);

            df1 = df1.map(item => {
                if (item[0] === 'Артикул' && !item[1] && !item[2]) return [item[item.length - 1]];
                return item;
            });

            df2 = df2.map(item => {
                if (item[0] && !item[4] && !item[6]) return [item[0]];
                return item;
            });

            df3 = df3.map(item => {
                if (item[0] && !item[1] && !item[2]) return [item[0]];
                return item;
            });

            // Извлечение колонок (оригинальная логика)
            const cycloneColumn = df1.map(row => row[0]);
            const mysteryColumn = df2.map(row => row[2] ?? row[0]);
            const thirdFileColumn = df3.map(row => row[0]?.toString() ?? '');

            const fourthColumnFile1 = df1.map(row => row[3]);
            const secondColumnFile2 = df2.map(row => row[2] ? row[1] : '');
            const thirdFileBrandColumn = df3.map(row => row[1]);

            const fifthColumnFile1 = df1.map(row => row[4]);
            const seventhColumnFile1 = df1.map(row => row[6]);
            const fifthColumnFile2 = df2.map(row => row[4]);
            const thirdFileModelColumn = df3.map(row => row[2]);

            const ninthColumnFile1 = df1.map(row => row[8]);
            const seventhColumnFile2 = df2.map(row => row[6]);
            const thirdFilePriceColumn = df3.map(row => row[5] ? Math.ceil((row[5] * 110))/100 : '');
            const thirdFileRetailPriceColumn = df3.map(row => row[6]);

            const plusSymbolColumn = df1.map(i => isNaN(i[6]) ? '' : '+');
            const tenthColumnFile2 = df2.map(row => row[9]);
            const thirdFileAvailabilityColumn = df3.map(row => row[1] ? isNaN(row[3]) ? '-' : '+' : '');

            // Объединение данных
            const combined = {
                A: [...cycloneColumn, ...mysteryColumn, ...thirdFileColumn],
                B: [...fourthColumnFile1, ...Array(mysteryColumn.length).fill('')],
                C: [...fifthColumnFile1, ...secondColumnFile2],
                D: [...seventhColumnFile1, ...fifthColumnFile2, ...thirdFilePriceColumn],
                E: [...ninthColumnFile1, ...seventhColumnFile2, ...thirdFileRetailPriceColumn],
                F: [...plusSymbolColumn, ...tenthColumnFile2, ...thirdFileAvailabilityColumn]
            };

            // Создаем книгу Excel
            const workbook = new ExcelJS.Workbook();
            const worksheet = workbook.addWorksheet('Combined Data');

            // Добавляем заголовки
            worksheet.addRow(['Артикул', 'Бренд', 'Модель', 'Оптова USD', 'Роздрібна UAH', 'Наявність']);

            // Стили для заголовков
            worksheet.getRow(1).eachCell(cell => {
                cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'FF8C00' }
                };
                cell.font = {
                    bold: true,
                    color: { argb: 'FFFFFF' },
                    size: 14
                };
                cell.alignment = {
                    vertical: 'middle',
                    horizontal: 'center'
                };
            });

            // Добавляем данные
            const maxLen = Math.max(...Object.values(combined).map(c => c.length));
            for (let i = 0; i < maxLen; i++) {
                const row = worksheet.addRow([
                    combined.A[i] || '',
                    combined.B[i] || '',
                    combined.C[i] || '',
                    combined.D[i] || '',
                    combined.E[i] || '',
                    combined.F[i] || ''
                ]);

                // Проверка на строку с одним значением
                const filledCells = row.values.slice(1).filter(v => v !== '').length;
                if (filledCells === 1) {
                    row.eachCell(cell => {
                        cell.fill = {
                            type: 'pattern',
                            pattern: 'solid',
                            fgColor: { argb: 'FFFF00' }
                        };
                        cell.font = { bold: true };
                    });
                    worksheet.mergeCells(`A${row.number}:F${row.number}`);
                }

                // Выравнивание для всех ячеек
                row.eachCell(cell => {
                    cell.alignment = {
                        vertical: 'middle',
                        horizontal: 'center'
                    };
                });
            }

            // Настройка ширины колонок
            worksheet.columns = [
                { width: 15 }, { width: 20 }, { width: 25 },
                { width: 12 }, { width: 15 }, { width: 10 }
            ];

            // Сохранение файла
            const buffer = await workbook.xlsx.writeBuffer();
            const blob = new Blob([buffer], {
                type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            });
            saveAs(blob, `DSP-Sound_price_${new Date().toISOString().split('T')[0]}.xlsx`);

        } catch (error) {
            alert(`Error: ${error.message}`);
        }
    });
});
