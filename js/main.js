document.addEventListener('DOMContentLoaded', function () {
    // Create hidden file inputs
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

    // Variables to store file data
    let file1 = null;
    let file2 = null;
    let file3 = null;
    let df1 = null; // Data from first file
    let df2 = null; // Data from second file
    let df3 = null; // Data from third file

    // Event listeners for buttons
    document.getElementById('selectFile1').addEventListener('click', function () {
        file1Input.click();
    });

    document.getElementById('selectFile2').addEventListener('click', function () {
        file2Input.click();
    });

    document.getElementById('selectFile3').addEventListener('click', function () {
        file3Input.click();
    });

    // Handle first file selection
    file1Input.addEventListener('change', function (e) {
        if (e.target.files.length > 0) {
            file1 = e.target.files[0];
            document.getElementById('file1Label').textContent = `Selected file: ${file1.name}`;

            const reader = new FileReader();
            reader.onload = function (e) {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, {type: 'array'});
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                df1 = XLSX.utils.sheet_to_json(firstSheet, {header: 1});
            };
            reader.readAsArrayBuffer(file1);
        }
    });

    // Handle second file selection
    file2Input.addEventListener('change', function (e) {
        if (e.target.files.length > 0) {
            file2 = e.target.files[0];
            document.getElementById('file2Label').textContent = `Selected file: ${file2.name}`;

            const reader = new FileReader();
            reader.onload = function (e) {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, {type: 'array'});
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                df2 = XLSX.utils.sheet_to_json(firstSheet, {header: 1});
            };
            reader.readAsArrayBuffer(file2);
        }
    });

    // Handle third file selection
    file3Input.addEventListener('change', function (e) {
        if (e.target.files.length > 0) {
            file3 = e.target.files[0];
            document.getElementById('file3Label').textContent = `Selected file: ${file3.name}`;

            const reader = new FileReader();
            reader.onload = function (e) {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, {type: 'array'});
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                df3 = XLSX.utils.sheet_to_json(firstSheet, {header: 1});
            };
            reader.readAsArrayBuffer(file3);
        }
    });

    // Merge button click handler
    document.getElementById('mergeBtn').addEventListener('click', function () {
        if (!df1 || !df2 || !df3 || df1.length === 0 || df2.length === 0 || df3.length === 0) {
            alert('Not all files are selected or they are empty');
            return;
        }

        try {
            // Remove empty rows from all files
            df1 = df1.filter(row => row.some(cell => cell !== null && cell !== ''));
            df2 = df2.filter(row => row.some(cell => cell !== null && cell !== ''));
            df3 = df3.filter(row => row.some(cell => cell !== null && cell !== ''));

            // Remove header rows
            df1 = df1.slice(1); // Remove first row from first file
            df2 = df2.slice(2); // Remove first two rows from second file
            df3 = df3.slice(2); // Remove first row from third file (adjust as needed)

            // Process rows with only "Артикул" (SKU) and empty values
            df1 = df1.map((item) => {
                if (item[0] === 'Артикул' && !item[1] && !item[2]) {
                    return [item[item.length - 1]];
                }
                return item;
            });

            df2 = df2.map((item) => {
                if (item[0] && !item[4] && !item[6]) {
                    return [item[0]];
                }
                return item;
            });

            // Process third file data (adjust column indexes as needed)
            df3 = df3.map((item) => {
                if (item[0] && !item[1] && !item[2]) {
                    return [item[0]];
                }
                return item;
            });

            // Extract needed columns by index
            const cycloneColumn = df1.map(row => row[0]); // Column A from file 1
            const mysteryColumn = df2.map(row => row[2] ?? row[0]); // Column C (fallback to A) from file 2
            const thirdFileColumn = df3.map(row => row[0] ? row[0].toString() : row[0]); // Column A from file 3 (adjust as needed)

            const fourthColumnFile1 = df1.map(row => row[3]); // Column D from first file
            const secondColumnFile2 = df2.map(row => row[2] ? row[0] : ''); // Column B from second file
            const thirdFileBrandColumn = df3.map(row => row[1]); // Column B from third file (adjust as needed)

            const fifthColumnFile1 = df1.map(row => row[4]); // Column E from first file
            const seventhColumnFile1 = df1.map(row => row[6]); // Column G from first file
            const fifthColumnFile2 = df2.map(row => row[4]); // Column E from second file
            const thirdFileModelColumn = df3.map(row => row[2]); // Column C from third file (adjust as needed)

            const ninthColumnFile1 = df1.map(row => row[8]); // Column I from first file
            const seventhColumnFile2 = df2.map(row => row[6]); // Column G from second file
            const thirdFilePriceColumn = df3.map(row => row[5] ? Math.ceil((row[5] * 110)) / 100 : ''); // Column F from third file (adjust as needed)
            const thirdFileRetailPriceColumn = df3.map(row => row[6]); // Column F from third file (adjust as needed)

            const plusSymbolColumn = df1.map(i => isNaN(i[6]) ? '' : '+'); // "+" symbol for first file rows
            const tenthColumnFile2 = df2.map(row => row[9]); // Column J from second file
            const thirdFileAvailabilityColumn = df3.map(row => row[1] ? isNaN(row[3]) ? '-' : '+' : ''); // Column E from third file (adjust as needed)

            // Combine columns from all three files
            const combinedColumnA = cycloneColumn.concat(mysteryColumn, thirdFileColumn);
            const combinedColumnB = fourthColumnFile1.concat(Array(mysteryColumn.length).fill(''));
            const combinedColumnC = fifthColumnFile1.concat(secondColumnFile2);
            const combinedColumnD = seventhColumnFile1.concat(fifthColumnFile2, thirdFilePriceColumn);
            const combinedColumnE = ninthColumnFile1.concat(seventhColumnFile2, thirdFileRetailPriceColumn);
            const combinedColumnPlus = plusSymbolColumn.concat(tenthColumnFile2, thirdFileAvailabilityColumn);

            // Create new data array
            const resultData = [];
            const maxLength = Math.max(
                combinedColumnA.length,
                combinedColumnB.length,
                combinedColumnC.length,
                combinedColumnD.length,
                combinedColumnE.length,
                combinedColumnPlus.length
            );

            // Add headers
            resultData.push([
                'Артикул', 'Бренд', 'Модель', 'Оптова USD', 'Роздрібна UAH', 'Наявність'
            ]);

            // Fill data rows
            for (let i = 0; i < maxLength; i++) {
                resultData.push([
                    i < combinedColumnA.length ? combinedColumnA[i] : '',
                    i < combinedColumnB.length ? combinedColumnB[i] : '',
                    i < combinedColumnC.length ? combinedColumnC[i] : '',
                    i < combinedColumnD.length ? combinedColumnD[i] : '',
                    i < combinedColumnE.length ? combinedColumnE[i] : '',
                    i < combinedColumnPlus.length ? combinedColumnPlus[i] : ''
                ]);
            }

            // Create new Excel workbook
            const wb = XLSX.utils.book_new();
            const ws = XLSX.utils.aoa_to_sheet(resultData);

            // Center alignment style
            const centerStyle = {
                alignment: {
                    horizontal: 'center',
                    vertical: 'center'
                }
            };

            // Header row style (dark orange)
            const headerStyle = {
                fill: {
                    patternType: "solid",
                    fgColor: {rgb: "FF8C00"} // Dark orange
                },
                font: {
                    bold: true,
                    sz: 14 // Larger font size
                }
            };

            // Single-cell row style (yellow)
            const singleCellStyle = {
                fill: {
                    patternType: "solid",
                    fgColor: {rgb: "FFFF00"} // Yellow
                },
                font: {
                    bold: true,
                    sz: 12 // Larger font size
                }
            };

            // Apply styles to all cells
            const range = XLSX.utils.decode_range(ws['!ref']);
            for (let R = range.s.r; R <= range.e.r; ++R) {
                for (let C = range.s.c; C <= range.e.c; ++C) {
                    const cell_address = {c: C, r: R};
                    const cell_ref = XLSX.utils.encode_cell(cell_address);

                    if (!ws[cell_ref]) continue;

                    // Base style (centering)
                    ws[cell_ref].s = {...centerStyle};

                    // Header row style
                    if (R === 0) {
                        ws[cell_ref].s = {...ws[cell_ref].s, ...headerStyle};
                    }

                    // Single-cell row style
                    if (R > 0 && C === 0) {
                        const row = resultData[R];
                        const hasSingleCell = row.filter(cell => cell).length === 1;
                        if (hasSingleCell) {
                            ws[cell_ref].s = {...ws[cell_ref].s, ...singleCellStyle};
                            // Apply yellow to entire row
                            for (let col = 0; col < row.length; col++) {
                                const col_ref = XLSX.utils.encode_cell({c: col, r: R});
                                if (!ws[col_ref]) continue;
                                ws[col_ref].s = {...ws[col_ref].s, ...singleCellStyle};
                            }
                        }
                    }
                }
            }

            // Array for merged cells
            const merges = [];

            // Analyze data for cell merging
            for (let r = 0; r < resultData.length; r++) {
                const row = resultData[r];
                let hasValue = false;
                let firstColWithValue = -1;
                let lastColWithValue = -1;

                for (let c = 0; c < row.length; c++) {
                    if (row[c] !== undefined && row[c] !== '' && row[c] !== null) {
                        if (firstColWithValue === -1) firstColWithValue = c;
                        lastColWithValue = c;
                        hasValue = true;
                    }
                }

                // Merge cells if only first cell has value
                if (hasValue && firstColWithValue === 0 && lastColWithValue === 0) {
                    merges.push({
                        s: {r: r, c: 0}, // Start cell
                        e: {r: r, c: row.length - 1} // End cell
                    });
                }
            }

            // Apply merges
            ws['!merges'] = merges;

            // Set column widths
            if (!ws['!cols']) ws['!cols'] = [];
            const colWidths = [
                {wch: 15}, // Column A (SKU)
                {wch: 20}, // Column B (Brand)
                {wch: 25}, // Column C (Model)
                {wch: 12}, // Column D (Wholesale USD)
                {wch: 15}, // Column E (Retail UAH)
                {wch: 10}  // Column F (Availability)
            ];
            ws['!cols'] = colWidths;

            // Add worksheet to workbook
            XLSX.utils.book_append_sheet(wb, ws, "Combined Data");

            // Generate filename with current date
            const currentDate = new Date().toISOString().split('T')[0];
            const fileName = `DSP-Sound_price_${currentDate}.xlsx`;

            // Save file
            XLSX.writeFile(wb, fileName);

        } catch (error) {
            alert(`Error: ${error.message}`);
        }
    });
});
