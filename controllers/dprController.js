const xlsx = require('xlsx');
const XLSXStyle = require('xlsx-style');
const fs = require('fs');
const path = require('path');
const os = require('os');
const { v4: uuidv4 } = require('uuid');
const vesselsData = require('./vesselsHP.json');

const countMGOInAllSheets = (filePath) => {
    const workbook = xlsx.readFile(filePath);
    const sheetNames = workbook.SheetNames;
    const mgoOccurrences = [];

    sheetNames.forEach((sheetName) => {
        const worksheet = workbook.Sheets[sheetName];
        const sheetData = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

        sheetData.forEach((row, rowIndex) => {
            row.forEach((cell, columnIndex) => {
                if (cell === 'M G O (m3)') {
                    const nextRows = sheetData.slice(rowIndex + 1, rowIndex + 6);
                    mgoOccurrences.push({
                        sheetName,
                        rowIndex: rowIndex + 2,
                        columnIndex: columnIndex + 1,
                        rowData: nextRows,
                    });
                }
            });
        });
    });

    const rowDataArray = mgoOccurrences.map((occurrence) => occurrence.rowData[3]);
    const secondFilledElements = rowDataArray.map((occurrence) => {
        const filledElements = occurrence.filter((element) => element !== '');
        const secondElement = filledElements[1];
        return secondElement !== undefined ? parseFloat(secondElement) : null;
    });

    const sum = secondFilledElements.reduce((accumulator, currentValue) => accumulator + currentValue, 0);
    return sum ? sum : 0;
};

const countHoursPortAllSheets = (filePath) => {
    const workbook = xlsx.readFile(filePath);
    const sheetNames = workbook.SheetNames;
    const hours = [];

    sheetNames.forEach((sheetName) => {
        const worksheet = workbook.Sheets[sheetName];
        const sheetData = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

        sheetData.forEach((row, rowIndex) => {
            row.forEach((cell, columnIndex) => {
                if (cell === 'Equipment/s') {
                    const nextRows = sheetData.slice(rowIndex + 1, rowIndex + 6);
                    hours.push({
                        sheetName,
                        rowIndex: rowIndex + 2,
                        columnIndex: columnIndex + 1,
                        rowData: nextRows,
                    });
                }
            });
        });
    });

    const rowDataArray = hours.map((occurrence) => occurrence.rowData[0]);
    const secondFilledElements = rowDataArray.map((occurrence) => {
        const filledElements = occurrence.filter((element) => element !== '');
        const secondElement = filledElements[3];
        return secondElement !== undefined ? parseFloat(secondElement) : null;
    });

    const sum = secondFilledElements.reduce((accumulator, currentValue) => accumulator + currentValue, 0);
    return sum ? sum : 0;;
};
const countHoursSTBDAllSheets = (filePath) => {
    const workbook = xlsx.readFile(filePath);
    const sheetNames = workbook.SheetNames;
    const hours = [];

    sheetNames.forEach((sheetName) => {
        const worksheet = workbook.Sheets[sheetName];
        const sheetData = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

        sheetData.forEach((row, rowIndex) => {
            row.forEach((cell, columnIndex) => {
                if (cell === 'Equipment/s') {
                    const nextRows = sheetData.slice(rowIndex + 1, rowIndex + 6);
                    hours.push({
                        sheetName,
                        rowIndex: rowIndex + 2,
                        columnIndex: columnIndex + 1,
                        rowData: nextRows,
                    });
                }
            });
        });
    });

    const rowDataArray = hours.map((occurrence) => occurrence.rowData[1]);
    const secondFilledElements = rowDataArray.map((occurrence) => {
        const filledElements = occurrence.filter((element) => element !== '');
        const secondElement = filledElements[3];
        return secondElement !== undefined ? parseFloat(secondElement) : null;
    });

    const sum = secondFilledElements.reduce((accumulator, currentValue) => accumulator + currentValue, 0);
    return sum ? sum : 0;;
};

const countHoursCEAllSheets = (filePath) => {
    const workbook = xlsx.readFile(filePath);
    const sheetNames = workbook.SheetNames;
    const mgoOccurrences = [];

    sheetNames.forEach((sheetName) => {
        const worksheet = workbook.Sheets[sheetName];
        const sheetData = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

        sheetData.forEach((row, rowIndex) => {
            row.forEach((cell, columnIndex) => {
                if (cell === 'Equipment/s') {
                    const nextRows = sheetData.slice(rowIndex + 1, rowIndex + 6);
                    mgoOccurrences.push({
                        sheetName,
                        rowIndex: rowIndex + 2,
                        columnIndex: columnIndex + 1,
                        rowData: nextRows,
                    });
                }
            });
        });
    });

    const rowDataArray = mgoOccurrences.map((occurrence) => occurrence.rowData[2]);
    const secondFilledElements = rowDataArray.map((occurrence) => {
        const filledElements = occurrence.filter((element) => element !== '');
        const secondElement = filledElements[3];
        return secondElement !== undefined ? parseFloat(secondElement) : null;
    });

    const sum = secondFilledElements.reduce((accumulator, currentValue) => accumulator + currentValue, 0);
    return sum ? sum : 0;

};
const countHoursDG1AllSheets = (filePath) => {
    const workbook = xlsx.readFile(filePath);
    const sheetNames = workbook.SheetNames;
    const mgoOccurrences = [];

    sheetNames.forEach((sheetName) => {
        const worksheet = workbook.Sheets[sheetName];
        const sheetData = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

        sheetData.forEach((row, rowIndex) => {
            row.forEach((cell, columnIndex) => {
                if (cell === 'Equipment/s') {
                    const nextRows = sheetData.slice(rowIndex + 1, rowIndex + 6);
                    mgoOccurrences.push({
                        sheetName,
                        rowIndex: rowIndex + 2,
                        columnIndex: columnIndex + 1,
                        rowData: nextRows,
                    });
                }
            });
        });
    });

    const rowDataArray = mgoOccurrences.map((occurrence) => occurrence.rowData[3]);
    const secondFilledElements = rowDataArray.map((occurrence) => {
        const filledElements = occurrence.filter((element) => element !== '');
        const secondElement = filledElements[3];
        return secondElement !== undefined ? parseFloat(secondElement) : null;
    });

    const sum = secondFilledElements.reduce((accumulator, currentValue) => accumulator + currentValue, 0);
    return sum ? sum : 0;

};
const countHoursDG2AllSheets = (filePath) => {
    const workbook = xlsx.readFile(filePath);
    const sheetNames = workbook.SheetNames;
    const mgoOccurrences = [];

    sheetNames.forEach((sheetName) => {
        const worksheet = workbook.Sheets[sheetName];
        const sheetData = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

        sheetData.forEach((row, rowIndex) => {
            row.forEach((cell, columnIndex) => {
                if (cell === 'Equipment/s') {
                    const nextRows = sheetData.slice(rowIndex + 1, rowIndex + 6);
                    mgoOccurrences.push({
                        sheetName,
                        rowIndex: rowIndex + 2,
                        columnIndex: columnIndex + 1,
                        rowData: nextRows,
                    });
                }
            });
        });
    });

    const rowDataArray = mgoOccurrences.map((occurrence) => occurrence.rowData[4]);
    const secondFilledElements = rowDataArray.map((occurrence) => {
        const filledElements = occurrence.filter((element) => element !== '');
        const secondElement = filledElements[3];
        return secondElement !== undefined ? parseFloat(secondElement) : null;
    });

    const sum = secondFilledElements.reduce((accumulator, currentValue) => accumulator + currentValue, 0);
    return sum ? sum : 0;

};

const countHoursDG3AllSheets = (filePath) => {
    const workbook = xlsx.readFile(filePath);
    const sheetNames = workbook.SheetNames;
    const mgoOccurrences = [];

    sheetNames.forEach((sheetName) => {
        const worksheet = workbook.Sheets[sheetName];
        const sheetData = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

        sheetData.forEach((row, rowIndex) => {
            row.forEach((cell, columnIndex) => {
                if (cell === 'Equipment/s') {
                    const nextRows = sheetData.slice(rowIndex + 1, rowIndex + 8);
                    mgoOccurrences.push({
                        sheetName,
                        rowIndex: rowIndex + 2,
                        columnIndex: columnIndex + 1,
                        rowData: nextRows,
                    });
                }
            });
        });
    });

    const rowDataArray = mgoOccurrences.map((occurrence) => occurrence.rowData[5]);
    const secondFilledElements = rowDataArray.map((occurrence) => {
        const filledElements = occurrence.filter((element) => element !== '');
        const secondElement = filledElements[3];
        return secondElement !== undefined ? parseFloat(secondElement) : null;
    });

    const sum = secondFilledElements.reduce((accumulator, currentValue) => accumulator + currentValue, 0);
    return sum ? sum : 0;

};
const countHoursDG4AllSheets = (filePath) => {
    const workbook = xlsx.readFile(filePath);
    const sheetNames = workbook.SheetNames;
    const mgoOccurrences = [];

    sheetNames.forEach((sheetName) => {
        const worksheet = workbook.Sheets[sheetName];
        const sheetData = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

        sheetData.forEach((row, rowIndex) => {
            row.forEach((cell, columnIndex) => {
                if (cell === 'Equipment/s') {
                    const nextRows = sheetData.slice(rowIndex + 1, rowIndex + 8);
                    mgoOccurrences.push({
                        sheetName,
                        rowIndex: rowIndex + 2,
                        columnIndex: columnIndex + 1,
                        rowData: nextRows,
                    });
                }
            });
        });
    });

    const rowDataArray = mgoOccurrences.map((occurrence) => occurrence.rowData[6]);
    const secondFilledElements = rowDataArray.map((occurrence) => {
        const filledElements = occurrence.filter((element) => element !== '');
        const secondElement = filledElements[3];
        return secondElement !== undefined ? parseFloat(secondElement) : null;
    });

    const sum = secondFilledElements.reduce((accumulator, currentValue) => accumulator + currentValue, 0);
    return sum ? sum : 0;

};
const countHoursDG5AllSheets = (filePath) => {
    const workbook = xlsx.readFile(filePath);
    const sheetNames = workbook.SheetNames;
    const mgoOccurrences = [];

    sheetNames.forEach((sheetName) => {
        const worksheet = workbook.Sheets[sheetName];
        const sheetData = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

        sheetData.forEach((row, rowIndex) => {
            row.forEach((cell, columnIndex) => {
                if (cell === 'Equipment/s') {
                    const nextRows = sheetData.slice(rowIndex + 1, rowIndex + 9);
                    mgoOccurrences.push({
                        sheetName,
                        rowIndex: rowIndex + 2,
                        columnIndex: columnIndex + 1,
                        rowData: nextRows,
                    });
                }
            });
        });
    });

    const rowDataArray = mgoOccurrences.map((occurrence) => occurrence.rowData[7]);
    const secondFilledElements = rowDataArray.map((occurrence) => {
        const filledElements = occurrence.filter((element) => element !== '');
        const secondElement = filledElements[3];
        return secondElement !== undefined ? parseFloat(secondElement) : null;
    });

    const sum = secondFilledElements.reduce((accumulator, currentValue) => accumulator + currentValue, 0);
    return sum ? sum : 0;

};
const countMilageAllSheets = (filePath) => {
    const workbook = xlsx.readFile(filePath);
    const sheetNames = workbook.SheetNames;
    const mgoOccurrences = [];

    sheetNames.forEach((sheetName) => {
        const worksheet = workbook.Sheets[sheetName];
        const sheetData = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

        sheetData.forEach((row, rowIndex) => {
            row.forEach((cell, columnIndex) => {
                if (cell === 'Stand By:') {
                    const nextRows = sheetData.slice(rowIndex + 1, rowIndex + 9);
                    mgoOccurrences.push({
                        sheetName,
                        rowIndex: rowIndex + 2,
                        columnIndex: columnIndex + 1,
                        rowData: nextRows,
                    });
                }
            });
        });
    });

    const rowDataArray = mgoOccurrences.map((occurrence) => occurrence.rowData[0]);
    const secondFilledElements = rowDataArray.map((occurrence) => {
        const filledElements = occurrence.filter((element) => element !== '');
        const secondElement = filledElements[2];

        return secondElement !== undefined ? secondElement : null;
    });

    const modifiedResult = secondFilledElements.map((str) => {
        if (typeof str === 'string') {
            const numberOnly = str.replace(/[^0-9]/g, '');
            const parsedNumber = parseInt(numberOnly, 10); // Parse the string as an integer with base 10

            return isNaN(parsedNumber) ? 0 : parsedNumber;
        }
        return str; // Keep existing numbers intact
    });

    const sum = modifiedResult.reduce((accumulator, currentValue) => parseFloat(accumulator) + parseFloat(currentValue), 0);

    return sum;
};




const readVesselsFileHP = (vesselName) => {
    // Remove empty elements from the firstRow array
    var hp = 0;
    var matchFound = false;

    // Access the vessel name and horsepower values directly from vesselsData
    vesselsData.vessels.forEach((vessel) => {
        const name = vessel.name;
        const horsepower = vessel.horsepower;

        if (vesselName.toUpperCase().includes(name.toUpperCase())) {
            hp = horsepower;
            matchFound = true;
            // You can perform additional operations if needed
            // within this if block
        }

        // Do something with the name and horsepower values
    });

    if (matchFound) {
        return hp;
    } else {
        return "no HP found";
    }
};
const readVesselsFileAUX = (vesselName) => {
    // Remove empty elements from the firstRow array
    var aux = 0;
    var matchFound = false;

    // Access the vessel name and horsepower values directly from vesselsData
    vesselsData.vessels.forEach((vessel) => {
        const name = vessel.name;
        const auxFile = vessel.consumptionAUX;

        if (vesselName.toUpperCase().includes(name.toUpperCase())) {
            aux = auxFile;
            matchFound = true;
            // You can perform additional operations if needed
            // within this if block
        }

        // Do something with the name and horsepower values
    });

    if (matchFound) {
        return aux;
    } else {
        return "no HP found";
    }
};



// Handle the uploaded files
const dprController = (req, res) => {
    const files = req.files;
    const fileDataArray = [];
    var serialNumber = 1;
    fileDataArray.unshift(['S/N', 'Vessel Name', 'HP', 'Monthly CONSM.Cu.M',
        'Monthly ME RH', 'ME CON./h', 'ME CON./D Cu.M', 'DG1 RH', 'DG2 RH', 'DG3 RH', 'DG4 RH',
        'DG5 RH', 'AUX R/H Total', 'NO AUX R/D', 'Aux CON./h', 'AUX CON./D Cu.M',
        'Estimated vessel CON AVG. Vessel Daily Consumption Sailing Cu.M', 'Ratio HP-L', 'Total DIST.']);

    // Process each uploaded file
    files.forEach((file) => {
        const totalConsumption = countMGOInAllSheets(file.path);
        const totalhoursPort = countHoursPortAllSheets(file.path);
        const totalhoursSTBD = countHoursSTBDAllSheets(file.path);
        const totalHour = totalhoursPort + totalhoursSTBD;
        const totalHours = totalHour / 2;
        const totalConsumptionPerHour = totalConsumption * 1000 / totalHours
        const totalHoursDG1 = countHoursDG1AllSheets(file.path);
        const totalHoursDG2 = countHoursDG2AllSheets(file.path);
        const totalHoursDG3 = countHoursDG3AllSheets(file.path);
        const totalHoursDG4 = countHoursDG4AllSheets(file.path);
        const totalHoursDG5 = countHoursDG5AllSheets(file.path);
        const totalDGHours = totalHoursDG1 + totalHoursDG2 + totalHoursDG3 + totalHoursDG4 + totalHoursDG5;
        const totalMile = countMilageAllSheets(file.path);
        // Read the Excel file
        const workbook = xlsx.readFile(file.path);

        // Get all sheet names
        var sheetNames = workbook.SheetNames;
        sheetNames = sheetNames.map((sheetName) => {
            if (sheetName.includes('Chart')) {
                return null; // or any other value you want to replace 'Chart' with
            }
            return sheetName;
        }).filter(Boolean);

        let monthDay = sheetNames.length;
        let daySheet = sheetNames[sheetNames.length - 1];
        if (sheetNames.includes('Monthly Bunkering & Consumption')) {
            let index = sheetNames.indexOf('Monthly Bunkering & Consumption');
            daySheet = sheetNames[index - 1];
            monthDay = index
            monthSheet = sheetNames[index];
        }
        const worksheet = workbook.Sheets[daySheet];
        var VesselRow;
        const sheetData = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

        sheetData.forEach((row, rowIndex) => {

            row.forEach((cell, colIndex) => {
                if (cell === 'Vessel Name:') {
                    VesselRow = row ? row.filter((element) => element !== '') : [];
                }

                if (fileDataArray.length === 1) {
                    fileDataArray.unshift(['', '', '', '', VesselRow[2], VesselRow[3]]);

                }



            });
        });
        const vesselName = VesselRow[1]
        const vesselHP = readVesselsFileHP(vesselName);
        const vesselAUXConsumption = readVesselsFileAUX(vesselName)
        const meCON = parseFloat(((totalConsumption * 1000 - vesselAUXConsumption * totalDGHours) / totalHours).toFixed(1));
        const auxCON = parseFloat((((totalDGHours.toFixed(1) * vesselAUXConsumption)) / 24 / monthDay).toFixed(1))
        const meCONperD = meCON * 24 / 1000
        const auxCONperD = auxCON * 24 / 1000
        fileDataArray.push([serialNumber, vesselName, vesselHP, parseFloat(totalConsumption.toFixed(1)), parseFloat(totalHours.toFixed(1)),
            meCON, meCONperD, parseFloat(totalHoursDG1.toFixed(1)), parseFloat(totalHoursDG2.toFixed(1)), parseFloat(totalHoursDG3.toFixed(1)),
            parseFloat(totalHoursDG4.toFixed(1)), parseFloat(totalHoursDG5.toFixed(1)), parseFloat(totalDGHours.toFixed(1)), parseFloat((totalDGHours / 24 / monthDay).toFixed(1)),
            parseFloat(auxCON).toFixed(1), parseFloat(auxCONperD).toFixed(1), parseFloat(meCONperD + auxCONperD).toFixed(1),
            parseFloat((meCONperD + auxCONperD) * 1000 / vesselHP).toFixed(1), parseFloat(totalMile.toFixed(1))])
        serialNumber++;


    });


    // Create a new workbook
    const newWorkbook = xlsx.utils.book_new();

    // Create a new worksheet
    const newWorksheet = xlsx.utils.aoa_to_sheet(fileDataArray);

    // ...


    const style2 = {
        alignment: {
            horizontal: 'center',
            vertical: 'center'
        },
        border: {
            top: { style: 'thin', color: { rgb: '00000000' } },
            bottom: { style: 'thin', color: { rgb: '00000000' } },
            left: { style: 'thin', color: { rgb: '00000000' } },
            right: { style: 'thin', color: { rgb: '00000000' } },
        },
    };
    const style = {
        ...style2,
        fill: {
            fgColor: {
                rgb: 'D6DCE4'
            }, // Blue color code
        },
    };
    const style3 = {
        ...style2,
        fill: {
            fgColor: {
                rgb: 'ffff79'
            }, // Blue color code
        },
    };
    const style4 = {
        ...style2,
        fill: {
            fgColor: {
                rgb: 'EEECE1'
            }, // Blue color code
        },
    };

    // Apply the style to each cell in the worksheet
    const range = xlsx.utils.decode_range(newWorksheet['!ref']);
    // Apply the style to each cell in the first and second columns of the worksheet

    for (let rowNum = range.s.r + 1; rowNum <= range.e.r; rowNum++) {
        let isEmptyRow = true; // Flag to track if the row is empty

        for (let colNum = range.s.c; colNum <= range.e.c; colNum++) {
            const cellAddress = xlsx.utils.encode_cell({ r: rowNum, c: colNum });
            const cell = newWorksheet[cellAddress] || { s: {} }; // Use an empty object if cell is undefined

            // Create a new style object with the desired properties
            const cellStyle = {
                ...style2, // Existing style properties

            };

            cell.s = cellStyle;
            if (colNum < 11) {
                if (cell.v !== undefined && cell.v !== '') {
                    isEmptyRow = false; // Row is not empty if at least one cell has a value
                }
                if (cell.v === undefined) {
                    cell.v = ""; // Set the value to a blank string if cell value is undefined
                }
                newWorksheet[cellAddress] = cell; // Update the cell in the worksheet
            }
        }

        // Remove the style from the row if it is empty

        const cellAddress = xlsx.utils.encode_cell({ r: rowNum, c: 0 }); // First column
        const cell = newWorksheet[cellAddress] || { s: {} }; // Use an empty object if cell is undefined

        cell.s = style;
        newWorksheet[cellAddress] = cell;
        if (isEmptyRow) {
            for (let colNum = range.s.c; colNum <= range.e.c; colNum++) {
                const cellAddress = xlsx.utils.encode_cell({ r: rowNum, c: colNum });
                const cell = newWorksheet[cellAddress] || { s: {} }; // Use an empty object if cell is undefined
                cell.s = {};
                cell.v = ''; // Remove the style
                newWorksheet[cellAddress] = cell; // Update the cell in the worksheet
            }
        }
    }

    // Set column widths
    const columnWidths = [
        { wpx: 30 },  // S/N
        { wpx: 130 }, // Vessel Name
        { wpx: 50 },  // HP
        { wpx: 150 }, // Monthly CONSM.Cu.M
        { wpx: 110 }, // Monthly ME RH
        { wpx: 90 }, // me CON./h
        { wpx: 90 }, // me CON./D
        { wpx: 50 }, // DG1 RH
        { wpx: 50 }, // DG2 RH
        { wpx: 50 }, // DG3 RH
        { wpx: 50 }, // DG4 RH
        { wpx: 50 }, // DG5 RH
        { wpx: 100 }, // AUX R/H Total
        { wpx: 100 }, // Number of AUX running per day
        { wpx: 90 }, // aux CON./h
        { wpx: 90 }, // aux CON./D
        { wpx: 90 }, //ratio HP-L
        { wpx: 80 }, // Total DIST.
    ];
    newWorksheet['!cols'] = columnWidths;

    const firstRow = 1; // Assuming the first row is at index 1
    const firstRowRange = xlsx.utils.decode_range(newWorksheet['!ref']);
    for (let col = firstRowRange.s.c + 2; col <= firstRowRange.e.c; col++) {
        const cellAddress = xlsx.utils.encode_cell({ r: firstRow, c: col });
        const cell = newWorksheet[cellAddress];
        cell.s = style; // Blue color code
    }

    //style dg columns
    fileDataArray.forEach((rowData, rowIndex) => {
        if (rowIndex > 0) { // Skip the first row
            for (let columnIndex = 7; columnIndex <= 11; columnIndex++) {
                const cellAddress = { c: columnIndex, r: rowIndex };
                const cell = newWorksheet[xlsx.utils.encode_cell(cellAddress)];
                cell.s = style4;
            }

        }

    });
    //style Estimated vessel CON AVG. Vessel Daily Consumption Sailing Cu.M
    fileDataArray.forEach((rowData, rowIndex) => {
        if (rowIndex > 0) { // Skip the first row
            let columnIndex = 16
            const cellAddress = { c: columnIndex, r: rowIndex };
            const cell = newWorksheet[xlsx.utils.encode_cell(cellAddress)];
            cell.s = style3;
        }
    });
    //style Estimated vessel CON Aux CON./Day Cu.M
    fileDataArray.forEach((rowData, rowIndex) => {
        if (rowIndex > 0) { // Skip the first row
            let columnIndex = 15
            const cellAddress = { c: columnIndex, r: rowIndex };
            const cell = newWorksheet[xlsx.utils.encode_cell(cellAddress)];
            cell.s = style;
        }
    });
    //style Estimated vessel CON ME CON./Day Cu.M
    fileDataArray.forEach((rowData, rowIndex) => {
        if (rowIndex > 0) { // Skip the first row
            let columnIndex = 6
            const cellAddress = { c: columnIndex, r: rowIndex };
            const cell = newWorksheet[xlsx.utils.encode_cell(cellAddress)];
            cell.s = style;
        }
    });
    fileDataArray.forEach((rowData, rowIndex) => {
        rowData.forEach((cellData, columnIndex) => {
            if (cellData === 'Date:') {
                const dateCellAddress = { c: columnIndex, r: rowIndex };
                const dateCell = newWorksheet[xlsx.utils.encode_cell(dateCellAddress)];
                dateCell.s = style;

                const nextCellAddress = { c: columnIndex + 1, r: rowIndex };
                const nextCell = newWorksheet[xlsx.utils.encode_cell(nextCellAddress)];
                nextCell.s = style2;
            }
        });
    });
    fileDataArray.forEach((rowData, rowIndex) => {
        rowData.forEach((cellData, columnIndex) => {
            if (cellData === 'Vessel Name') {
                const dateCellAddress = { c: columnIndex, r: rowIndex };
                const dateCell = newWorksheet[xlsx.utils.encode_cell(dateCellAddress)];
                dateCell.s = style;


            }
        });
    });
    // Add the worksheet to the workbook
    xlsx.utils.book_append_sheet(newWorkbook, newWorksheet, 'Data');

    // Generate Excel file buffer
    const excelBuffer = XLSXStyle.write(newWorkbook, { type: 'buffer', bookType: 'xlsx' });

    // Create the 'downloads' directory if it doesn't exist
    const generateUniqueFileName = () => {
        const uniqueId = uuidv4();
        return `data-${uniqueId}.xlsx`;
    };
    const tempDir = os.tmpdir();

    // Save the buffer to a file
    const filePath = path.join(tempDir, generateUniqueFileName());
    fs.writeFileSync(filePath, excelBuffer);

    // Send the file for download
    res.download(filePath, 'data.xlsx', (err) => {
        if (err) {
            console.error('Error sending file:', err);
        }

        // Remove the temporary file
        fs.unlinkSync(filePath);
    });
};

module.exports = dprController;
