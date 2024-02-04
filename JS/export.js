document.addEventListener('DOMContentLoaded', function () {
    // Get reference to file input and column select
    var fileInput = document.getElementById('fileInput');
    var columnSelect = document.getElementById('columnSelect');
    var shuffleColumnsSelect = document.getElementById('shuffleColumnsSelect');


    var excelData;

    // Add event listener for file input change
    fileInput.addEventListener('change', function (e) {
        var file = e.target.files[0];

        // Read the Excel file
        readExcel(file, function (data) {
            // Display data in the table
            excelData = data;
            displayData(data);

            // Populate sort column dropdown
            populateColumnDropdown(columnSelect, data[0]);

            // Populate preference columns dropdown
            populateColumnDropdown(shuffleColumnsSelect, data[0], true);

        });
    });

    // Function to read Excel file

    function readExcel(file, callback) {
        var reader = new FileReader();

        reader.onload = function (e) {
            var data = new Uint8Array(e.target.result);
            var workbook = XLSX.read(data, { type: 'array' });

            // Get the first sheet
            var sheet = workbook.Sheets[workbook.SheetNames[0]];

            // Convert sheet data to array of objects
            var jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

            // Filter out empty rows
            var nonEmptyRows = jsonData.filter(row => row.some(cell => cell !== undefined && cell !== null && cell !== ''));

            callback(nonEmptyRows);
        };

        reader.readAsArrayBuffer(file);
    }


    // Function to display data in the table
    function displayData(data) {
        var table = document.getElementById('dataTable');
        table.innerHTML = '';

        // Create table header
        var thead = document.createElement('thead');
        var headerRow = document.createElement('tr');

        data[0].forEach(function (cell) {
            var th = document.createElement('th');
            th.textContent = cell;
            headerRow.appendChild(th);

        });

        thead.appendChild(headerRow);
        table.appendChild(thead);

        // Create table body
        var tbody = document.createElement('tbody');

        for (var i = 1; i < data.length; i++) {
            var row = document.createElement('tr');

            data[i].forEach(function (cell) {
                var td = document.createElement('td');
                td.textContent = cell;
                row.appendChild(td);
            });

            tbody.appendChild(row);
        }

        table.appendChild(tbody);
    }


    // Function to populate column dropdown
    function populateColumnDropdown(select, columns, multiple = false) {
        select.innerHTML = '';

        columns.forEach(function (column, index) {
            var option = document.createElement('option');
            option.value = index;
            option.textContent = column;
            select.appendChild(option);
        });

        select.multiple = multiple;
    }


    //   // Populate the column select dropdown
    //   var columnSelect = document.getElementById('columnSelect');
    //   for (var i = 0; i < excelData[0].length; i++) {
    //     var option = document.createElement('option');
    //     option.value = i;
    //     option.text = excelData[0][i];
    //     columnSelect.add(option);
    //   }

    // Function to add an empty column to the Excel sheet
    function addEmptyColumnToSheet(data, columnName) {
        // Add the new column name to the header row
        data[0].push(columnName);

        // Add an empty value to each row in the new column
        for (var i = 1; i < data.length; i++) {
            data[i].push('');
        }

        return data;
    }

    // Function to handle the button click
    window.addNewColumn = function () {
        var columnName = document.getElementById('columnName').value;
        if (columnName.trim() === '') {
            alert('Please enter a column name.');
            return;
        }

        // Add the new column to the Excel data
        excelData = addEmptyColumnToSheet(excelData, columnName);

        displayData(excelData);

         // Populate sort column dropdown
        populateColumnDropdown(columnSelect, data[0]);
    }


    // Function to randomize a column within a specified range
    window.randomizeColumn = function () {
        var columnIndex = parseInt(document.getElementById('columnSelect').value);
        var minValue = parseInt(document.getElementById('minValue').value);
        var maxValue = parseInt(document.getElementById('maxValue').value);

        // Set default min and max values to the column min and max
        var columnData = excelData.slice(1).map(row => convertToNumber(row[columnIndex]));

        if (isNaN(minValue) || isNaN(maxValue)) {
            // Set default values for the input fields
            minValue = Math.min(...columnData);
            maxValue = Math.max(...columnData);
        }

        // Randomize the selected column within the specified range
        for (var i = 1; i < excelData.length; i++) {
            var randomValue = Math.floor(Math.random() * (maxValue - minValue + 1)) + minValue;
            excelData[i][columnIndex] = randomValue;
        }

        // Display the updated data
        displayData(excelData);
    }

    // Function to convert string values to numbers
    function convertToNumber(value) {
        if (typeof value === 'string') {
            // Remove non-numeric characters and map to 0, 1, 2, ...
            value = value.replace(/[^0-9]/g, '');
            if (value === '') {
                return 0; // Default to 0 if the string was not numeric
            }
            return parseInt(value, 10);
        }
        return value; // If it's not a string, leave it as is
    }

    // Fisher-Yates shuffle algorithm
    function shuffleArray(array) {
        for (let i = array.length - 1; i > 0; i--) {
            const j = (Math.floor(Math.random() * (i + 1))) % array.length;
            [array[i], array[j]] = [array[j], array[i]];
        }
        return array;
    }

    // Function to shuffle selected columns in an Excel sheet
    function shuffleColumns(data, selectedColumns) {

        // Extract the values from the selected columns
        const selectedColumnValues = selectedColumns.map(columnIndex => {
            return {
                index: columnIndex,
                values: data.slice(1).map(row => row[columnIndex])
            };
        });

        // Shuffle the values within each selected column independently
        selectedColumnValues.forEach(column => {
            column.values = shuffleArray(column.values);
        });

        // Update the original data with the shuffled values
        data.slice(1).forEach((row, rowIndex) => {
            selectedColumnValues.forEach(column => {

                row[column.index] = column.values[rowIndex];
            });

        });

    }



    window.shuffledSubjects = function () {


        // Replace with the indices of the columns you want to shuffle
        const selectedColumnsToShuffle = Array.from(shuffleColumnsSelect.selectedOptions)
            .map(option => parseInt(option.value));

        if (!(selectedColumnsToShuffle.empty)) {


            // Replace with the indices of the columns you want to shuffle
            shuffleColumns(excelData, selectedColumnsToShuffle);

            // Display the shuffled data
            // console.log(shuffledData);

            displayData(excelData);


        } else {
            alert('Please select a column for sorting.');
        }
    };

    window.downloadData = function () {
        // Create a new workbook with the sorted data
        var newWorkbook = XLSX.utils.book_new();
        var newSheet = XLSX.utils.json_to_sheet(excelData);
        XLSX.utils.book_append_sheet(newWorkbook, newSheet, 'Sheet1');

        // Save the workbook to a file
        XLSX.writeFile(newWorkbook, 'output.xlsx');

    };




});
