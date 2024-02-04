document.addEventListener('DOMContentLoaded', function () {
    // Get reference to file input and column select
    var fileInput = document.getElementById('fileInput');
    var resultColumnSelect = document.getElementById('resultColumnSelect');
    var columnSelect = document.getElementById('columnSelect');
    var preferenceColumnsSelect = document.getElementById('preferenceColumnsSelect');

    var studentsPerSubjectInput = document.getElementById('studentsPerSubject');

    var sortedData;
    var sortedData1;
    // Add event listener for file input change
    fileInput.addEventListener('change', function (e) {
        var file = e.target.files[0];

        // Read the Excel file
        readExcel(file, function (data) {
            // Display data in the table
            displayData(data);

            // Populate sort column dropdown
            populateColumnDropdown(columnSelect, data[0]);

            // Populate preference columns dropdown
            populateColumnDropdown(preferenceColumnsSelect, data[0], true);

            // Populate assigned subject column dropdown
            populateColumnDropdown(resultColumnSelect, data[0]);
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

            callback(jsonData);
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

    function assignMinValueColumn(data, resultColumnIndex, selectedColumns, subjectAvailability) {
        data.forEach(function (row) {

            var currentPreference = 1;
            var selectedSubject = null;

            // Try to find a subject with available slots based on preferences
            while (currentPreference <= selectedColumns.length) {
                var minColumnIndex = findMinValueColumn(row, selectedColumns, currentPreference);
                if (minColumnIndex !== null) {
                    selectedSubject = data[0][minColumnIndex];
                    if (subjectAvailability[selectedSubject] > 0) {
                        row[resultColumnIndex] = selectedSubject;
                        subjectAvailability[selectedSubject]--;
                        break; // Break the loop if the subject is assigned
                    } else {
                        currentPreference++; // Move to the next preference if the current one is full
                    }
                } else {
                    break; // Break the loop if no more preferences are available
                }
            }

            // Assign 'No Preference' if all preferences are full
            if (currentPreference > selectedColumns.length) {
                row[resultColumnIndex] = 'No Preference (All Full)';
            }
        });
    }

    // Function to find the column index with the minimum value based on preferences
    function findMinValueColumn(row, selectedColumns, preference) {
        var minColumnIndex = null;


        selectedColumns.forEach(function (colIndex) {
            if (row[colIndex] === preference.toString()) {

                minColumnIndex = colIndex;
                // console.log(row[colIndex], colIndex, preference.toString());
                return minColumnIndex;
            }
        });
        // console.log(minColumnIndex);
        return minColumnIndex;
    }




    // Function to sort and download modified Excel file
    window.sortAndDownload = function () {
        var columnIndex = parseInt(columnSelect.value);

        if (!isNaN(columnIndex)) {
            // Sort the data by the selected column index
            sortedData1 = sortData(columnIndex);
            // Update the displayed data in the table
            displayData(sortedData1);
            document.getElementById("myButton2").disabled = true;


        } else {
            alert('Please select a column for sorting.');
        }
    };


    window.allocatedSubjects = function () {
        var columnIndex = parseInt(columnSelect.value);

        var resultColumnIndex = parseInt(resultColumnSelect.value);

        var maxStudent = parseInt(studentsPerSubjectInput.value);

        if (!isNaN(columnIndex)) {
            // Sort the data by the selected column index
            sortedData = sortData(columnIndex);

            // Get selected columns from user input
            var selectedColumns = Array.from(preferenceColumnsSelect.selectedOptions)
                .map(option => parseInt(option.value));

            var subjectAvailability = {};
            // Initialize subject availability for each subject
            selectedColumns.forEach(function (colIndex) {
                var subject = sortedData[0][colIndex];
                subjectAvailability[subject] = maxStudent;
                // console.log(colIndex);
            });

            // Assign the column with the minimum value to the result column for each row
            assignMinValueColumn(sortedData, resultColumnIndex, selectedColumns, subjectAvailability);

            displayData(sortedData);
            document.getElementById("myButton").disabled = true;

        } else {
            alert('Please select a column for sorting.');
        }
    };

    window.downloadData = function () {
        // Create a new workbook with the sorted data
        var newWorkbook = XLSX.utils.book_new();
        var newSheet = XLSX.utils.json_to_sheet(sortedData);
        XLSX.utils.book_append_sheet(newWorkbook, newSheet, 'Sheet1');

        // Save the workbook to a file
        XLSX.writeFile(newWorkbook, 'output.xlsx');
    };

    window.downloadData2 = function () {
        // Create a new workbook with the sorted data
        var newWorkbook = XLSX.utils.book_new();
        var newSheet = XLSX.utils.json_to_sheet(sortedData1);
        XLSX.utils.book_append_sheet(newWorkbook, newSheet, 'Sheet1');

        // Save the workbook to a file
        XLSX.writeFile(newWorkbook, 'output_sorted.xlsx');
    };

    // Function to sort data by a specific column index
    function sortData(columnIndex) {
        var table = document.getElementById('dataTable');
        var tbody = table.querySelector('tbody');

        var rows = Array.from(tbody.getElementsByTagName('tr'));
        rows.shift(); // Remove header row

        // Sort rows based on the selected column index
        rows.sort(function (a, b) {
            var cellA = a.cells[columnIndex].textContent;
            var cellB = b.cells[columnIndex].textContent;
            return cellB.localeCompare(cellA, undefined, { numeric: true, sensitivity: 'base' });
        });

        // Create a new array with the sorted data
        var sortedData = rows.map(function (row) {
            return Array.from(row.cells).map(function (cell) {
                return cell.textContent;
            });
        });

        // Add the header row back to the sorted data
        sortedData.unshift(Array.from(table.querySelector('thead tr').cells).map(function (cell) {
            return cell.textContent;
        }));

        return sortedData;
    }
});
