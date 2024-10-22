// Handle export button click event
     $("#export-btn").click(function() {
            // Create a new workbook and a new sheet
            var workbook = new ExcelJS.Workbook();
            var worksheet = workbook.addWorksheet('Styled Sheet');

        // Collect data from headings, paragraphs, lists, and tables
        $("h1, h2,h3,h6, p, ul li, table").each(function() {
            if ($(this).is("table")) {
                // If it's a table, process headers and rows
                var headers = [];
                var rowData = [];

                // Get headers
                $(this).find("th").each(function() {
                    headers.push($(this).text());
                });

                // Add headers if present
                if (headers.length > 0) {
                    var headerRow = worksheet.addRow(headers);

                    // Style the header row
                    headerRow.eachCell(function(cell, colNumber) {
                        cell.font = { bold: true, color: { argb: 'FFFFFF' } }; // White bold text
                        cell.fill = {
                            type: 'pattern',
                            pattern: 'solid',
                            fgColor: { argb: 'FF326099' }  // Blue background
                        };
                        cell.alignment = { horizontal: 'center', vertical: 'middle' };
                    });
                }

                // Get rows
                $(this).find("tbody tr").each(function() {
                    rowData = []; // Reset for each row
                    $(this).find("td").each(function() {
                        rowData.push($(this).text());
                    });

                    // Add row data and style it
                    var dataRow = worksheet.addRow(rowData);

                    // Style each row
                    dataRow.eachCell(function(cell, colNumber) {
                        cell.font = { color: { argb: 'FF000000' } }; // Black text
                        cell.border = {
                            top: { style: 'thin' },
                            left: { style: 'thin' },
                            bottom: { style: 'thin' },
                            right: { style: 'thin' }
                        };
                    });
                });

                worksheet.addRow([]); // Optional blank row between tables
				worksheet.addRow([]); 
				worksheet.addRow([]); 
            } else {
                // If it is not a table, just get the text
                var text = $(this).text().trim();
                var textRow = worksheet.addRow([text]); // Wrap in an array to create a row in Excel

                // Style for text (e.g., for paragraphs or headers)
                textRow.eachCell(function(cell) {
                    cell.font = {  bold: true,  size: 14, color: { argb: 'FF000000' } }; 
					
                });
            }
        });
		
		// Apply styles to the first row (header)
            worksheet.getRow(1).eachCell((cell) => {
                cell.font = { name: 'Arial', bold: true, color: { argb: 'FF326099' }, size: 20 };
                
            });
		 
		
		// Set the width for all columns (for example, set width to 20 for all)
        for (let i = 1; i <= worksheet.columnCount; i++) {
            worksheet.getColumn(i).width = 25; // Adjust the width as necessary
        }
		
		
            // Generate the Excel file
            workbook.xlsx.writeBuffer().then(function(data) {
                var blob = new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
                var url = window.URL.createObjectURL(blob);
                var anchor = document.createElement('a');
                anchor.href = url;
                anchor.download = "report" + new Date().toISOString().replace(/[\-\:\.]/g, "") + ".xlsx";
                document.body.appendChild(anchor);
                anchor.click();
                document.body.removeChild(anchor);
            });
        });
