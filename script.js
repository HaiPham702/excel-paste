document.addEventListener('DOMContentLoaded', function() {
    const pasteArea = document.getElementById('pasteArea');
    const dataTable = document.getElementById('dataTable');
    
    // Focus the paste area when clicking on it
    pasteArea.addEventListener('click', function() {
        pasteArea.focus();
    });
    
    // Handle paste event
    pasteArea.addEventListener('paste', function(e) {
        e.preventDefault();
        
        // Get clipboard data
        const clipboardData = e.clipboardData || window.clipboardData;
        
        // Try to get HTML data first
        let pastedData = clipboardData.getData('text/html');
        debugger
        
        // If no HTML data is available, fall back to plain text
        if (!pastedData || !pastedData.trim()) {
            pastedData = clipboardData.getData('text');
            
            if (!pastedData.trim()) {
                return;
            }
            
            // Process as plain text
            processExcelTextData(pastedData);
        } else {
            // Process as HTML
            processExcelHtmlData(pastedData);
        }
        
        // Clear placeholder text
        pasteArea.innerHTML = '';
    });
    
    function processExcelTextData(data) {
        // Split the pasted text into rows
        const rows = data.trim().split(/[\r\n]+/);
        
        // Check if we have any data
        if (rows.length === 0) {
            return;
        }
        
        // Clear existing table data
        dataTable.innerHTML = '';
        
        // Create table header row
        const headerRow = rows[0].split('\t');
        const thead = document.createElement('thead');
        const headerTr = document.createElement('tr');
        
        headerRow.forEach(headerText => {
            const th = document.createElement('th');
            th.textContent = headerText;
            headerTr.appendChild(th);
        });
        
        thead.appendChild(headerTr);
        dataTable.appendChild(thead);
        
        // Create table body
        const tbody = document.createElement('tbody');
        
        // Start from index 1 to skip the header row
        for (let i = 1; i < rows.length; i++) {
            const row = rows[i].split('\t');
            
            // Skip empty rows
            if (row.join('').trim() === '') {
                continue;
            }
            
            const tr = document.createElement('tr');
            
            row.forEach(cellText => {
                const td = document.createElement('td');
                td.textContent = cellText;
                tr.appendChild(td);
            });
            
            tbody.appendChild(tr);
        }
        
        dataTable.appendChild(tbody);
        
        // Update the UI to show the data has been processed
        pasteArea.innerHTML = '<p>Data processed! Click here to paste new data.</p>';
        pasteArea.classList.add('success');
    }
    
    function processExcelHtmlData(htmlData) {
        debugger
        // Create a temporary div to parse the HTML
        const tempDiv = document.createElement('div');
        tempDiv.innerHTML = htmlData.trim();
        
        // Find the table in the pasted HTML
        const excelTable = tempDiv.querySelector('table');
        
        if (!excelTable) {
            // If no table found, try processing as text
            const textData = tempDiv.textContent;
            if (textData.trim()) {
                processExcelTextData(textData);
            }
            return;
        }
        
        // Clear existing table data
        dataTable.innerHTML = '';
        
        // Create a new table structure
        const thead = document.createElement('thead');
        const tbody = document.createElement('tbody');
        
        // Process the rows
        const rows = excelTable.querySelectorAll('tr');
        
        if (rows.length > 0) {
            // First row is header
            const headerRow = rows[0];
            const headerCells = headerRow.querySelectorAll('td, th');
            
            const headerTr = document.createElement('tr');
            
            headerCells.forEach(cell => {
                const th = document.createElement('th');
                th.innerHTML = cell.innerHTML; // Preserve HTML formatting
                
                // Preserve merged cells by copying rowspan and colspan attributes
                if (cell.hasAttribute('rowspan')) {
                    th.setAttribute('rowspan', cell.getAttribute('rowspan'));
                }
                if (cell.hasAttribute('colspan')) {
                    th.setAttribute('colspan', cell.getAttribute('colspan'));
                }
                
                // Copy any style attributes to preserve cell formatting
                if (cell.hasAttribute('style')) {
                    th.setAttribute('style', cell.getAttribute('style'));
                }
                
                headerTr.appendChild(th);
            });
            
            thead.appendChild(headerTr);
            dataTable.appendChild(thead);
            
            // Process the rest of the rows for the body
            for (let i = 1; i < rows.length; i++) {
                const row = rows[i];
                const cells = row.querySelectorAll('td, th');
                
                // Skip empty rows
                if (cells.length === 0) {
                    continue;
                }
                
                const tr = document.createElement('tr');
                
                cells.forEach(cell => {
                    const td = document.createElement('td');
                    td.innerHTML = cell.innerHTML; // Preserve HTML formatting
                    
                    // Preserve merged cells by copying rowspan and colspan attributes
                    if (cell.hasAttribute('rowspan')) {
                        td.setAttribute('rowspan', cell.getAttribute('rowspan'));
                    }
                    if (cell.hasAttribute('colspan')) {
                        td.setAttribute('colspan', cell.getAttribute('colspan'));
                    }
                    
                    // Copy any style attributes to preserve cell formatting
                    if (cell.hasAttribute('style')) {
                        td.setAttribute('style', cell.getAttribute('style'));
                    }
                    
                    // Copy class attributes if any
                    if (cell.hasAttribute('class')) {
                        td.setAttribute('class', cell.getAttribute('class'));
                    }
                    
                    tr.appendChild(td);
                });
                
                tbody.appendChild(tr);
            }
            
            dataTable.appendChild(tbody);
        }
        
        // Update the UI to show the data has been processed
        pasteArea.innerHTML = '<p>Data processed! Click here to paste new data.</p>';
        pasteArea.classList.add('success');
    }
}); 