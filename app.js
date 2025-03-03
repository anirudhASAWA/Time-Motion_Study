// Time and Motion Study Tool - Frontend JavaScript

// Main application state
const state = {
    columns: 0,
    columnNames: [],
    rows: [],
    timers: {},
    intervalId: null
  };
  
  // DOM Elements
  const columnCountInput = document.getElementById('columnCount');
  const projectNameInput = document.getElementById('projectName');
  const actionPanel = document.getElementById('actionPanel');
  const addRowBtn = document.getElementById('addRowBtn');
  const saveProjectBtn = document.getElementById('saveProjectBtn');
  const exportBtn = document.getElementById('exportBtn');
  const tableContainer = document.getElementById('tableContainer');
  
  // Initialize application
  function init() {
    // Set up event listeners
    columnCountInput.addEventListener('change', handleColumnCountChange);
    addRowBtn.addEventListener('click', addRow);
    saveProjectBtn.addEventListener('click', saveProject);
    exportBtn.addEventListener('click', exportToExcel);
    
    // Start timer update interval
    state.intervalId = setInterval(updateTimers, 100);
  }
  
  // Handle column count change
  function handleColumnCountChange() {
    const count = parseInt(columnCountInput.value) || 0;
    state.columns = count;
    
    // Initialize column names
    state.columnNames = Array(count).fill('').map((_, i) => `Process ${i + 1}`);
    
    // Reset rows and timers
    state.rows = [];
    state.timers = {};
    
    // Show/hide action panel
    actionPanel.classList.toggle('hidden', count === 0);
    
    // Render the table
    renderTable();
  }
  
  // Add a new row
  function addRow() {
    const rowId = Date.now().toString();
    const newRow = { 
      id: rowId,
      name: `Task ${state.rows.length + 1}`
    };
    
    state.rows.push(newRow);
    
    // Initialize timers for this row
    state.columnNames.forEach((col, colIndex) => {
      const timerId = `${rowId}-${colIndex}`;
      state.timers[timerId] = {
        time: 0,
        isRunning: false,
        startTime: null
      };
    });
    
    // Re-render the table
    renderTable();
  }
  
  // Render the table
  function renderTable() {
    if (state.columns === 0) {
      tableContainer.innerHTML = '';
      return;
    }
    
    let tableHTML = `
      <table class="w-full border-collapse">
        <thead>
          <tr class="bg-gray-100 table-header">
            <th class="border p-2 w-48">Task Name</th>
    `;
    
    // Add column headers
    state.columnNames.forEach((colName, colIndex) => {
      tableHTML += `
        <th class="border p-2">
          <input
            type="text"
            value="${colName}"
            data-column-index="${colIndex}"
            class="column-name-input w-full p-1 border rounded"
            placeholder="Process ${colIndex + 1}"
          />
        </th>
      `;
    });
    
    tableHTML += `
          </tr>
        </thead>
        <tbody>
    `;
    
    // Add rows
    state.rows.forEach((row) => {
      tableHTML += `
        <tr>
          <td class="border p-2">
            <input
              type="text"
              value="${row.name}"
              data-row-id="${row.id}"
              class="row-name-input w-full p-1 border rounded"
            />
          </td>
      `;
      
      // Add cells for each column
      state.columnNames.forEach((_, colIndex) => {
        const timerId = `${row.id}-${colIndex}`;
        const timer = state.timers[timerId] || { time: 0, isRunning: false };
        
        tableHTML += `
          <td class="border p-2 text-center">
            <div class="timer-display" id="display-${timerId}">
              ${formatTime(timer.time)}
            </div>
            <div class="timer-controls">
              ${!timer.isRunning ? 
                `<button
                   class="btn-start btn-action"
                   data-action="start"
                   data-row-id="${row.id}"
                   data-col-index="${colIndex}"
                 >
                   Start
                 </button>` : 
                `<button
                   class="btn-stop btn-action"
                   data-action="stop"
                   data-row-id="${row.id}"
                   data-col-index="${colIndex}"
                 >
                   Stop
                 </button>`
              }
              <button
                class="btn-reset btn-action"
                data-action="reset"
                data-row-id="${row.id}"
                data-col-index="${colIndex}"
              >
                Reset
              </button>
            </div>
          </td>
        `;
      });
      
      tableHTML += `
        </tr>
      `;
    });
    
    tableHTML += `
        </tbody>
      </table>
    `;
    
    // Update the table container
    tableContainer.innerHTML = tableHTML;
    
    // Add event listeners to inputs and buttons
    document.querySelectorAll('.column-name-input').forEach(input => {
      input.addEventListener('change', (e) => {
        const colIndex = parseInt(e.target.dataset.columnIndex);
        state.columnNames[colIndex] = e.target.value;
      });
    });
    
    document.querySelectorAll('.row-name-input').forEach(input => {
      input.addEventListener('change', (e) => {
        const rowId = e.target.dataset.rowId;
        const row = state.rows.find(r => r.id === rowId);
        if (row) {
          row.name = e.target.value;
        }
      });
    });
    
    document.querySelectorAll('[data-action]').forEach(button => {
      button.addEventListener('click', (e) => {
        const action = e.target.dataset.action;
        const rowId = e.target.dataset.rowId;
        const colIndex = parseInt(e.target.dataset.colIndex);
        
        if (action === 'start') {
          startTimer(rowId, colIndex);
        } else if (action === 'stop') {
          stopTimer(rowId, colIndex);
        } else if (action === 'reset') {
          resetTimer(rowId, colIndex);
        }
      });
    });
  }
  
  // Timer functions
  function startTimer(rowId, colIndex) {
    const timerId = `${rowId}-${colIndex}`;
    state.timers[timerId] = {
      ...state.timers[timerId],
      isRunning: true,
      startTime: Date.now() - (state.timers[timerId]?.time || 0)
    };
    renderTable();
  }
  
  function stopTimer(rowId, colIndex) {
    const timerId = `${rowId}-${colIndex}`;
    state.timers[timerId] = {
      ...state.timers[timerId],
      isRunning: false
    };
    renderTable();
  }
  
  function resetTimer(rowId, colIndex) {
    const timerId = `${rowId}-${colIndex}`;
    state.timers[timerId] = {
      time: 0,
      isRunning: false,
      startTime: null
    };
    renderTable();
  }
  
  // Update all running timers
  function updateTimers() {
    let updated = false;
    
    Object.keys(state.timers).forEach(timerId => {
      if (state.timers[timerId].isRunning) {
        state.timers[timerId].time = Date.now() - state.timers[timerId].startTime;
        updated = true;
        
        // Update display without re-rendering the entire table
        const displayElement = document.getElementById(`display-${timerId}`);
        if (displayElement) {
          displayElement.textContent = formatTime(state.timers[timerId].time);
        }
      }
    });
    
    // Only re-render if no direct DOM updates happened
    if (updated && !document.querySelector('.timer-display')) {
      renderTable();
    }
  }
  
  // Format time in mm:ss.ms format
  function formatTime(ms) {
    const minutes = Math.floor(ms / 60000);
    const seconds = Math.floor((ms % 60000) / 1000);
    const milliseconds = Math.floor((ms % 1000) / 10);
    
    return `${minutes.toString().padStart(2, '0')}:${seconds.toString().padStart(2, '0')}.${milliseconds.toString().padStart(2, '0')}`;
  }
  
  // Export to Excel (CSV format)
  function exportToExcel() {
    // Create header row
    let csvContent = "Task Name,";
    state.columnNames.forEach(name => {
      csvContent += `${name},`;
    });
    csvContent += "\n";
    
    // Add data rows
    state.rows.forEach(row => {
      csvContent += `${row.name},`;
      state.columnNames.forEach((_, colIndex) => {
        const timerId = `${row.id}-${colIndex}`;
        const timer = state.timers[timerId] || { time: 0 };
        csvContent += `${formatTime(timer.time)},`;
      });
      csvContent += "\n";
    });
    
    // Create download link
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.setAttribute('href', url);
    link.setAttribute('download', `${projectNameInput.value.replace(/\s+/g, '_')}_export.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  }
  
  // Save project data to server
  async function saveProject() {
    const projectName = projectNameInput.value;
    
    try {
      const studyData = {
        projectName,
        columnNames: state.columnNames,
        rows: state.rows,
        timerData: Object.entries(state.timers).reduce((acc, [key, value]) => {
          acc[key] = {
            time: value.time,
            isRunning: false // Always save as stopped
          };
          return acc;
        }, {})
      };
      
      // Make API call to backend
      const response = await fetch('/api/save-project', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(studyData)
      });
      
      if (response.ok) {
        const result = await response.json();
        alert(`Project "${projectName}" saved successfully!`);
      } else {
        throw new Error('Failed to save project');
      }
    } catch (error) {
      console.error("Error saving project:", error);
      alert("Failed to save project. See console for details.");
    }
  }
  
  // Initialize the application when DOM is loaded
  document.addEventListener('DOMContentLoaded', init);