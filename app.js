const processStepInput = document.getElementById('processStepInput');
const failureModeInput = document.getElementById('failureModeInput');
const effectInput = document.getElementById('effectInput');
const causeInput = document.getElementById('causeInput');
const detectionInput = document.getElementById('detectionInput');
const actionInput = document.getElementById('actionInput');

const toggleBtn = document.getElementById('toggleIncidentBtn');
const exportBtn = document.getElementById('exportBtn');
const statusText = document.getElementById('statusText');
const tableBody = document.getElementById('incidentTableBody');

const incidents = [];
let activeIncident = null;

function formatTime(dateObj) {
  return new Intl.DateTimeFormat('en-GB', {
    hour: '2-digit',
    minute: '2-digit',
    hour12: false,
  }).format(dateObj);
}

function currentFormValues() {
  return {
    processStep: processStepInput.value.trim(),
    failureMode: failureModeInput.value.trim(),
    effect: effectInput.value.trim(),
    cause: causeInput.value.trim(),
    detection: detectionInput.value.trim(),
    action: actionInput.value.trim(),
  };
}

function clearFormValues() {
  processStepInput.value = '';
  failureModeInput.value = '';
  effectInput.value = 'Line stop';
  causeInput.value = '';
  detectionInput.value = '';
  actionInput.value = '';
}

function updateTable() {
  tableBody.innerHTML = '';
  incidents.forEach((row) => {
    const tr = document.createElement('tr');
    [
      row.downtime,
      row.processStep,
      row.failureMode,
      row.effect,
      row.cause,
      row.detection,
      row.action,
    ].forEach((value) => {
      const td = document.createElement('td');
      td.textContent = value || '';
      tr.appendChild(td);
    });
    tableBody.appendChild(tr);
  });
}

function startIncident() {
  activeIncident = { start: new Date() };
  toggleBtn.textContent = 'Stop & Save Incident / 停止并保存';
  statusText.textContent = `Timing started at ${formatTime(activeIncident.start)}.`;
}

function endIncident() {
  const end = new Date();
  const downtime = `${formatTime(activeIncident.start)}-${formatTime(end)}`;

  incidents.push({
    downtime,
    ...currentFormValues(),
  });

  activeIncident = null;
  toggleBtn.textContent = 'Start Incident Timing / 开始计时';
  statusText.textContent = `Incident saved (${downtime}). Ready for next incident.`;
  clearFormValues();
  updateTable();
}

function exportExcel() {
  if (incidents.length === 0) {
    alert('No incidents to export. / 暂无可导出的记录');
    return;
  }

  const worksheetData = [
    ['Downtime', 'Process Step', 'Failure Mode', 'Effect', 'Cause', 'Detection', 'Action'],
    ...incidents.map((r) => [
      r.downtime,
      r.processStep,
      r.failureMode,
      r.effect,
      r.cause,
      r.detection,
      r.action,
    ]),
  ];

  const ws = XLSX.utils.aoa_to_sheet(worksheetData);
  ws['!cols'] = [
    { wch: 14 },
    { wch: 20 },
    { wch: 28 },
    { wch: 14 },
    { wch: 35 },
    { wch: 20 },
    { wch: 28 },
  ];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Downtime');

  const now = new Date();
  const date = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;
  const shift = now.getHours() < 12 ? 'Morning' : now.getHours() < 18 ? 'Afternoon' : 'Night';
  const filename = `Downtime_${date}_${shift}.xlsx`;

  XLSX.writeFile(wb, filename);
}

toggleBtn.addEventListener('click', () => {
  if (activeIncident) {
    endIncident();
  } else {
    startIncident();
  }
});

exportBtn.addEventListener('click', exportExcel);
