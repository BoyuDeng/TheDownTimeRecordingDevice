const input = document.getElementById('incidentInput');
const toggleBtn = document.getElementById('toggleIncidentBtn');
const exportBtn = document.getElementById('exportBtn');
const statusText = document.getElementById('statusText');
const tableBody = document.getElementById('incidentTableBody');

const incidents = [];
let activeIncident = null;

const labelRules = [
  { regex: /^(故障工位|工位|process\s*step|station)\s*[:：]\s*(.+)$/i, field: 'processStep' },
  { regex: /^(故障原因|原因|cause|root\s*cause)\s*[:：]\s*(.+)$/i, field: 'cause' },
  { regex: /^(故障模式|failure\s*mode|mode)\s*[:：]\s*(.+)$/i, field: 'failureMode' },
  { regex: /^(影响|effect)\s*[:：]\s*(.+)$/i, field: 'effect' },
  { regex: /^(检出|发现方式|detection)\s*[:：]\s*(.+)$/i, field: 'detection' },
  { regex: /^(处置|措施|action)\s*[:：]\s*(.+)$/i, field: 'action' },
];

function parseDetails(raw) {
  const parsed = {
    processStep: '',
    failureMode: '',
    effect: 'Line stop',
    cause: '',
    detection: '',
    action: '',
  };

  const lines = raw
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);

  const leftovers = [];
  for (const line of lines) {
    let matched = false;
    for (const rule of labelRules) {
      const m = line.match(rule.regex);
      if (m) {
        parsed[rule.field] = m[2].trim();
        matched = true;
        break;
      }
    }
    if (!matched && !/^停线时间\s*[:：]/.test(line)) {
      leftovers.push(line);
    }
  }

  if (!parsed.failureMode && leftovers.length > 0) parsed.failureMode = leftovers[0];
  if (!parsed.cause && leftovers.length > 1) parsed.cause = leftovers.slice(1).join(' / ');

  return parsed;
}

function formatTime(dateObj) {
  return new Intl.DateTimeFormat('en-GB', {
    hour: '2-digit',
    minute: '2-digit',
    hour12: false,
  }).format(dateObj);
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
  const text = input.value.trim();
  if (!text) {
    alert('Please enter incident details first. / 请先输入故障信息');
    return;
  }

  activeIncident = {
    start: new Date(),
    details: parseDetails(text),
  };

  toggleBtn.textContent = 'End Incident / 结束停线';
  statusText.textContent = `Incident started at ${formatTime(activeIncident.start)}.`;
}

function endIncident() {
  const end = new Date();
  const downtime = `${formatTime(activeIncident.start)}-${formatTime(end)}`;
  incidents.push({
    downtime,
    ...activeIncident.details,
  });

  activeIncident = null;
  input.value = '';
  toggleBtn.textContent = 'Start Incident / 开始停线';
  statusText.textContent = `Incident saved (${downtime}). Ready for next incident.`;
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
