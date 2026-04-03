const input = document.getElementById('incidentInput');
const toggleBtn = document.getElementById('toggleIncidentBtn');
const exportBtn = document.getElementById('exportBtn');
const statusText = document.getElementById('statusText');
const tableBody = document.getElementById('incidentTableBody');

const incidents = [];
let activeIncident = null;

function textAfterLabel(line) {
  const chunks = line.split(/[:：]/);
  return chunks.length > 1 ? chunks.slice(1).join(':').trim() : line.trim();
}

function matchStationToken(text) {
  const m = text.match(/\b(STN[A-Z0-9-]*\s*[A-Z0-9-]*)\b/i);
  return m ? m[1].replace(/\s+/g, ' ').trim() : '';
}

function parseDetails(raw) {
  const parsed = {
    processStep: '',
    failureMode: '',
    effect: 'Line stop',
    cause: '',
    detection: '',
    action: '',
    rawNotes: raw.trim(),
  };

  const lines = raw
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);

  const unlabeled = [];

  for (const line of lines) {
    const lower = line.toLowerCase();

    if (/^(停线时间|downtime|time)\s*[:：]/i.test(line)) {
      continue;
    }

    if (/^(故障工位|工位|station|process\s*step)\s*[:：]/i.test(line)) {
      parsed.processStep = textAfterLabel(line);
      continue;
    }

    if (/^(故障原因|原因|cause|root\s*cause)\s*[:：]/i.test(line)) {
      parsed.cause = textAfterLabel(line);
      continue;
    }

    if (/^(故障模式|failure\s*mode|mode)\s*[:：]/i.test(line)) {
      parsed.failureMode = textAfterLabel(line);
      continue;
    }

    if (/^(影响|effect)\s*[:：]/i.test(line)) {
      parsed.effect = textAfterLabel(line);
      continue;
    }

    if (/^(检出|发现方式|detection)\s*[:：]/i.test(line)) {
      parsed.detection = textAfterLabel(line);
      continue;
    }

    if (/^(处置|措施|action)\s*[:：]/i.test(line)) {
      parsed.action = textAfterLabel(line);
      continue;
    }

    if (!parsed.processStep) {
      const station = matchStationToken(line);
      if (station) {
        parsed.processStep = station;
      }
    }

    if (!parsed.cause && /(原因|异常|错误|故障|fault|error|issue|jam|alarm|二维码|qr)/i.test(lower + line)) {
      parsed.cause = line;
      continue;
    }

    if (!parsed.failureMode && /(停线|line\s*stop|communication|vision|inspection|unloading|carryover|feeder|断线)/i.test(lower + line)) {
      parsed.failureMode = line;
      continue;
    }

    unlabeled.push(line);
  }

  if (!parsed.failureMode && unlabeled.length > 0) {
    parsed.failureMode = unlabeled[0];
  }

  if (!parsed.cause && unlabeled.length > 1) {
    parsed.cause = unlabeled.slice(1).join(' / ');
  }

  if (!parsed.processStep && parsed.rawNotes) {
    parsed.processStep = matchStationToken(parsed.rawNotes);
  }

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
  activeIncident = {
    start: new Date(),
  };

  toggleBtn.textContent = 'Stop & Save Incident / 停止并保存';
  statusText.textContent = `Timing started at ${formatTime(activeIncident.start)}. You can enter details now.`;
}

function endIncident() {
  const end = new Date();
  const downtime = `${formatTime(activeIncident.start)}-${formatTime(end)}`;
  const details = parseDetails(input.value || '');

  incidents.push({
    downtime,
    processStep: details.processStep,
    failureMode: details.failureMode,
    effect: details.effect,
    cause: details.cause,
    detection: details.detection,
    action: details.action,
  });

  activeIncident = null;
  input.value = '';
  toggleBtn.textContent = 'Start Incident Timing / 开始计时';
  statusText.textContent = `Incident saved (${downtime}). Auto-sort finished. Ready for next incident.`;
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
