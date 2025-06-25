// Meeting Attendance Report App
// Author: Copilot

const REQUIRED_COLUMNS = [
    'Full Name',
    'Email Address',
    'Join Time',
    'Leave Time',
    'Duration (minutes)',
    'Role'
];
const MIN_ATTENDANCE_MINUTES = 2;

document.getElementById('processBtn').addEventListener('click', handleFiles);
document.getElementById('exportBtn').addEventListener('click', exportToExcel);

function handleFiles() {
    const files = document.getElementById('csvFiles').files;
    clearError();
    if (!files.length) {
        showError('Please select at least one CSV file.');
        return;
    }
    const filePromises = Array.from(files).map(file => readFileAsText(file));
    Promise.all(filePromises)
        .then(fileContents => {
            try {
                const allMeetings = [];
                fileContents.forEach((content, idx) => {
                    const { meetingDate, records } = parseCsv(content, idx, files[idx].name);
                    allMeetings.push({ meetingDate, records });
                });
                const summary = generateSummary(allMeetings);
                renderSummaryTable(summary);
                document.getElementById('summaryTableSection').style.display = 'block';
            } catch (err) {
                showError(err.message);
            }
        })
        .catch(err => showError('Error reading files: ' + err.message));
}

function readFileAsText(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = e => resolve(e.target.result);
        reader.onerror = e => reject(e);
        reader.readAsText(file);
    });
}

function parseCsv(content, fileIdx, fileName) {
    // Simple CSV parser (assumes no commas in fields)
    const lines = content.split(/\r?\n/).filter(Boolean);
    if (!lines.length) throw new Error(`File ${fileName} is empty.`);
    const headers = lines[0].split(',').map(h => h.trim());
    const missing = REQUIRED_COLUMNS.filter(col => !headers.includes(col));
    if (missing.length) {
        throw new Error(`File ${fileName} is missing required columns: ${missing.join(', ')}`);
    }
    const colIdx = Object.fromEntries(headers.map((h, i) => [h, i]));
    const records = lines.slice(1).map(line => {
        const cells = line.split(',');
        return Object.fromEntries(headers.map((h, i) => [h, cells[i] ? cells[i].trim() : '']));
    });
    // Meeting date: use first Join Time (or Organizer's Join Time if available)
    let meetingDate = '';
    for (const rec of records) {
        if (rec['Join Time']) {
            meetingDate = rec['Join Time'].split(' ')[0];
            break;
        }
    }
    if (!meetingDate) meetingDate = `Meeting ${fileIdx+1}`;
    return { meetingDate, records };
}

function generateSummary(allMeetings) {
    // Collect unique participants and meetings
    const participants = new Set();
    const meetings = [];
    const attendance = {};
    allMeetings.forEach(({ meetingDate, records }) => {
        meetings.push(meetingDate);
        // Find organizer's presence window
        const organizerRecords = records.filter(r => r['Role'].toLowerCase() === 'organizer');
        if (!organizerRecords.length) throw new Error(`No organizer found for meeting on ${meetingDate}.`);
        const orgJoin = parseDate(organizerRecords[0]['Join Time']);
        const orgLeave = parseDate(organizerRecords[0]['Leave Time']);
        if (!orgJoin || !orgLeave) throw new Error(`Invalid organizer times for meeting on ${meetingDate}.`);
        // For each participant, sum attendance within organizer window
        const meetingAttendance = {};
        records.forEach(r => {
            if (!r['Full Name'] || !r['Join Time'] || !r['Leave Time']) return;
            if (r['Role'].toLowerCase() === 'organizer') return;
            const join = parseDate(r['Join Time']);
            const leave = parseDate(r['Leave Time']);
            if (!join || !leave) return;
            // Clamp to organizer window
            const actualJoin = join < orgJoin ? orgJoin : join;
            const actualLeave = leave > orgLeave ? orgLeave : leave;
            let duration = (actualLeave - actualJoin) / 60000; // ms to min
            if (duration < 0) duration = 0;
            if (duration >= MIN_ATTENDANCE_MINUTES) {
                const key = r['Full Name'] + (r['Email Address'] ? ` (${r['Email Address']})` : '');
                participants.add(key);
                meetingAttendance[key] = (meetingAttendance[key] || 0) + duration;
            }
        });
        attendance[meetingDate] = meetingAttendance;
    });
    return { participants: Array.from(participants).sort(), meetings, attendance };
}

function parseDate(str) {
    // Expects format: YYYY-MM-DD HH:MM AM/PM
    if (!str) return null;
    const d = new Date(str.replace(/\u200E/g, ''));
    return isNaN(d) ? null : d;
}

function renderSummaryTable({ participants, meetings, attendance }) {
    const container = document.getElementById('summaryTableContainer');
    let html = '<table><thead><tr><th>Participant</th>';
    meetings.forEach(m => html += `<th>${m}</th>`);
    html += '</tr></thead><tbody>';
    participants.forEach(p => {
        html += `<tr><td>${p}</td>`;
        meetings.forEach(m => {
            const mins = attendance[m][p] ? Math.round(attendance[m][p]) : '';
            html += `<td>${mins}</td>`;
        });
        html += '</tr>';
    });
    html += '</tbody></table>';
    container.innerHTML = html;
}

function exportToExcel() {
    const table = document.querySelector('#summaryTableContainer table');
    if (!table) return;
    const wb = XLSX.utils.table_to_book(table, { sheet: 'Attendance Summary' });
    XLSX.writeFile(wb, 'attendance_summary.xlsx');
}

function showError(msg) {
    document.getElementById('errorMsg').textContent = msg;
    document.getElementById('summaryTableSection').style.display = 'none';
}
function clearError() {
    document.getElementById('errorMsg').textContent = '';
}
