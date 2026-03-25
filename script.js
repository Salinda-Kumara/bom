// ===== CHART.JS DEFAULTS — PREMIUM MODERN UI =====
Chart.defaults.color = '#64748b';
Chart.defaults.font.family = "'Inter', system-ui, sans-serif";
Chart.defaults.font.size = 13;
Chart.defaults.plugins.legend.labels.usePointStyle = true;
Chart.defaults.plugins.legend.labels.pointStyleWidth = 12;
Chart.defaults.plugins.legend.labels.padding = 24;
Chart.defaults.elements.bar.borderRadius = 8;
Chart.defaults.elements.arc.borderWidth = 0;
Chart.register(ChartDataLabels);

// Premium Modern color palette (keeping 'PPT' variable name to avoid refactoring)
const PPT = {
    blue: '#3b82f6', orange: '#f59e0b', gray: '#94a3b8',
    gold: '#fbbf24', dblue: '#1d4ed8', green: '#10b981',
    red: '#ef4444', teal: '#14b8a6', purple: '#8b5cf6',
    ltblue: '#60a5fa', dgreen: '#047857', brown: '#b45309',
};
const PPT_COLORS = [PPT.blue, PPT.purple, PPT.orange, PPT.teal, PPT.green, PPT.ltblue, PPT.gold, PPT.red];

const gridColor = '#f1f5f9';
const tooltipConfig = {
    backgroundColor: '#ffffff', titleColor: '#0f172a', bodyColor: '#475569',
    borderColor: '#e2e8f0', borderWidth: 1,
    titleFont: { weight: '800', size: 14 }, bodyFont: { size: 13, weight: '500' },
    padding: 16, cornerRadius: 12, displayColors: true, boxPadding: 6,
};

let chartInstances = {};
function destroyCharts() {
    Object.values(chartInstances).forEach(c => c.destroy());
    chartInstances = {};
}

// ===== FILE UPLOAD =====
const uploadOverlay = document.getElementById('uploadOverlay');
const dropzone = document.getElementById('dropzone');
const fileInput = document.getElementById('fileInput');
const uploadProgress = document.getElementById('uploadProgress');
const progressFill = document.getElementById('progressFill');
const progressText = document.getElementById('progressText');
const sidebar = document.getElementById('sidebar');
const mainContent = document.getElementById('mainContent');
const reloadBtn = document.getElementById('reloadBtn');

dropzone.addEventListener('click', () => fileInput.click());
dropzone.addEventListener('dragover', e => { e.preventDefault(); dropzone.classList.add('dragover'); });
dropzone.addEventListener('dragleave', () => dropzone.classList.remove('dragover'));
dropzone.addEventListener('drop', e => {
    e.preventDefault(); dropzone.classList.remove('dragover');
    if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
});
fileInput.addEventListener('change', e => { if (e.target.files.length) handleFile(e.target.files[0]); });

reloadBtn.addEventListener('click', () => {
    destroyCharts();
    sidebar.style.display = 'none';
    mainContent.style.display = 'none';
    uploadOverlay.classList.remove('hidden');
    uploadProgress.style.display = 'none';
    progressFill.style.width = '0%';
    fileInput.value = '';
});

function handleFile(file) {
    if (!file.name.match(/\.xlsx?$/i)) { alert('Please upload an Excel file (.xlsx or .xls)'); return; }
    uploadProgress.style.display = 'block';
    progressFill.style.width = '20%';
    progressText.textContent = 'Reading file...';

    const reader = new FileReader();
    reader.onload = function (e) {
        progressFill.style.width = '50%';
        progressText.textContent = 'Parsing data...';
        setTimeout(() => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array', cellDates: true });
                progressFill.style.width = '80%';
                progressText.textContent = 'Building dashboard...';
                setTimeout(() => {
                    processWorkbook(workbook);
                    progressFill.style.width = '100%';
                    progressText.textContent = 'Done!';
                    setTimeout(() => {
                        uploadOverlay.classList.add('hidden');
                        sidebar.style.display = 'flex';
                        mainContent.style.display = 'block';
                        initNavigation();
                        initAnimations();
                    }, 400);
                }, 200);
            } catch (err) {
                alert('Error parsing file: ' + err.message);
                uploadProgress.style.display = 'none';
                progressFill.style.width = '0%';
            }
        }, 100);
    };
    reader.readAsArrayBuffer(file);
}

// ===== PROCESS WORKBOOK =====
function processWorkbook(wb) {
    destroyCharts();
    const sheets = {};
    wb.SheetNames.forEach(name => {
        sheets[name] = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: null });
    });

    // Find sheet by flexible name matching
    function findSheet(keyword) {
        const key = Object.keys(sheets).find(k => k.toLowerCase().replace(/\s+/g, '').includes(keyword.toLowerCase().replace(/\s+/g, '')));
        return key ? sheets[key] : null;
    }

    const lectureData = parseLectureBreakdown(sheets['Lecture_Breakdown']);
    const examData = parseExamStatus(sheets['Exam_Status']);
    const enrollData = parseEnrollments(sheets['Enrollments']);
    const batchTransfers = parseSumMetric(sheets['Batch_Transfers']);
    const dropouts = parseSumMetric(sheets['Dropouts']);
    const commencements = parseSchedule(sheets['Lecture_Commencements']);
    const endings = parseSchedule(sheets['Lecture_Endings']);
    const highlightsData = parseHighlights(sheets['Highlights'] || findSheet('Highlight'));
    const internshipData = parseInternship(sheets['Internship'] || findSheet('Internship'));
    const counsellingData = parseCounselling(findSheet('Counselling'));
    const meetingsData = parseHRTable(findSheet('Meetings') || findSheet('Meeting'));
    const recruitmentData = parseHRTable(findSheet('Recruitment') || findSheet('Recruit'));

    buildKPIs(lectureData, enrollData, examData, batchTransfers, dropouts, counsellingData);
    buildLectureChart(lectureData);
    buildMetricPills(lectureData, batchTransfers, dropouts);
    buildScheduleTable('commencementsTable', commencements);
    buildScheduleTable('endingsTable', endings);
    buildHighlightsSection(highlightsData);
    buildEnrollmentTable(enrollData);
    buildEnrollmentCharts(enrollData);
    buildExamCharts(examData);
    buildExamCards(examData);
    buildInternshipSection(internshipData);
    buildCounsellingChart(counsellingData);
    // buildWelfareSummary(internshipData, counsellingData);
    buildHRSection(meetingsData, recruitmentData);

    // Trigger staggered animations for modern UI effect
    setTimeout(() => {
        document.querySelectorAll('.card, .kpi-card, .event-card').forEach((el, i) => {
            el.classList.add('fade-in');
            el.style.animationDelay = `${(i % 10) * 0.05}s`;
        });
    }, 100);
}

// ===== PARSERS =====
function parseLectureBreakdown(rows) {
    if (!rows || rows.length < 2) return [];
    const result = [];
    for (let i = 1; i < rows.length; i++) {
        const r = rows[i];
        if (r && r[0] && typeof r[0] === 'string' && r[0] !== 'Degree')
            result.push({ degree: r[0], mode: r[1], count: Number(r[2]) || 0 });
    }
    return result;
}

function parseExamStatus(rows) {
    if (!rows || rows.length < 2) return [];
    const result = [];
    for (let i = 1; i < rows.length; i++) {
        const r = rows[i]; if (!r) continue;
        const cycle = r[0];
        if (!cycle || typeof cycle !== 'string' || cycle === 'Exam_Cycle') continue;
        const papers = Number(r[1]) || 0;
        const released = Number(r[2]) || 0;
        // percentage can be a formula string like "=C2/B2*100" or a number
        let pct = r[3];
        if (typeof pct === 'string' && pct.startsWith('=')) {
            pct = papers > 0 ? ((released / papers) * 100) : 0;
        } else {
            pct = Number(pct) || (papers > 0 ? ((released / papers) * 100) : 0);
        }
        result.push({ cycle, papers, released, percentage: Math.round(pct * 10) / 10 });
    }
    return result;
}
function parseEnrollments(rows) {
    if (!rows || rows.length < 5) return { budgetActual: [], mohePipeline: [] };
    const budgetActual = [];
    const mohePipeline = [];

    // Parse Budget vs Actual (Rows 2 to 13)
    let currentProgramme = '';
    for (let i = 2; i <= 13; i++) {
        const r = rows[i]; if (!r) continue;
        if (r[0] && String(r[0]).trim()) {
            currentProgramme = String(r[0]).trim();
        }
        const type = r[1] ? String(r[1]).trim() : '';
        
        let finalNumber = 0;
        for (let j = r.length - 1; j >= 2; j--) {
            if (typeof r[j] === 'number') {
                finalNumber = r[j];
                break;
            }
        }

        if (type === 'Budget') {
            budgetActual.push({ name: currentProgramme, budget: finalNumber, actual: 0 });
        } else if (type === 'Actual') {
            const entry = budgetActual.find(b => b.name === currentProgramme);
            if (entry) entry.actual = finalNumber;
        }
    }

    // Parse MoHE Admissions Pipeline
    let pipelineStart = -1;
    for (let i = 10; i < rows.length; i++) {
        const r = rows[i];
        if (r && r[1] && String(r[1]).toLowerCase().includes('applications received')) {
            pipelineStart = i + 1;
            break;
        }
    }
    if (pipelineStart > -1) {
        for (let i = pipelineStart; i < rows.length; i++) {
            const r = rows[i]; if (!r || !r[0]) continue;
            const name = String(r[0]).trim();
            if (name.toLowerCase() === 'total' || name.toLowerCase() === 'degree') continue;
            
            const received = parseInt(String(r[1]).replace(/[^0-9]/g, '')) || 0;
            const selected = parseInt(String(r[2]).replace(/[^0-9]/g, '')) || 0;
            const registered = parseInt(String(r[3]).replace(/[^0-9]/g, '')) || 0;
            
            mohePipeline.push({ name, received, selected, registered });
        }
    }

    return { budgetActual, mohePipeline };
}

// Sum all numeric values across all rows (for Batch_Transfers, Dropouts with monthly data)
function parseSumMetric(rows) {
    if (!rows) return 0;
    let total = 0;
    for (let i = 1; i < rows.length; i++) { // skip header
        const r = rows[i]; if (!r) continue;
        for (let c = 1; c < r.length; c++) {
            if (typeof r[c] === 'number') total += r[c];
        }
    }
    return total || 0;
}

// Parse Commencements / Endings: columns are [Label, Date, Batch, Y:S]
function parseSchedule(rows) {
    if (!rows || rows.length < 2) return [];
    const result = [];
    for (let i = 1; i < rows.length; i++) {
        const r = rows[i]; if (!r) continue;
        const batch = r[2] ? String(r[2]).trim().replace(/\u00a0/g, ' ') : '';
        const semester = r[3] ? String(r[3]).trim() : '';
        if (!batch) continue;

        let dateStr = '';
        const dateVal = r[1];
        if (dateVal instanceof Date) {
            dateStr = dateVal.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
        } else if (dateVal) {
            const d = new Date(dateVal);
            if (!isNaN(d)) dateStr = d.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
            else dateStr = String(dateVal).split('T')[0];
        }
        result.push({ batch, semester, date: dateStr });
    }
    return result;
}

// Parse Highlights sheet: guest lectures (rows 1-4 with 9 cols) + events (rows 7+)
function parseHighlights(rows) {
    if (!rows || rows.length < 2) return { guestLectures: [], events: [] };
    const guestLectures = [];
    const events = [];

    // Guest lectures: row[0]=No, [1]=Date, [2]=Time, [3]=Venue, [4]=Resource Person, [5]=Qualifications, [6]=Topic, [7]=Subject, [8]=Batch
    for (let i = 1; i < rows.length; i++) {
        const r = rows[i]; if (!r) continue;
        // Detect event section (row with "Event" in col 1)
        if (r[1] && String(r[1]).trim().toLowerCase() === 'event') {
            // Parse events from here
            for (let j = i + 1; j < rows.length; j++) {
                const ev = rows[j]; if (!ev) continue;
                const name = ev[1] ? String(ev[1]).trim() : '';
                
                // Stop parsing if we reach the Achievements section header
                if (name.toLowerCase().includes('achivement') || name.toLowerCase().includes('achievement')) {
                    break;
                }
                
                const date = ev[2] ? String(ev[2]).trim() : '';
                if (name) events.push({ name, date });
            }
            break;
        }
        // Guest lecture row
        if (r[0] && (typeof r[0] === 'number' || !isNaN(Number(r[0])))) {
            guestLectures.push({
                no: r[0],
                date: r[1] ? String(r[1]).trim() : '',
                time: r[2] ? String(r[2]).trim() : '',
                venue: r[3] ? String(r[3]).trim() : '',
                person: r[4] ? String(r[4]).trim() : '',
                qualifications: r[5] ? String(r[5]).trim() : '',
                topic: r[6] ? String(r[6]).trim() : '',
                subject: r[7] ? String(r[7]).trim() : '',
                batch: r[8] ? String(r[8]).trim() : '',
            });
        }
    }
    return { guestLectures, events };
}

// Parse Internship sheet
function parseInternship(rows) {
    if (!rows || rows.length < 2) return null;
    const data = {
        total: 0,
        audit: { total: 0, big3: 0, other: 0 },
        nonAudit: 0,
        auditPct: 0, nonAuditPct: 0,
        caRegistered: 0, caNotRegistered: 0,
        caRegPct: 0, caNotRegPct: 0,
    };

    for (let i = 0; i < rows.length; i++) {
        const r = rows[i]; if (!r || !r[0]) continue;
        const label = String(r[0]).trim().toLowerCase();
        const val = Number(r[1]) || 0;
        const pct = r[2] != null ? Number(r[2]) : null;

        if (label === 'audit' && pct !== null) {
            // This is the ratio row (Audit with percentage)
            data.auditPct = pct;
        } else if (label === 'audit') {
            data.audit.total = val;
        } else if (label === 'big 3') {
            data.audit.big3 = val;
        } else if (label.includes('other')) {
            data.audit.other = val;
        } else if (label === 'non audit' && pct !== null) {
            data.nonAuditPct = pct;
        } else if (label === 'non audit') {
            data.nonAudit = val;
        } else if (label === 'total') {
            data.total = val;
        } else if (label.includes('ca registered') || label.includes('ca registere')) {
            data.caRegistered = val;
            if (pct !== null) data.caRegPct = pct;
        } else if (label.includes('ca not')) {
            data.caNotRegistered = val;
            if (pct !== null) data.caNotRegPct = pct;
        }
    }
    return data;
}

// Parse Counselling sessions
function parseCounselling(rows) {
    if (!rows || rows.length < 2) return [];
    const result = [];
    for (let i = 1; i < rows.length; i++) {
        const r = rows[i]; if (!r) continue;
        const month = r[0] ? String(r[0]).trim() : '';
        const count = Number(r[1]) || 0;
        if (month) result.push({ month, count });
    }
    return result;
}

// ===== BUILDERS =====
function buildKPIs(lectures, enrollments, exams, transfers, dropouts, counselling) {
    const totalLectures = lectures.reduce((s, l) => s + l.count, 0);
    const lectureSub = lectures.map(l => `${l.degree} ${l.mode}: ${l.count}`).join(' | ');
    let totalBudget = 0, totalActual = 0;
    enrollments.budgetActual.forEach(p => {
        if (p.name.toUpperCase() === 'TOTAL') {
            totalBudget = p.budget || 0;
            totalActual = p.actual || 0;
        }
    });
    const totalPapers = exams.reduce((s, e) => s + e.papers, 0);
    const totalReleased = exams.reduce((s, e) => s + e.released, 0);
    const months = [...new Set(exams.map(e => e.cycle))];
    const totalCounselling = counselling.reduce((s, c) => s + c.count, 0);

    document.getElementById('kpiRow').innerHTML = `
        <div class="kpi-card">
            <div class="kpi-icon"><svg width="26" height="26" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M2 3h6a4 4 0 0 1 4 4v14a3 3 0 0 0-3-3H2z"/><path d="M22 3h-6a4 4 0 0 0-4 4v14a3 3 0 0 1 3-3h7z"/></svg></div>
            <div class="kpi-value">${totalLectures}</div>
            <div class="kpi-label">Total Lectures</div>
            <div class="kpi-sub">${lectureSub || '—'}</div>
        </div>
        <div class="kpi-card">
            <div class="kpi-icon"><svg width="26" height="26" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg></div>
            <div class="kpi-value">${totalActual || totalBudget}</div>
            <div class="kpi-label">Total Enrollments</div>
            <div class="kpi-sub">Budget: ${totalBudget} | Actual: ${totalActual}</div>
        </div>
        <div class="kpi-card">
            <div class="kpi-icon"><svg width="26" height="26" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/></svg></div>
            <div class="kpi-value">${totalPapers}</div>
            <div class="kpi-label">Exam Papers</div>
            <div class="kpi-sub">${totalReleased} Results Released | ${months.join(', ')}</div>
        </div>
        <div class="kpi-card">
            <div class="kpi-icon"><svg width="26" height="26" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="17 1 21 5 17 9"/><path d="M3 11V9a4 4 0 0 1 4-4h14"/><polyline points="7 23 3 19 7 15"/><path d="M21 13v2a4 4 0 0 1-4 4H3"/></svg></div>
            <div class="kpi-value">${transfers}</div>
            <div class="kpi-label">Batch Transfers</div>
            <div class="kpi-sub">Dropouts: ${dropouts} | Counselling: ${totalCounselling}</div>
        </div>`;
    document.getElementById('heroMonthBadge').textContent = 'Data Loaded Successfully';
}

function buildLectureChart(lectures) {
    if (!lectures.length) return;
    const labels = lectures.map(l => `${l.degree} - ${l.mode}`);
    const data = lectures.map(l => l.count);

    chartInstances['lecture'] = new Chart(document.getElementById('lectureBreakdownChart'), {
        type: 'bar',
        data: {
            labels,
            datasets: [{
                label: 'Lecture Count',
                data,
                backgroundColor: data.map((_, i) => PPT_COLORS[i % PPT_COLORS.length]),
                borderColor: data.map((_, i) => PPT_COLORS[i % PPT_COLORS.length]),
                borderWidth: 1, borderRadius: 4, borderSkipped: false,
            }]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            plugins: {
                legend: { display: false }, tooltip: tooltipConfig,
                datalabels: { anchor: 'end', align: 'top', color: '#1a1d23', font: { weight: '700', size: 13 } }
            },
            scales: {
                y: { beginAtZero: true, grid: { color: gridColor }, border: { display: false } },
                x: { grid: { display: false }, border: { display: false } }
            }
        }
    });
}

function buildMetricPills(lectures, transfers, dropouts) {
    const degreeGroups = {};
    lectures.forEach(l => degreeGroups[l.degree] = (degreeGroups[l.degree] || 0) + l.count);
    let html = '';
    const colors = ['blue', 'green', 'gold', 'red'];
    let i = 0;
    for (const [deg, count] of Object.entries(degreeGroups)) {
        html += `<div class="metric-pill"><span class="pill-label">${deg} Total</span><span class="pill-value ${colors[i % colors.length]}">${count}</span></div>`;
        i++;
    }
    html += `<div class="metric-pill"><span class="pill-label">Batch Transfers</span><span class="pill-value gold">${transfers}</span></div>`;
    html += `<div class="metric-pill"><span class="pill-label">Dropouts</span><span class="pill-value red">${dropouts}</span></div>`;
    document.getElementById('metricPills').innerHTML = html;
}

function buildScheduleTable(containerId, items) {
    const el = document.getElementById(containerId);
    if (!items.length) { el.innerHTML = '<p style="color:var(--text-muted);font-size:13px;padding:16px;">No data</p>'; return; }
    let html = '<table class="data-table compact"><thead><tr><th>Date</th><th>Batch</th><th>Year : Semester</th></tr></thead><tbody>';
    items.forEach(it => {
        html += `<tr><td>${it.date}</td><td><strong>${it.batch}</strong></td><td>${it.semester}</td></tr>`;
    });
    el.innerHTML = html + '</tbody></table>';
}

// ===== HIGHLIGHTS BUILDERS =====
function buildHighlightsSection(data) {
    const lectureEl = document.getElementById('guestLecturesTable');
    const eventsEl = document.getElementById('eventsGrid');

    if (!data || !data.guestLectures.length) {
        lectureEl.innerHTML = '<p style="color:var(--text-muted);font-size:13px;padding:16px;">No guest lecture data</p>';
    } else {
        let html = '<table class="data-table guest-lecture-table"><thead><tr><th>#</th><th>Date</th><th>Topic</th><th>Resource Person</th><th>Subject</th><th>Batch</th></tr></thead><tbody>';
        data.guestLectures.forEach(gl => {
            html += `<tr>
                <td class="center">${gl.no}</td>
                <td class="nowrap">${gl.date}</td>
                <td>
                    <div class="gl-topic">${gl.topic}</div>
                    <div class="gl-meta">${gl.time} • ${gl.venue}</div>
                </td>
                <td>
                    <div class="gl-person">${gl.person}</div>
                    <div class="gl-quals">${gl.qualifications}</div>
                </td>
                <td>${gl.subject}</td>
                <td>${gl.batch}</td>
            </tr>`;
        });
        lectureEl.innerHTML = html + '</tbody></table>';
    }

    if (!data || !data.events.length) {
        eventsEl.innerHTML = '<p style="color:var(--text-muted);font-size:13px;padding:16px;">No events data</p>';
    } else {
        const imageMapping = [
            { key: 'thai pongal', file: 'Thai pongal.jpg' },
            { key: 'blood', file: 'Blood Doation Campaign.jpg' },
            { key: 'iftar', file: 'Iftar Celebration.jpg' },
            { key: 'inauguration', file: 'Inauguration Ceremony Intake 20A & 4A.jpg' },
            { key: 'registration', file: '10th IFSL Intake Students Registration.jpg' },
            { key: 'ifsl', file: '10th IFSL Intake Students Registration.jpg' }
        ];

        function getEventImage(name) {
            if (!name) return 'https://images.unsplash.com/photo-1540575467063-178a50c2df87?w=600&q=80';
            const lower = name.toLowerCase();
            for (let map of imageMapping) {
                if (lower.includes(map.key)) return encodeURI(map.file);
            }
            // fallback
            return 'https://images.unsplash.com/photo-1540575467063-178a50c2df87?w=600&q=80';
        }
        
        let html = '<div class="roadmap"><div class="roadmap-line"></div>';
        data.events.forEach((ev, idx) => {
            const imgUrl = getEventImage(ev.name);
            const delay = (idx % 10) * 0.1;
            html += `<div class="roadmap-item fade-in" style="animation-delay: ${delay}s">
                <div class="roadmap-marker"></div>
                <div class="roadmap-content">
                    <div class="roadmap-image-wrap">
                        <img src="${imgUrl}" alt="${ev.name}" class="roadmap-image" />
                    </div>
                    <div class="roadmap-details">
                        <div class="roadmap-date"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" style="vertical-align:text-top; margin-right:4px;"><rect x="3" y="4" width="18" height="18" rx="2" ry="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg> ${ev.date}</div>
                        <div class="roadmap-name">${ev.name}</div>
                        
                    </div>
                </div>
            </div>`;
        });
        html += '</div>';
        eventsEl.innerHTML = html;
        eventsEl.className = 'roadmap-container'; // Override the previous grid class
    }
}

// ===== ENROLLMENT BUILDERS =====
function buildEnrollmentTable(enrollData) {
    const wrap = document.getElementById('enrollmentTableWrap');
    if (!enrollData.budgetActual || !enrollData.budgetActual.length) { wrap.innerHTML = '<p style="color:var(--text-muted);padding:16px;">No enrollment data</p>'; return; }
    
    let html = '<table class="data-table enrollment-table" style="width:100%;"><thead><tr><th>Degree Programme</th><th>Budget</th><th>Actual</th></tr></thead><tbody>';
    
    enrollData.budgetActual.forEach(p => {
        const isTotal = p.name.toLowerCase() === 'total';
        const variance = p.actual - p.budget;
        
        let varColor = variance < 0 ? 'var(--c-red)' : 'var(--c-green)';
        let varIcon = variance < 0 ? '↓' : '↑';
        if (variance === 0) { varColor = 'var(--text-sec)'; varIcon = '-'; }
        
        let rowStyle = isTotal ? 'background:#f8fafc;' : '';
        let nameStyle = isTotal ? 'font-weight:800;' : 'font-weight:600; color:var(--c-blue);';
        let valStyle = isTotal ? 'font-weight:800;' : '';
        
        html += `<tr style="${rowStyle}">`;
        html += `<td style="${nameStyle}">${p.name}</td>`;
        html += `<td style="${valStyle}">${p.budget}</td>`;
        html += `<td style="${valStyle}">${p.actual}</td>`;
        // html += `<td style="color:${varColor}; font-weight:700;">${varIcon} ${Math.abs(variance)}</td>`;
        html += '</tr>';
    });
    html += '</tbody></table>';
    wrap.innerHTML = html;
}

function buildEnrollmentCharts(enrollData) {
    if (!enrollData.budgetActual || !enrollData.mohePipeline) return;

    const progNames = [], budgetData = [], actualData = [];
    enrollData.budgetActual.forEach(p => {
        if (p.name.toLowerCase() === 'total') return;
        const name = p.name.length > 20 ? p.name.substring(0, 20) + '...' : p.name;
        progNames.push(name);
        budgetData.push(p.budget);
        actualData.push(p.actual);
    });

    chartInstances['enrollment'] = new Chart(document.getElementById('enrollmentChart'), {
        type: 'bar',
        data: {
            labels: progNames,
            datasets: [
                { label: 'Budget', data: budgetData, backgroundColor: PPT.blue, borderColor: PPT.dblue, borderWidth: 1, borderRadius: 4 },
                { label: 'Actual', data: actualData, backgroundColor: PPT.orange, borderColor: '#c66520', borderWidth: 1, borderRadius: 4 }
            ]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            plugins: {
                legend: { position: 'top' }, tooltip: tooltipConfig,
                datalabels: { anchor: 'end', align: 'top', color: '#1a1d23', font: { weight: '700', size: 12 } }
            },
            scales: {
                y: { beginAtZero: true, grid: { color: gridColor }, border: { display: false } },
                x: { grid: { display: false }, border: { display: false } }
            }
        }
    });

    const moheNames = [], moheReceived = [], moheSelected = [], moheRegistered = [];
    enrollData.mohePipeline.forEach(p => {
        moheNames.push(p.name);
        moheReceived.push(p.received);
        moheSelected.push(p.selected);
        moheRegistered.push(p.registered);
    });

    chartInstances['mohePipeline'] = new Chart(document.getElementById('mohePipelineChart'), {
        type: 'bar',
        data: {
            labels: moheNames,
            datasets: [
                { label: 'Applications Received', data: moheReceived, backgroundColor: 'rgba(56, 189, 248, 0.6)', borderColor: '#38bdf8', borderWidth: 1, borderRadius: 4 },
                { label: 'Selected applicants', data: moheSelected, backgroundColor: 'rgba(96, 165, 250, 0.8)', borderColor: '#60a5fa', borderWidth: 1, borderRadius: 4 },
                { label: 'Registered', data: moheRegistered, backgroundColor: PPT.blue, borderColor: PPT.dblue, borderWidth: 1, borderRadius: 4 }
            ]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            plugins: {
                legend: { position: 'bottom', labels: { boxWidth: 12 } }, tooltip: tooltipConfig,
                datalabels: { anchor: 'center', align: 'center', color: '#fff', font: { weight: '800', size: 11 } }
            },
            scales: {
                y: { beginAtZero: true, grid: { color: gridColor }, border: { display: false } },
                x: { grid: { display: false }, border: { display: false } }
            }
        }
    });
}

// ===== EXAM BUILDERS =====
function buildExamCharts(examData) {
    if (!examData.length) return;
    const labels = examData.map(e => e.cycle);

    chartInstances['exam'] = new Chart(document.getElementById('examStatusChart'), {
        type: 'bar',
        data: {
            labels,
            datasets: [
                {
                    label: 'No of Exam Papers',
                    data: examData.map(e => e.papers),
                    backgroundColor: PPT.blue,
                    borderColor: PPT.dblue,
                    borderWidth: 1, borderRadius: 3, borderSkipped: false,
                },
                {
                    label: 'Results Released',
                    data: examData.map(e => e.released),
                    backgroundColor: PPT.green,
                    borderColor: PPT.dgreen,
                    borderWidth: 1, borderRadius: 3, borderSkipped: false,
                }
            ]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            plugins: {
                legend: { position: 'top' }, tooltip: tooltipConfig,
                datalabels: { anchor: 'end', align: 'top', color: '#1a1d23', font: { weight: '600', size: 11 } }
            },
            scales: {
                y: { beginAtZero: true, grid: { color: gridColor }, border: { display: false }, title: { display: true, text: 'Count', color: '#5a6275' } },
                x: { grid: { display: false }, border: { display: false } }
            }
        }
    });
}

function buildExamCards(examData) {
    if (!examData.length) { document.getElementById('examCardsGrid').innerHTML = ''; return; }
    let html = '';
    
    const glassyColors = [
        'rgba(59, 130, 246, 0.1)',   // blue
        'rgba(16, 185, 129, 0.1)',   // green
        'rgba(245, 158, 11, 0.1)',   // orange
        'rgba(139, 92, 246, 0.1)',   // purple
        'rgba(239, 68, 68, 0.1)',    // red
        'rgba(20, 184, 166, 0.1)'    // teal
    ];
    
    examData.forEach((e, idx) => {
        const pctColor = e.percentage >= 100 ? 'var(--accent-green)' : e.percentage >= 50 ? 'var(--accent-orange)' : 'var(--accent-red)';
        const bgColor = glassyColors[idx % glassyColors.length];
        const cardStyle = `background: ${bgColor}; backdrop-filter: blur(12px); -webkit-backdrop-filter: blur(12px); border: 1px solid rgba(255,255,255,0.5);`;
        
        html += `<div class="exam-month-card" style="${cardStyle}">
            <div class="exam-month-header" style="background: transparent; border-bottom: 1px solid rgba(255,255,255,0.3); color: #1e293b;">${e.cycle}</div>
            <div class="exam-month-body">
                <div class="exam-stat"><span>Exam Papers</span><span class="exam-val">${e.papers}</span></div>
                <div class="exam-stat"><span>Results Released</span><span class="exam-val">${e.released}</span></div>
                <div class="exam-stat"><span>Completion</span><span class="exam-val" style="color:${pctColor}">${e.percentage}%</span></div>
            </div>
        </div>`;
    });
    document.getElementById('examCardsGrid').innerHTML = html;
}

// ===== INTERNSHIP BUILDERS =====
function buildInternshipSection(data) {
    const statsEl = document.getElementById('internshipStats');
    if (!data || !data.total) {
        statsEl.innerHTML = '<p style="color:var(--text-muted);font-size:13px;padding:16px;">No internship data</p>';
        return;
    }

    statsEl.innerHTML = `
        <div class="intern-stat-card intern-total">
            <div class="intern-stat-value">${data.total}</div>
            <div class="intern-stat-label">Total Internships</div>
        </div>
        <div class="intern-stat-card intern-audit">
            <div class="intern-stat-value">${data.audit.total}</div>
            <div class="intern-stat-label">Audit Firms</div>
            <div class="intern-stat-sub">Big 3: ${data.audit.big3} | Other: ${data.audit.other}</div>
        </div>
        <div class="intern-stat-card intern-nonaudit">
            <div class="intern-stat-value">${data.nonAudit}</div>
            <div class="intern-stat-label">Non-Audit</div>
        </div>
        <div class="intern-stat-card intern-ca">
            <div class="intern-stat-value">${data.caRegistered}</div>
            <div class="intern-stat-label">CA Registered</div>
            <div class="intern-stat-sub">${Math.round(data.caRegPct * 100)}%</div>
        </div>
    `;

    // Audit vs Non-Audit Pie
    chartInstances['internshipPie'] = new Chart(document.getElementById('internshipPieChart'), {
        type: 'doughnut',
        data: {
            labels: ['Audit', 'Non-Audit'],
            datasets: [{
                data: [data.audit.total, data.nonAudit],
                backgroundColor: [PPT.blue, PPT.orange],
                borderColor: '#ffffff', borderWidth: 3,
            }]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            cutout: '55%',
            plugins: {
                legend: { position: 'bottom', labels: { padding: 14, color: '#1a1d23' } },
                tooltip: tooltipConfig,
                datalabels: {
                    color: '#fff', font: { weight: '700', size: 14 },
                    formatter: (value, ctx) => {
                        const total = ctx.dataset.data.reduce((a, b) => a + b, 0);
                        return Math.round((value / total) * 100) + '%';
                    }
                }
            }
        }
    });

    // CA Registration Pie
    chartInstances['caRegistration'] = new Chart(document.getElementById('caRegistrationChart'), {
        type: 'doughnut',
        data: {
            labels: ['CA Registered', 'Not Registered'],
            datasets: [{
                data: [data.caRegistered, data.caNotRegistered],
                backgroundColor: [PPT.green, PPT.red],
                borderColor: '#ffffff', borderWidth: 3,
            }]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            cutout: '55%',
            plugins: {
                legend: { position: 'bottom', labels: { padding: 14, color: '#1a1d23' } },
                tooltip: tooltipConfig,
                datalabels: {
                    color: '#fff', font: { weight: '700', size: 14 },
                    formatter: (value, ctx) => {
                        const total = ctx.dataset.data.reduce((a, b) => a + b, 0);
                        return Math.round((value / total) * 100) + '%';
                    }
                }
            }
        }
    });
}

// ===== COUNSELLING BUILDER =====
function buildCounsellingChart(data) {
    if (!data || !data.length) return;
    const labels = data.map(d => d.month);
    const values = data.map(d => d.count);

    chartInstances['counselling'] = new Chart(document.getElementById('counsellingChart'), {
        type: 'bar',
        data: {
            labels,
            datasets: [{
                label: 'Sessions',
                data: values,
                backgroundColor: values.map((_, i) => PPT_COLORS[(i + 5) % PPT_COLORS.length]),
                borderColor: values.map((_, i) => PPT_COLORS[(i + 5) % PPT_COLORS.length]),
                borderWidth: 1, borderRadius: 6, borderSkipped: false,
                barThickness: 48,
            }]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            plugins: {
                legend: { display: false }, tooltip: tooltipConfig,
                datalabels: { anchor: 'end', align: 'top', color: '#1a1d23', font: { weight: '700', size: 14 } }
            },
            scales: {
                y: { beginAtZero: true, grid: { color: gridColor }, border: { display: false }, title: { display: true, text: 'No. of Sessions', color: '#5a6275' } },
                x: { grid: { display: false }, border: { display: false } }
            }
        }
    });
}

// ===== WELFARE SUMMARY =====
function buildWelfareSummary(internship, counselling) {
    const el = document.getElementById('welfareSummary');
    const totalCounselling = counselling ? counselling.reduce((s, c) => s + c.count, 0) : 0;

    let html = '<div class="welfare-cards">';
    html += `<div class="welfare-item">
        <div class="welfare-item-icon" style="background:rgba(68,114,196,0.1);color:${PPT.blue};">
            <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/></svg>
        </div>
        <div>
            <div class="welfare-item-value">${internship ? internship.total : 0}</div>
            <div class="welfare-item-label">Students in Internships</div>
        </div>
    </div>`;
    html += `<div class="welfare-item">
        <div class="welfare-item-icon" style="background:rgba(112,173,71,0.1);color:${PPT.green};">
            <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg>
        </div>
        <div>
            <div class="welfare-item-value">${internship ? internship.caRegistered : 0}</div>
            <div class="welfare-item-label">CA Registered Students</div>
        </div>
    </div>`;
    html += `<div class="welfare-item">
        <div class="welfare-item-icon" style="background:rgba(112,48,160,0.1);color:${PPT.purple};">
            <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M20.84 4.61a5.5 5.5 0 0 0-7.78 0L12 5.67l-1.06-1.06a5.5 5.5 0 0 0-7.78 7.78l1.06 1.06L12 21.23l7.78-7.78 1.06-1.06a5.5 5.5 0 0 0 0-7.78z"/></svg>
        </div>
        <div>
            <div class="welfare-item-value">${totalCounselling}</div>
            <div class="welfare-item-label">Total Counselling Sessions</div>
        </div>
    </div>`;

    if (counselling && counselling.length) {
        html += `<div class="welfare-item">
            <div class="welfare-item-icon" style="background:rgba(69,181,170,0.1);color:${PPT.teal};">
                <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 20V10"/><path d="M18 20V4"/><path d="M6 20v-4"/></svg>
            </div>
            <div>
                <div class="welfare-item-value">${Math.round(totalCounselling / counselling.length)}</div>
                <div class="welfare-item-label">Avg Sessions / Month</div>
            </div>
        </div>`;
    }
    html += '</div>';
    el.innerHTML = html;
}

// ===== NAVIGATION =====
function initNavigation() {
    const navLinks = document.querySelectorAll('.nav-link');
    const sections = document.querySelectorAll('.section');
    const observer = new IntersectionObserver(entries => {
        entries.forEach(entry => {
            if (entry.isIntersecting) {
                const id = entry.target.id;
                navLinks.forEach(link => link.classList.toggle('active', link.getAttribute('data-section') === id));
            }
        });
    }, { rootMargin: '-20% 0px -60% 0px' });
    sections.forEach(s => observer.observe(s));
    navLinks.forEach(link => link.addEventListener('click', e => {
        e.preventDefault();
        const t = document.querySelector(link.getAttribute('href'));
        if (t) t.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }));
}

function initAnimations() {
    // Only observe sections for scroll-spy and basic fade up
    const obs = new IntersectionObserver(entries => {
        entries.forEach(entry => { 
            if (entry.isIntersecting) { 
                entry.target.style.opacity = '1'; 
                entry.target.style.transform = 'translateY(0)'; 
            } 
        });
    }, { threshold: 0.05 });
    document.querySelectorAll('.section').forEach(el => {
        el.style.opacity = '0'; el.style.transform = 'translateY(20px)';
        el.style.transition = 'opacity 0.6s ease, transform 0.6s ease';
        obs.observe(el);
    });
}

// ===== HR & OPERATIONS BUILDERS & PARSERS =====
function parseHRTable(rows) {
    if (!rows || rows.length < 2) return { headers: [], data: [] };
    let headerIdx = -1;
    for (let i = 0; i < Math.min(6, rows.length); i++) {
        if (rows[i] && rows[i].some(c => c && typeof c === 'string' && !c.includes('Unnamed'))) {
            headerIdx = i; break;
        }
    }
    if (headerIdx < 0) return { headers: [], data: [] };
    
    const headers = rows[headerIdx].filter(h => h && typeof h === 'string' && !h.includes('Unnamed')).map(h => h.trim());
    const data = [];
    
    for (let i = headerIdx + 1; i < rows.length; i++) {
        const r = rows[i]; if (!r) continue;
        const rowData = [];
        let hasData = false;
        
        for (let j = 0; j < Math.min(r.length, rows[headerIdx].length); j++) {
            const head = rows[headerIdx][j];
            if (head && typeof head === 'string' && !head.includes('Unnamed')) {
                let val = r[j] != null ? r[j] : '';
                if (val instanceof Date) {
                    val = val.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
                }
                rowData.push(val);
                if (val !== '') hasData = true;
            }
        }
        if (hasData) data.push(rowData);
    }
    return { headers, data };
}

function buildHRSection(meetings, recruitment) {
    function buildTable(containerId, tblData) {
        const el = document.getElementById(containerId);
        if (!el) return;
        if (!tblData || !tblData.headers.length || !tblData.data.length) {
            el.innerHTML = '<p style="color:var(--text-mut);padding:16px;">No data available</p>';
            return;
        }
        let html = '<table class="data-table"><thead><tr>';
        tblData.headers.forEach(h => html += `<th>${h}</th>`);
        html += '</tr></thead><tbody>';
        tblData.data.forEach(row => {
            html += '<tr>';
            row.forEach((cell, i) => {
                let styling = i === 0 ? 'font-weight:700; color:#1e293b;' : '';
                html += `<td style="${styling}">${cell}</td>`;
            });
            html += '</tr>';
        });
        html += '</tbody></table>';
        el.innerHTML = html;
    }
    
    buildTable('meetingsTableWrap', meetings);
    buildTable('recruitmentTableWrap', recruitment);
}
