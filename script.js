let attendanceData = {};

// MODIFIED: Added 'downloadUrl' parameter
async function processExcelFile(arrayBuffer, fileName, downloadUrl) {
    try {
        const data = new Uint8Array(arrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        
        const courseName = (jsonData[2] && jsonData[2][2]) || '';
        console.log('Processing:', fileName, 'Course:', courseName);
        
        if (!courseName) {
            console.warn('No course name found in:', fileName);
            return { courseName: null, data: {} };
        }
        
        const studentData = {};
        for (let i = 6; i < jsonData.length; i++) {
            const row = jsonData[i];
            if (row && row[1]) {
                const rollNo = String(row[1]).trim().toUpperCase(); // Standardize roll numbers
                studentData[rollNo] = {
                    name: row[2] || '',
                    courseName: courseName,
                    section: row[3] || '',
                    totalAbsent: Number(row[4]) || 0,
                    totalPresent: Number(row[5]) || 0,
                    sessions: row.slice(6).filter(x => x === 'P' || x === 'A'),
                    // NEW: Store the source file URL
                    sourceUrl: downloadUrl 
                };
            }
        }
        return { courseName, data: studentData };
    } catch (error) {
        console.error('Error processing file:', fileName, error);
        return { courseName: null, data: {} };
    }
}

async function fetchGitHubDirectory() {
    // This is the auto-generated list of files
    const fileStructure = { "ACEL": [ "ACEL  Attendance - A.xlsx", "ACEL  Attendance - B.xlsx" ], "ADMA": [ "ADMA Attendance.xlsx" ], "AIMLB": [ "AIMLB Attendance - A.xlsx", "AIMLB Attendance - B.xlsx" ], "B2B": [ "B2B Attendance - A.xlsx", "B2B Attendance - B.xlsx" ], "BF": [ "BF Attendance - A.xlsx", "BF Attendance - B.xlsx" ], "BM": [ "BM Attendance - A.xlsx", "BM Attendance - B.xlsx" ], "BMOD": [ "BMOD Attendance - A.xlsx", "BMOD Attendance - B.xlsx" ], "CB": [ "CB Attendance - A.xlsx", "CB Attendance - B.xlsx" ], "CCIT": [ "CCIT Attendance.xlsx" ], "CEDA": [ "CEDA Attendance.xlsx" ], "CHW": [ "CHW Attendance.xlsx" ], "DIW": [ "DIW Attendance.xlsx" ], "DV": [ "DV Attendance - A.xlsx", "DV Attendance - B.xlsx" ], "EsIndia": [ "EsIndia Attendance - A.xlsx", "EsIndia Attendance - B.xlsx" ], "ExO": [ "ExO Attendance.xlsx" ], "FAMA": [ "FAMA Attendance - A.xlsx", "FAMA Attendance - B.xlsx" ], "FAuR": [ "FAuR Attendance.xlsx" ], "FBETDM": [ "FBETDM Attendance.xlsx" ], "FMac": [ "FMac Attendance.xlsx" ], "GS": [ "GS Attendance.xlsx" ], "IAPM": [ "IAPM Attendance - A.xlsx", "IAPM Attendance - B.xlsx" ], "IF": [ "IF Attendance - A.xlsx", "IF Attendance - B.xlsx" ], "ITA": [ "ITA Attendance.xlsx" ], "IoTB": [ "IoTB Attendance.xlsx" ], "LMPO": [ "LMPO Attendance.xlsx" ], "LMW": [ "LMW Attendance - A.xlsx", "LMW Attendance - B.xlsx" ], "LSS": [ "LSS Attendance.xlsx" ], "M&A": [ "M&A Attendance.xlsx" ], "MACR": [ "MACR Attendance.xlsx" ], "MLVW": [ "MLVW Attendance.xlsx" ], "MSN": [ "MSN Attendance - A.xlsx", "MSN Attendance - B.xlsx" ], "OCMC": [ "OCMC Attendance.xlsx" ], "OFDG": [ "OFDG Attendance - A.xlsx", "OFDG Attendance - B.xlsx" ], "PM": [ "PM Attendance - A.xlsx", "PM Attendance - B.xlsx" ], "People": [ "People Attendance.xlsx" ], "Pricing": [ "Pricing Attendance - A.xlsx", "Pricing Attendance - B.xlsx" ], "Python": [ "Python Attendance - A.xlsx", "Python Attendance - B.xlsx" ], "SA": [ "SA Attendance.xlsx" ], "SAR": [ "SAR Attendance.xlsx" ], "SM": [ "SM Attendance - A.xlsx", "SM Attendance - B.xlsx" ], "SS": [ "SS Attendance - A.xlsx", "SS Attendance - B.xlsx", "SS Attendance - C.xlsx", "SS Attendance - D.xlsx", "SS Attendance - E.xlsx", "SS Attendance - F.xlsx", "SS Attendance - G.xlsx", "SS Attendance - H.xlsx", "SS Attendance - I.xlsx", "SS Attendance - J.xlsx" ], "TOC": [ "TOC Attendance.xlsx" ] };
    
    // UPDATED: Changed username to 'maddydev7' and repo name to 'attendance'
    const baseUrl = 'https://raw.githubusercontent.com/maddydev7/attendance/main/attendance-files';
    let allFiles = [];

    for (const [subject, files] of Object.entries(fileStructure)) {
        for (const file of files) {
            allFiles.push({
                name: file,
                path: subject,
                download_url: `${baseUrl}/${subject}/${encodeURIComponent(file)}`
            });
        }
    }

    console.log(`Total files to process: ${allFiles.length}`);
    return allFiles;
}

async function loadAttendanceData() {
    try {
        console.log('Starting to fetch files...');
        const files = await fetchGitHubDirectory();
        attendanceData = {};
        let processed = 0;

        await Promise.all(files.map(async (file) => {
            try {
                const response = await fetch(file.download_url);
                if (!response.ok) {
                    console.error(`Failed to fetch ${file.download_url}: ${response.statusText}`);
                    return;
                }
                const arrayBuffer = await response.arrayBuffer();
                // MODIFIED: Pass the download_url to the processor
                const { courseName, data: studentData } = await processExcelFile(arrayBuffer, file.name, file.download_url);
                
                if (courseName) {
                    if (!attendanceData[courseName]) {
                        attendanceData[courseName] = {};
                    }
                    Object.assign(attendanceData[courseName], studentData);
                    processed++;
                }
            } catch (err) {
                console.error(`Error processing ${file.name}:`, err);
            }
        }));

        console.log(`Processing complete. Success: ${processed}/${files.length}`);
        console.log('Courses found:', Object.keys(attendanceData));
        
    } catch (error) {
        console.error('Error in loadAttendanceData:', error);
    }
}

function createProgressCircle(containerId, percentage) {
    const canvas = document.createElement('canvas');
    const ctx = canvas.getContext('2d');
    const container = document.getElementById(containerId);
    if (!container) return;
    container.innerHTML = '';
    container.appendChild(canvas);

    function createGradient(percentage) {
        const gradient = ctx.createLinearGradient(0, 0, 150, 0);
        if (percentage >= 75) {
            gradient.addColorStop(0, '#22c55e');
            gradient.addColorStop(1, '#16a34a');
        } else if (percentage >= 60) {
            gradient.addColorStop(0, '#fbbf24');
            gradient.addColorStop(1, '#d97706');
        } else {
            gradient.addColorStop(0, '#ef4444');
            gradient.addColorStop(1, '#dc2626');
        }
        return gradient;
    }

    new Chart(ctx, {
        type: 'doughnut',
        data: {
            datasets: [{
                data: [percentage, 100 - percentage],
                backgroundColor: [createGradient(percentage), '#E5E7EB'],
                borderWidth: 0,
                borderRadius: { outerStart: 5, outerEnd: 5, innerStart: 5, innerEnd: 5 }
            }]
        },
        options: {
            cutout: '85%',
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { display: false }, tooltip: { enabled: false } }
        }
    });
}

function checkAttendance() {
    const rollInput = document.getElementById('rollNumber').value.trim().toUpperCase();
    const resultDiv = document.getElementById('result');
    const studentInfoDiv = document.getElementById('studentInfo');
    
    if (!rollInput) {
        resultDiv.innerHTML = '<p class="text-center text-gray-500 col-span-full">Please enter a roll number.</p>';
        studentInfoDiv.classList.add('hidden');
        return;
    }
    
    if (!attendanceData || Object.keys(attendanceData).length === 0) {
        resultDiv.innerHTML = '<p class="text-center text-red-500 col-span-full">Attendance data is still loading or failed to load. Please try again in a moment.</p>';
        return;
    }

    let studentSubjects = [];
    let studentName = '';
    let totalOverallClasses = 0;
    let totalOverallPresent = 0;

    Object.entries(attendanceData).forEach(([courseName, courseData]) => {
        if (courseData[rollInput]) {
            const subjectInfo = { courseName, ...courseData[rollInput] };
            studentSubjects.push(subjectInfo);
            if (!studentName && subjectInfo.name) {
                studentName = subjectInfo.name;
            }
            totalOverallClasses += subjectInfo.totalPresent + subjectInfo.totalAbsent;
            totalOverallPresent += subjectInfo.totalPresent;
        }
    });

    if (studentSubjects.length === 0) {
        resultDiv.innerHTML = '<p class="text-center text-red-500 col-span-full">Roll number not found. Please check the roll number and try again.</p>';
        studentInfoDiv.classList.add('hidden');
        return;
    }

    const overallPercentage = totalOverallClasses ? ((totalOverallPresent / totalOverallClasses) * 100).toFixed(1) : '0';

    studentInfoDiv.classList.remove('hidden');
    studentInfoDiv.innerHTML = `
        <div class="student-info-card">
            <div class="info-header">
                <div class="student-details">
                    <h2 class="student-name">${studentName}</h2>
                    <p class="roll-number">${rollInput}</p>
                </div>
                <div class="overall-stats">
                    <div class="stat"><span class="stat-number">${totalOverallClasses}</span><span class="stat-label">Total Classes</span></div>
                    <div class="stat"><span class="stat-number">${totalOverallPresent}</span><span class="stat-label">Present</span></div>
                    <div class="stat"><span class="stat-number">${totalOverallClasses - totalOverallPresent}</span><span class="stat-label">Absent</span></div>
                </div>
            </div>
            <div class="progress-container">
                <div class="progress-bar"><div class="progress-fill ${overallPercentage >= 75 ? 'high' : overallPercentage >= 60 ? 'medium' : 'low'}" style="width: ${overallPercentage}%"></div></div>
                <span class="progress-label">Overall: ${overallPercentage}%</span>
            </div>
        </div>`;

    resultDiv.innerHTML = studentSubjects.map((subject, index) => {
        const totalClasses = subject.totalPresent + subject.totalAbsent;
        const attendancePercent = totalClasses ? ((subject.totalPresent / totalClasses) * 100).toFixed(2) : 0;
        const containerID = `progress-${index}`;

        // NEW: Construct the Office Live viewer link
        const officeViewerBase = 'https://view.officeapps.live.com/op/view.aspx?src=';
        const encodedUrl = encodeURIComponent(subject.sourceUrl);
        const viewSheetLink = officeViewerBase + encodedUrl;

        // MODIFIED: Added the 'View Sheet' link to the card
        return `
            <div class="subject-card">
                <h3 class="subject-name">${subject.courseName}</h3>
                <div id="${containerID}" class="progress-circle"></div>
                <div class="section-badge">${subject.section}</div>
                <div class="stats-container">
                    <div class="stat-item"><span class="stat-value">${totalClasses}</span><span class="stat-label">Total</span></div>
                    <div class="stat-item"><span class="stat-value text-green-600">${subject.totalPresent}</span><span class="stat-label">Present</span></div>
                    <div class="stat-item"><span class="stat-value text-red-600">${subject.totalAbsent}</span><span class="stat-label">Absent</span></div>
                    <div class="stat-item"><span class="stat-value">${attendancePercent}%</span><span class="stat-label">Attend.</span></div>
                </div>
                <div class="mt-auto pt-4">
                    <a href="${viewSheetLink}" target="_blank" class="view-sheet-link">
                        View Sheet
                        <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 20 20" fill="currentColor" class="w-5 h-5 ml-1"><path d="M12.232 4.232a2.5 2.5 0 013.536 3.536l-1.225 1.224a.75.75 0 001.061 1.06l1.224-1.224a4 4 0 00-5.656-5.656l-3 3a4 4 0 00.225 5.865.75.75 0 00.977-1.138 2.5 2.5 0 01-.142-3.665l3-3z"></path><path d="M8.603 3.799a4.49 4.49 0 014.49 4.49v.03a4.49 4.49 0 01-4.49 4.49h-1.982a4.5 4.5 0 01-4.49-4.49v-1.982a4.5 4.5 0 014.49-4.49h1.982zm-1.982 1.5a3 3 0 00-3 3v1.982a3 3 0 003 3h1.982a3 3 0 003-3v-1.982a3 3 0 00-3-3h-1.982z"></path></svg>
                    </a>
                </div>
            </div>
        `;
    }).join('');

    studentSubjects.forEach((subject, index) => {
        const totalClasses = subject.totalPresent + subject.totalAbsent;
        const attendancePercent = totalClasses ? ((subject.totalPresent / totalClasses) * 100) : 0;
        createProgressCircle(`progress-${index}`, attendancePercent);
    });
}

document.addEventListener('DOMContentLoaded', () => {
    const checkBtn = document.getElementById('checkBtn');
    const rollInput = document.getElementById('rollNumber');
    loadAttendanceData();
    if (checkBtn) checkBtn.addEventListener('click', checkAttendance);
    if (rollInput) rollInput.addEventListener('keypress', e => e.key === 'Enter' && checkAttendance());
});