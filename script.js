let attendanceData = {};

// MODIFIED: Now accepts the full 'file' object to get its path and name
async function processExcelFile(arrayBuffer, file) { 
    try {
        const data = new Uint8Array(arrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        
        const courseName = (jsonData[2] && jsonData[2][2]) || '';
        console.log('Processing:', file.name, 'Course:', courseName);
        
        if (!courseName) {
            console.warn('No course name found in:', file.name);
            return { courseName: null, data: {} };
        }
        
        const studentData = {};
        for (let i = 6; i < jsonData.length; i++) {
            const row = jsonData[i];
            if (row && row[1]) {
                const rollNo = String(row[1]).trim().toUpperCase();
                studentData[rollNo] = {
                    name: row[2] || '',
                    courseName: courseName,
                    section: row[3] || '',
                    totalAbsent: Number(row[4]) || 0,
                    totalPresent: Number(row[5]) || 0,
                    sessions: row.slice(6).filter(x => x === 'P' || x === 'A'),
                    // NEW: Store the source file information for creating the link later
                    sourceFile: {
                        name: file.name,
                        path: file.path
                    }
                };
            }
        }
        return { courseName, data: studentData };
    } catch (error) {
        console.error('Error processing file:', file.name, error);
        return { courseName: null, data: {} };
    }
}

async function fetchGitHubDirectory() {
    // UPDATED with the data from our EDA
    const fileStructure = { "ACEL": [ "ACEL  Attendance - A.xlsx", "ACEL  Attendance - B.xlsx" ], "ADMA": [ "ADMA Attendance.xlsx" ], "AIMLB": [ "AIMLB Attendance - A.xlsx", "AIMLB Attendance - B.xlsx" ], "B2B": [ "B2B Attendance - A.xlsx", "B2B Attendance - B.xlsx" ], "BF": [ "BF Attendance - A.xlsx", "BF Attendance - B.xlsx" ], "BM": [ "BM Attendance - A.xlsx", "BM Attendance - B.xlsx" ], "BMOD": [ "BMOD Attendance - A.xlsx", "BMOD Attendance - B.xlsx" ], "CB": [ "CB Attendance - A.xlsx", "CB Attendance - B.xlsx" ], "CCIT": [ "CCIT Attendance.xlsx" ], "CEDA": [ "CEDA Attendance.xlsx" ], "CHW": [ "CHW Attendance.xlsx" ], "DIW": [ "DIW Attendance.xlsx" ], "DV": [ "DV Attendance - A.xlsx", "DV Attendance - B.xlsx" ], "EsIndia": [ "EsIndia Attendance - A.xlsx", "EsIndia Attendance - B.xlsx" ], "ExO": [ "ExO Attendance.xlsx" ], "FAMA": [ "FAMA Attendance - A.xlsx", "FAMA Attendance - B.xlsx" ], "FAuR": [ "FAuR Attendance.xlsx" ], "FBETDM": [ "FBETDM Attendance.xlsx" ], "FMac": [ "FMac Attendance.xlsx" ], "GS": [ "GS Attendance.xlsx" ], "IAPM": [ "IAPM Attendance - A.xlsx", "IAPM Attendance - B.xlsx" ], "IF": [ "IF Attendance - A.xlsx", "IF Attendance - B.xlsx" ], "ITA": [ "ITA Attendance.xlsx" ], "IoTB": [ "IoTB Attendance.xlsx" ], "LMPO": [ "LMPO Attendance.xlsx" ], "LMW": [ "LMW Attendance - A.xlsx", "LMW Attendance - B.xlsx" ], "LSS": [ "LSS Attendance.xlsx" ], "M&A": [ "M&A Attendance.xlsx" ], "MACR": [ "MACR Attendance.xlsx" ], "MLVW": [ "MLVW Attendance.xlsx" ], "MSN": [ "MSN Attendance - A.xlsx", "MSN Attendance - B.xlsx" ], "OCMC": [ "OCMC Attendance.xlsx" ], "OFDG": [ "OFDG Attendance - A.xlsx", "OFDG Attendance - B.xlsx" ], "PM": [ "PM Attendance - A.xlsx", "PM Attendance - B.xlsx" ], "People": [ "People Attendance.xlsx" ], "Pricing": [ "Pricing Attendance - A.xlsx", "Pricing Attendance - B.xlsx" ], "Python": [ "Python Attendance - A.xlsx", "Python Attendance - B.xlsx" ], "SA": [ "SA Attendance.xlsx" ], "SAR": [ "SAR Attendance.xlsx" ], "SM": [ "SM Attendance - A.xlsx", "SM Attendance - B.xlsx" ], "SS": [ "SS Attendance - A.xlsx", "SS Attendance - B.xlsx", "SS Attendance - C.xlsx", "SS Attendance - D.xlsx", "SS Attendance - E.xlsx", "SS Attendance - F.xlsx", "SS Attendance - G.xlsx", "SS Attendance - H.xlsx", "SS Attendance - I.xlsx", "SS Attendance - J.xlsx" ], "TOC": [ "TOC Attendance.xlsx" ] };
    
    // UPDATED: New username and repo name
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
        console.log(`Found total ${files.length} Excel files`);
        attendanceData = {};
        let processed = 0;

        await Promise.all(files.map(async (file) => {
            try {
                const response = await fetch(file.download_url);
                if (!response.ok) {
                    console.error(`Failed to fetch ${file.download_url}: ${response.statusText}`);
                    return;
                };
                
                const arrayBuffer = await response.arrayBuffer();
                // MODIFIED: Pass the whole 'file' object
                const { courseName, data: studentData } = await processExcelFile(arrayBuffer, file);
                
                if (courseName) {
                    if (!attendanceData[courseName]) {
                        attendanceData[courseName] = {};
                    }
                    Object.assign(attendanceData[courseName], studentData);
                    processed++;
                    console.log(`Successfully processed (${processed}/${files.length}): ${courseName} from ${file.name}`);
                }
            } catch (err) {
                console.error(`Error processing ${file.name}:`, err);
            }
        }));

        console.log(`Processing complete. Success: ${processed}`);
        
    } catch (error) {
        console.error('Error in loadAttendanceData:', error);
    }
}

function checkAttendance() {
    const rollInput = document.getElementById('rollNumber').value.trim().toUpperCase();
    const resultDiv = document.getElementById('result');
    const studentInfoDiv = document.getElementById('studentInfo');
    
    if (!rollInput) {
        resultDiv.innerHTML = '';
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
            const subjectData = courseData[rollInput];
            studentSubjects.push(subjectData); // The whole object, including sourceFile
            if (!studentName && subjectData.name) {
                studentName = subjectData.name;
            }
            totalOverallClasses += subjectData.totalPresent + subjectData.totalAbsent;
            totalOverallPresent += subjectData.totalPresent;
        }
    });

    if (studentSubjects.length === 0) {
        resultDiv.innerHTML = '<p class="text-center text-red-500 col-span-full">Roll number not found. Please check the roll number and try again.</p>';
        studentInfoDiv.classList.add('hidden');
        return;
    }
    
    // ... (The studentInfoDiv generation code remains the same) ...
    const overallPercentage = totalOverallClasses ? 
        ((totalOverallPresent / totalOverallClasses) * 100).toFixed(1) : '0';

    studentInfoDiv.classList.remove('hidden');
    studentInfoDiv.innerHTML = `
        <div class="student-info-card">
            <div class="info-header">
                <div class="student-details">
                    <h2 class="student-name">${studentName}</h2>
                    <p class="roll-number">${rollInput}</p>
                </div>
                <div class="overall-stats">
                    <div class="stat">
                        <span class="stat-number">${totalOverallClasses}</span>
                        <span class="stat-label">Total Classes</span>
                    </div>
                    <div class="stat">
                        <span class="stat-number">${totalOverallPresent}</span>
                        <span class="stat-label">Present</span>
                    </div>
                    <div class="stat">
                        <span class="stat-number">${totalOverallClasses - totalOverallPresent}</span>
                        <span class="stat-label">Absent</span>
                    </div>
                </div>
            </div>
            <div class="progress-container">
                <div class="progress-bar">
                    <div class="progress-fill ${overallPercentage >= 75 ? 'high' : overallPercentage >= 60 ? 'medium' : 'low'}"
                         style="width: ${overallPercentage}%"></div>
                </div>
                <span class="progress-label">Overall: ${overallPercentage}%</span>
            </div>
        </div>`;


    // MODIFIED: Subject Card generation now includes the "View Sheet" button
    resultDiv.innerHTML = studentSubjects.map((subject, index) => {
        const totalClasses = subject.totalPresent + subject.totalAbsent;
        const attendancePercent = totalClasses ? ((subject.totalPresent / totalClasses) * 100).toFixed(2) : 0;
        const containerID = `progress-${index}`;

        // Construct the GitHub raw file URL
        const githubFileUrl = `https://raw.githubusercontent.com/maddydev7/attendance/main/attendance-files/${subject.sourceFile.path}/${subject.sourceFile.name}`;
        
        // Construct the final Office Live viewer URL
        const officeViewerUrl = `https://view.officeapps.live.com/op/view.aspx?src=${encodeURIComponent(githubFileUrl)}`;

        return `
            <div class="subject-card">
                <h3 class="subject-name">${subject.courseName}</h3>
                <div id="${containerID}" class="progress-circle"></div>
                <div class="stats-container">
                    <div class="stat-item">
                        <span class="stat-value">${totalClasses}</span>
                        <span class="stat-label">Total</span>
                    </div>
                    <div class="stat-item">
                        <span class="stat-value text-green-600">${subject.totalPresent}</span>
                        <span class="stat-label">Present</span>
                    </div>
                    <div class="stat-item">
                        <span class="stat-value text-red-600">${subject.totalAbsent}</span>
                        <span class="stat-label">Absent</span>
                    </div>
                    <div class="stat-item">
                        <span class="stat-value">${attendancePercent}%</span>
                        <span class="stat-label">Attend.</span>
                    </div>
                </div>
                <!-- NEW: View Sheet Button -->
                <a href="${officeViewerUrl}" target="_blank" class="view-sheet-btn">
                    View Sheet
                    <svg class="w-4 h-4 ml-1.5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M10 6H6a2 2 0 00-2 2v10a2 2 0 002 2h10a2 2 0 002-2v-4M14 4h6m0 0v6m0-6L10 14"></path></svg>
                </a>
            </div>
        `;
    }).join('');

    // Initialize progress circles
    studentSubjects.forEach((subject, index) => {
        const totalClasses = subject.totalPresent + subject.totalAbsent;
        const attendancePercent = totalClasses ? 
            ((subject.totalPresent / totalClasses) * 100) : 0;
        createProgressCircle(`progress-${index}`, attendancePercent);
    });
}

// ... (createProgressCircle and DOMContentLoaded functions remain the same) ...
document.addEventListener('DOMContentLoaded', () => {
    const checkBtn = document.getElementById('checkBtn');
    const rollInput = document.getElementById('rollNumber');

    loadAttendanceData();

    if (checkBtn) checkBtn.addEventListener('click', checkAttendance);
    if (rollInput) {
        rollInput.addEventListener('keypress', (e) => {
            if (e.key === 'Enter') {
                checkAttendance();
            }
        });
    }
});