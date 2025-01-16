// Comment out showLoading function

function showLoading() {
    const resultDiv = document.getElementById('result');
    resultDiv.innerHTML = `
        <div class="fixed inset-0 flex items-center justify-center bg-white bg-opacity-90 z-50">
            <div class="w-96 text-center p-8">
                <div class="mb-4 text-lg font-semibold text-gray-700">Loading attendance data...</div>
                <div class="w-full bg-gray-200 rounded-full h-2.5">
                    <div class="bg-blue-600 h-2.5 rounded-full transition-all duration-300" id="progressBar" style="width: 0%"></div>
                </div>
                <p class="mt-2 text-sm text-gray-600" id="loadingStatus"></p>
            </div>
        </div>
    `;
}

async function processExcelFile(arrayBuffer, fileName) {
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
                const rollNo = String(row[1]).trim();
                studentData[rollNo] = {
                    name: row[2] || '',
                    courseName: courseName,
                    section: row[3] || '',
                    totalAbsent: Number(row[4]) || 0,
                    totalPresent: Number(row[5]) || 0,
                    sessions: row.slice(6).filter(x => x === 'P' || x === 'A')
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
    const fileStructure = {
        'DT': ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'].map(section => `DT (${section}) Attendance Sheet.xlsx`),
        'FA-II': ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'].map(section => `FA-II (${section}) Attendance Sheet.xlsx`),
        'FIM': [
            'FIM-D (ABCD) Attendance Sheet.xlsx',
            'FIM-D (EFGH) Attendance Sheet.xlsx',
            'FIM-U (ABCD) Attendance Sheet.xlsx',
            'FIM-U (EFGH) Attendance Sheet.xlsx'
        ],
        'HRM': ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'].map(section => `HRM (${section}) Attendance Sheet.xlsx`),
        'LAB': ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'].map(section => `LAB (${section}) Attendance Sheet.xlsx`),
        'MR': [
            'MR-A (ABCD) Attendance Sheet.xlsx',
            'MR-A (EFGH) Attendance Sheet.xlsx',
            'MR-B (ABCD) Attendance Sheet.xlsx',
            'MR-B (EFGH) Attendance Sheet.xlsx',
            'MR-S (ABCD) Attendance Sheet.xlsx',
            'MR-S (EFGH) Attendance Sheet.xlsx'
        ],
        'SCM': [
            'SCM-H Attendance Sheet.xlsx',
            'SCM-R (ABCD) Attendance.xlsx',
            'SCM-R (EFGH) Attendance.xlsx'
        ],
        'SDM': [
            'SDM-A (ABCD) Attendance Sheet.xlsx',
            'SDM-A (EFGH) Attendance Sheet.xlsx',
            'SDM-M (ABCD) Attendance Sheet.xlsx',
            'SDM-M (EFGH) Attendance Sheet.xlsx'
        ],
        'SIP': ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'].map(section => `SIP (${section}) Attendance Sheet.xlsx`),
        'SM-II': ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'].map(section => `SM-II (${section}) Attendance Sheet.xlsx`)
    };

    const baseUrl = 'https://raw.githubusercontent.com/iimindore/attendance-system/main/attendance-files';
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
                console.log(`Processing (${processed + 1}/${files.length}): ${file.path}/${file.name}`);
                const response = await fetch(file.download_url);
                if (!response.ok) return;
                
                const arrayBuffer = await response.arrayBuffer();
                const { courseName, data: studentData } = await processExcelFile(arrayBuffer, file.name);
                
                if (courseName) {
                    if (!attendanceData[courseName]) {
                        attendanceData[courseName] = {};
                    }
                    Object.assign(attendanceData[courseName], studentData);
                    processed++;
                    console.log(`Successfully processed: ${courseName}`);
                }
            } catch (err) {
                console.error(`Error processing ${file.name}:`, err);
            }
        }));

        console.log(`Processing complete. Success: ${processed}`);
        console.log('Courses found:', Object.keys(attendanceData));
        localStorage.setItem('attendanceCache', JSON.stringify(attendanceData));  // Add this line
        
    } catch (error) {
        console.error('Error in loadAttendanceData:', error);
    }
}

function createProgressCircle(containerId, percentage) {
    const canvas = document.createElement('canvas');
    const ctx = canvas.getContext('2d');
    const container = document.getElementById(containerId);
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
                borderRadius: 5
            }]
        },
        options: {
            cutout: '85%',
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false }
            }
        }
    });
}

function checkAttendance() {
    const rollInput = document.getElementById('rollNumber').value.trim();
    const resultDiv = document.getElementById('result');
    const studentInfoDiv = document.getElementById('studentInfo');
    
    if (!attendanceData || Object.keys(attendanceData).length === 0) {
        resultDiv.innerHTML = '<p class="text-center text-red-500">No attendance data available</p>';
        return;
    }

    let studentSubjects = [];
    let studentName = '';
    let totalOverallClasses = 0;
    let totalOverallPresent = 0;

    Object.entries(attendanceData).forEach(([courseName, courseData]) => {
        if (courseData[rollInput]) {
            studentSubjects.push({
                courseName: courseName,
                ...courseData[rollInput]
            });
            if (!studentName && courseData[rollInput].name) {
                studentName = courseData[rollInput].name;
            }
            totalOverallClasses += courseData[rollInput].totalPresent + courseData[rollInput].totalAbsent;
            totalOverallPresent += courseData[rollInput].totalPresent;
        }
    });

    console.log('Found subjects:', studentSubjects.length);

    if (studentSubjects.length === 0) {
        resultDiv.innerHTML = '<p class="text-center text-red-500">Roll number not found</p>';
        studentInfoDiv.classList.add('hidden');
        return;
    }

    // Student Info Bar
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

    // Subject Cards
    resultDiv.innerHTML = studentSubjects.map((subject, index) => {
        const totalClasses = subject.totalPresent + subject.totalAbsent;
        const attendancePercent = totalClasses ? ((subject.totalPresent / totalClasses) * 100).toFixed(2) : 0;
        const containerID = `progress-${index}`;

        return `
            <div class="subject-card">
                <h3 class="subject-name">${subject.courseName}</h3>
                <div id="${containerID}" class="progress-circle"></div>
                <div class="section-badge">${subject.section}</div>
                <div class="stats-container">
                    <div class="stat-item">
                        <span class="stat-value">${totalClasses}</span>
                        <span class="stat-label">Total Classes</span>
                    </div>
                    <div class="stat-item">
                        <span class="stat-value">${subject.totalPresent}</span>
                        <span class="stat-label">Present</span>
                    </div>
                    <div class="stat-item">
                        <span class="stat-value">${subject.totalAbsent}</span>
                        <span class="stat-label">Absent</span>
                    </div>
                    <div class="stat-item">
                        <span class="stat-value">${attendancePercent}%</span>
                        <span class="stat-label">Attendance</span>
                    </div>
                </div>
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