let attendanceData = {};
let processedFiles = 0;
let totalFiles = 0;

async function processExcelFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const worksheet = workbook.Sheets[workbook.SheetNames[0]];
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                
                const courseName = (jsonData[2] && jsonData[2][2]) || '';
                if (!courseName) {
                    console.warn(`No course name found in file: ${file.name}`);
                    resolve({ courseName: null, data: {} });
                    return;
                }
                
                console.log(`Processing ${file.name} for course ${courseName}`);
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
                
                resolve({ courseName, data: studentData });
            } catch (error) {
                console.error(`Error processing file ${file.name}:`, error);
                resolve({ courseName: null, data: {} });
            }
        };
        reader.onerror = () => reject(new Error(`Failed to read file ${file.name}`));
        reader.readAsArrayBuffer(file);
    });
}

async function fetchGitHubDirectory() {
    const subjects = ['DT', 'FA-II', 'FIM', 'HRM', 'LAB', 'MR', 'SCM', 'SDM', 'SIP', 'SM-II'];
    const baseUrl = 'https://raw.githubusercontent.com/iimindore/attendance-system/main/attendance-files';
    let allFiles = [];
    
    try {
        for (const subject of subjects) {
            // Each subject has its sections in Excel files
            const fileUrl = `${baseUrl}/${subject}/attendance.xlsx`;
            try {
                const response = await fetch(fileUrl);
                if (response.ok) {
                    allFiles.push({
                        name: `${subject}.xlsx`,
                        path: subject,
                        download_url: fileUrl
                    });
                }
            } catch (err) {
                console.error(`Error fetching ${subject}:`, err);
            }
        }
        
        console.log(`Found ${allFiles.length} Excel files`);
        return allFiles;
    } catch (error) {
        console.error('Error fetching files:', error);
        return [];
    }
}

async function loadAttendanceData() {
    try {
        console.log('Starting to fetch files...');
        const files = await fetchGitHubDirectory();
        console.log(`Found total ${files.length} Excel files`);
        attendanceData = {};
        
        for (const file of files) {
            try {
                console.log(`Processing: ${file.path}/${file.name}`);
                const response = await fetch(file.download_url);
                if (!response.ok) {
                    console.error(`Failed to fetch ${file.name}`);
                    continue;
                }
                
                const arrayBuffer = await response.arrayBuffer();
                const data = new Uint8Array(arrayBuffer);
                
                const { courseName, data: studentData } = await processExcelFile(data);
                if (courseName) {
                    attendanceData[courseName] = {
                        ...attendanceData[courseName],
                        ...studentData
                    };
                }
                console.log(`Processed ${courseName}, found ${Object.keys(studentData).length} students`);
            } catch (err) {
                console.error(`Error processing ${file.name}:`, err);
            }
        }
        
        console.log('Final processed courses:', Object.keys(attendanceData));
        console.log('Total courses:', Object.keys(attendanceData).length);
        localStorage.setItem('attendanceData', JSON.stringify(attendanceData));
        
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
    
    console.log('Checking roll number:', rollInput);
    
    const storedData = localStorage.getItem('attendanceData');
    if (!storedData) {
        console.log('No stored data found');
        resultDiv.innerHTML = '<p class="text-center text-red-500">No attendance data available</p>';
        return;
    }

    const data = JSON.parse(storedData);
    console.log('Available courses:', Object.keys(data));
    
    let studentSubjects = [];
    let studentName = '';
    let totalOverallClasses = 0;
    let totalOverallPresent = 0;

    Object.entries(data).forEach(([courseName, courseData]) => {
        console.log('Checking course:', courseName);
        if (courseData[rollInput]) {
            console.log('Found student in course:', courseName);
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

    // Load attendance data when page loads
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