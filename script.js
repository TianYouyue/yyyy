// 全局变量
let scoreData = [];          // 成绩数据
let configData = {};         // 配置数据
let selectedStudents = [];   // 筛选后的学生列表
let currentPreviewStudent = null; // 当前预览的学生

// 初始化
document.addEventListener('DOMContentLoaded', function() {
    // 绑定事件
    document.getElementById('scoreFile').addEventListener('change', handleScoreFile);
    document.getElementById('configFile').addEventListener('change', handleConfigFile);
    document.getElementById('previewBtn').addEventListener('click', previewScoreCard);
    document.getElementById('generateBtn').addEventListener('click', batchGenerate);
    document.getElementById('studentSelect').addEventListener('change', changePreviewStudent);
    document.getElementById('downloadSingleBtn').addEventListener('click', downloadSingleScoreCard);
    document.getElementById('paperSize').addEventListener('change', updatePreview);
    document.getElementById('paperOrientation').addEventListener('change', updatePreview);
    document.getElementById('gradeSelect').addEventListener('change', updateClassOptions);
    document.getElementById('cancelBtn').addEventListener('click', cancelBatch);
});

// 处理成绩文件上传
function handleScoreFile(event) {
    const file = event.target.files[0];
    if (!file) return;
    
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {type: 'array'});
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            scoreData = XLSX.utils.sheet_to_json(firstSheet);
            
            showStatus('成绩文件加载成功！共' + scoreData.length + '条记录', 'success');
            checkFilesLoaded();
        } catch (error) {
            showStatus('文件解析失败：' + error.message, 'error');
        }
    };
    reader.readAsArrayBuffer(file);
}

// 处理配置文件上传
function handleConfigFile(event) {
    const file = event.target.files[0];
    if (!file) return;
    
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {type: 'array'});
            
            // 读取考次表
            const examSheet = workbook.Sheets['考次'];
            configData.exams = XLSX.utils.sheet_to_json(examSheet, {header: 1})
                .flat()
                .filter(exam => exam && exam !== '考次');
            
            // 读取班级表
            const classSheet = workbook.Sheets['班级'];
            const classRows = XLSX.utils.sheet_to_json(classSheet, {header: 1});
            configData.classes = {};
            classRows.slice(1).forEach(row => {
                if (row[0] && row[1]) {
                    if (!configData.classes[row[0]]) {
                        configData.classes[row[0]] = [];
                    }
                    configData.classes[row[0]].push(row[1]);
                }
            });
            
            // 更新UI
            updateConfigUI();
            showStatus('配置文件加载成功！', 'success');
            checkFilesLoaded();
        } catch (error) {
            showStatus('配置文件解析失败：' + error.message, 'error');
        }
    };
    reader.readAsArrayBuffer(file);
}

// 检查文件是否都已加载
function checkFilesLoaded() {
    if (scoreData.length > 0 && configData.exams) {
        document.querySelector('.config-section').style.display = 'block';
        document.querySelector('.preview-section').style.display = 'block';
    }
}

// 更新配置UI
function updateConfigUI() {
    // 更新年级选择
    const gradeSelect = document.getElementById('gradeSelect');
    gradeSelect.innerHTML = '<option value="all">全部年级</option>';
    Object.keys(configData.classes).forEach(grade => {
        const option = document.createElement('option');
        option.value = grade;
        option.textContent = grade;
        gradeSelect.appendChild(option);
    });
    
    // 更新考次选择
    const examContainer = document.getElementById('examCheckboxes');
    examContainer.innerHTML = '';
    configData.exams.forEach(exam => {
        const div = document.createElement('div');
        div.className = 'checkbox-item';
        div.innerHTML = `
            <input type="checkbox" id="exam_${exam}" value="${exam}" checked>
            <label for="exam_${exam}">${exam}</label>
        `;
        examContainer.appendChild(div);
    });
    
    // 更新科目选择（从数据中提取所有科目）
    updateSubjectOptions();
}

// 从数据中提取所有科目
function updateSubjectOptions() {
    if (scoreData.length === 0) return;
    
    // 排除前6列（考次、学号、姓名、年级、班级、总分）
    const firstRow = scoreData[0];
    const subjectKeys = Object.keys(firstRow).slice(5);
    
    const subjectContainer = document.getElementById('subjectCheckboxes');
    subjectContainer.innerHTML = '';
    
    subjectKeys.forEach(subject => {
        if (subject.trim()) {
            const div = document.createElement('div');
            div.className = 'checkbox-item';
            div.innerHTML = `
                <input type="checkbox" id="subject_${subject}" value="${subject}" checked>
                <label for="subject_${subject}">${subject}</label>
            `;
            subjectContainer.appendChild(div);
        }
    });
}

// 更新班级选项（基于选择的年级）
function updateClassOptions() {
    const gradeSelect = document.getElementById('gradeSelect');
    const classSelect = document.getElementById('classSelect');
    const selectedGrades = Array.from(gradeSelect.selectedOptions).map(opt => opt.value);
    
    classSelect.innerHTML = '<option value="all">全部班级</option>';
    
    if (selectedGrades.includes('all') || selectedGrades.length === 0) {
        // 显示所有班级
        Object.values(configData.classes).flat().forEach(className => {
            const option = document.createElement('option');
            option.value = className;
            option.textContent = className;
            classSelect.appendChild(option);
        });
    } else {
        // 只显示选中年级的班级
        selectedGrades.forEach(grade => {
            if (configData.classes[grade]) {
                configData.classes[grade].forEach(className => {
                    const option = document.createElement('option');
                    option.value = className;
                    option.textContent = className;
                    classSelect.appendChild(option);
                });
            }
        });
    }
}

// 筛选学生
function filterStudents() {
    const gradeSelect = document.getElementById('gradeSelect');
    const classSelect = document.getElementById('classSelect');
    
    const selectedGrades = Array.from(gradeSelect.selectedOptions).map(opt => opt.value);
    const selectedClasses = Array.from(classSelect.selectedOptions).map(opt => opt.value);
    
    // 获取唯一的学生列表
    const studentMap = new Map();
    scoreData.forEach(row => {
        const studentKey = `${row.学号}|${row.姓名}|${row.年级}|${row.班级}`;
        if (!studentMap.has(studentKey)) {
            studentMap.set(studentKey, {
                学号: row.学号,
                姓名: row.姓名,
                年级: row.年级,
                班级: row.班级
            });
        }
    });
    
    let students = Array.from(studentMap.values());
    
    // 按年级筛选
    if (!selectedGrades.includes('all') && selectedGrades.length > 0) {
        students = students.filter(student => selectedGrades.includes(student.年级));
    }
    
    // 按班级筛选
    if (!selectedClasses.includes('all') && selectedClasses.length > 0) {
        students = students.filter(student => selectedClasses.includes(student.班级));
    }
    
    selectedStudents = students;
    updateStudentSelect();
}

// 更新学生选择下拉框
function updateStudentSelect() {
    const studentSelect = document.getElementById('studentSelect');
    const studentCount = document.getElementById('studentCount');
    
    studentSelect.innerHTML = '';
    studentCount.textContent = `共${selectedStudents.length}名学生`;
    
    selectedStudents.forEach((student, index) => {
        const option = document.createElement('option');
        option.value = index;
        option.textContent = `${student.班级} - ${student.姓名} (${student.学号})`;
        studentSelect.appendChild(option);
    });
    
    if (selectedStudents.length > 0) {
        studentSelect.disabled = false;
        currentPreviewStudent = selectedStudents[0];
    } else {
        studentSelect.disabled = true;
    }
}

// 预览成绩单
function previewScoreCard() {
    filterStudents();
    if (selectedStudents.length === 0) {
        showStatus('请先选择年级和班级！', 'error');
        return;
    }
    
    const studentIndex = document.getElementById('studentSelect').value || 0;
    currentPreviewStudent = selectedStudents[studentIndex];
    updatePreview();
}

// 更新预览
function updatePreview() {
    if (!currentPreviewStudent) return;
    
    const previewContainer = document.getElementById('previewContainer');
    previewContainer.innerHTML = '';
    
    // 创建成绩单HTML
    const scoreCard = createScoreCardHTML(currentPreviewStudent);
    previewContainer.appendChild(scoreCard);
}

// 创建单个学生的成绩单HTML
function createScoreCardHTML(student) {
    const paperSize = document.getElementById('paperSize').value;
    const orientation = document.getElementById('paperOrientation').value;
    
    // 获取选中的考次和科目
    const selectedExams = Array.from(document.querySelectorAll('#examCheckboxes input:checked'))
        .map(cb => cb.value);
    const selectedSubjects = Array.from(document.querySelectorAll('#subjectCheckboxes input:checked'))
        .map(cb => cb.value);
    
    // 过滤该学生的成绩数据
    const studentScores = scoreData.filter(row => 
        row.学号 === student.学号 && 
        selectedExams.includes(row.考次)
    );
    
    // 创建表格HTML
    let tableHTML = '<table class="score-table">';
    
    // 表头
    tableHTML += '<thead><tr>';
    tableHTML += '<th>科目</th>';
    selectedExams.forEach(exam => {
        tableHTML += `<th>${exam}</th>`;
    });
    tableHTML += '</tr></thead>';
    
    // 表格内容
    tableHTML += '<tbody>';
    selectedSubjects.forEach(subject => {
        tableHTML += '<tr>';
        tableHTML += `<td>${subject}</td>`;
        
        selectedExams.forEach(exam => {
            const scoreRow = studentScores.find(row => row.考次 === exam);
            const score = scoreRow ? (scoreRow[subject] || '') : '';
            tableHTML += `<td>${score}</td>`;
        });
        
        tableHTML += '</tr>';
    });
    tableHTML += '</tbody></table>';
    
    // 创建成绩单容器
    const card = document.createElement('div');
    card.className = `score-card ${paperSize} ${orientation}`;
    card.id = 'scoreCardToCapture';
    card.innerHTML = `
        <div class="score-header">
            <h1 class="score-title">成绩单</h1>
            <div class="student-info">
                <p>班级：${student.班级} | 姓名：${student.姓名} | 学号：${student.学号}</p>
            </div>
        </div>
        <div class="score-content">
            ${tableHTML}
        </div>
    `;
    
    // 根据纸张设置样式
    if (paperSize === 'a4') {
        card.style.width = orientation === 'landscape' ? '29.7cm' : '21cm';
        card.style.height = orientation === 'landscape' ? '21cm' : '29.7cm';
    } else {
        card.style.width = orientation === 'landscape' ? '14cm' : '10cm';
        card.style.height = orientation === 'landscape' ? '10cm' : '14cm';
    }
    
    return card;
}

// 切换预览的学生
function changePreviewStudent() {
    const studentIndex = document.getElementById('studentSelect').value;
    if (studentIndex >= 0 && studentIndex < selectedStudents.length) {
        currentPreviewStudent = selectedStudents[studentIndex];
        updatePreview();
    }
}

// 下载单个成绩单
function downloadSingleScoreCard() {
    if (!currentPreviewStudent) {
        showStatus('请先预览成绩单！', 'error');
        return;
    }
    
    const card = document.getElementById('scoreCardToCapture');
    if (!card) {
        showStatus('未找到成绩单元素！', 'error');
        return;
    }
    
    html2canvas(card, {
        scale: 2,
        useCORS: true,
        backgroundColor: '#ffffff'
    }).then(canvas => {
        canvas.toBlob(blob => {
            const filename = `${currentPreviewStudent.班级}-${currentPreviewStudent.姓名}-${currentPreviewStudent.学号}.png`;
            saveAs(blob, filename);
            showStatus('成绩单下载成功！', 'success');
        });
    }).catch(error => {
        showStatus('生成图片失败：' + error.message, 'error');
    });
}

// 批量生成成绩单
let isGenerating = false;
let cancelGeneration = false;

async function batchGenerate() {
    if (isGenerating) {
        showStatus('正在生成中，请稍候...', 'error');
        return;
    }
    
    filterStudents();
    if (selectedStudents.length === 0) {
        showStatus('没有可选的学生！', 'error');
        return;
    }
    
    isGenerating = true;
    cancelGeneration = false;
    
    // 显示进度条
    const progressSection = document.getElementById('batchProgress');
    const progressFill = document.getElementById('progressFill');
    const progressText = document.getElementById('progressText');
    
    progressSection.style.display = 'block';
    progressFill.style.width = '0%';
    progressText.textContent = '准备中...';
    
    const zip = new JSZip();
    const imgFolder = zip.folder("成绩单");
    
    const total = selectedStudents.length;
    let completed = 0;
    
    for (let i = 0; i < total; i++) {
        if (cancelGeneration) break;
        
        const student = selectedStudents[i];
        
        // 更新进度
        completed++;
        const percent = Math.round((completed / total) * 100);
        progressFill.style.width = percent + '%';
        progressText.textContent = `正在生成：${student.班级} - ${student.姓名} (${completed}/${total})`;
        
        // 创建成绩单并转为图片
        try {
            const card = createScoreCardHTML(student);
            document.body.appendChild(card);
            
            const canvas = await html2canvas(card, {
                scale: 2,
                useCORS: true,
                backgroundColor: '#ffffff'
            });
            
            document.body.removeChild(card);
            
            canvas.toBlob(blob => {
                const filename = `${student.班级}-${student.姓名}-${student.学号}.png`;
                imgFolder.file(filename, blob);
                
                // 如果是最后一个，生成ZIP文件
                if (completed === total || cancelGeneration) {
                    if (!cancelGeneration) {
                        progressText.textContent = '正在打包ZIP文件...';
                        zip.generateAsync({type: "blob"})
                            .then(function(content) {
                                saveAs(content, "学生成绩单.zip");
                                progressText.textContent = '批量生成完成！';
                                showStatus('批量生成完成，ZIP文件已下载！', 'success');
                                
                                // 延迟隐藏进度条
                                setTimeout(() => {
                                    progressSection.style.display = 'none';
                                }, 2000);
                            });
                    }
                    isGenerating = false;
                }
            });
        } catch (error) {
            console.error(`生成${student.姓名}的成绩单失败：`, error);
        }
        
        // 添加延迟避免阻塞
        await new Promise(resolve => setTimeout(resolve, 100));
    }
}

// 取消批量生成
function cancelBatch() {
    cancelGeneration = true;
    document.getElementById('batchProgress').style.display = 'none';
    showStatus('批量生成已取消', 'error');
    isGenerating = false;
}

// 显示状态消息
function showStatus(message, type) {
    const statusEl = document.getElementById('fileStatus');
    statusEl.textContent = message;
    statusEl.className = 'status ' + type;
    
    // 3秒后自动清除
    setTimeout(() => {
        statusEl.textContent = '';
        statusEl.className = 'status';
    }, 3000);
}