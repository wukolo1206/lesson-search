const CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQmguIi7xm9XNY7T__XkUjl1NxtLeDRm1ZP5SiRQWpPaQS9BUkSheMh6RGQhWx9v2xKIVNvSJMSJza_/pub?gid=748523163&single=true&output=csv";

// Global State
let quizData = []; // All valid questions
let groupedArticles = {}; // { originalTitleText: { grade, originalTitle, articles: { transferTitle: { content, questions: [] } } } }
let currentDraft = null; // Currently taking quiz
let studentInfo = {};

// DOM Elements
const els = {
    overlay: document.getElementById('loading-overlay'),
    screenLogin: document.getElementById('screen-login'),
    screenQuiz: document.getElementById('screen-quiz'),
    screenResult: document.getElementById('screen-result'),
    
    // Login
    inputClass: document.getElementById('input-class'),
    inputNumber: document.getElementById('input-number'),
    inputName: document.getElementById('input-name'),
    selectGrade: document.getElementById('select-grade'),
    selectArticle: document.getElementById('select-article'),
    btnStart: document.getElementById('btn-start'),
    btnTeacherGuide: document.getElementById('btn-teacher-guide'),
    
    // Quiz
    quizStudentInfo: document.getElementById('quiz-student-info'),
    quizArticleTitle: document.getElementById('quiz-article-title'),
    articleDisplayTitle: document.getElementById('article-display-title'),
    articleContent: document.getElementById('article-content'),
    questionsContainer: document.getElementById('questions-container'),
    btnSubmit: document.getElementById('btn-submit'),
    
    // Result
    resultTime: document.getElementById('result-time'),
    resultScore: document.getElementById('result-score'),
    resultStudent: document.getElementById('result-student'),
    resultArticleTitle: document.getElementById('result-article-title'),
    resultDetails: document.getElementById('result-details'),
    btnBackHome: document.getElementById('btn-back-home'),
    btnDownload: document.getElementById('btn-download'),
    reportCard: document.getElementById('report-card')
};

// Environment Detection
const IS_GAS = typeof google !== 'undefined' && google.script && google.script.run;

// Application Init
async function initApp() {
    try {
        if (IS_GAS) {
            console.log("Running in Google Apps Script environment.");
            loadFromGAS();
        } else {
            console.log("Running in local environment. Loading via CSV.");
            loadCSVData();
        }
    } catch (e) {
        console.error(e);
        Swal.fire('錯誤', '載入題庫發生異常。', 'error');
    }
}

function loadFromGAS() {
    google.script.run
        .withSuccessHandler(function(resp) {
            if (!resp || !resp.transfer) {
                console.error("No transfer data returned");
                Swal.fire('錯誤', '題庫載入失敗 (無資料返回)。', 'error');
                els.overlay.classList.add('hidden');
                return;
            }
            
            // Format GAS objects into string maps similar to PapaParse output
            let mockCSVData = resp.transfer.map(row => {
                let obj = {};
                for (let k in row) {
                    obj[String(k).trim()] = String(row[k] || '');
                }
                return obj;
            });
            
            processData(mockCSVData);
            els.overlay.classList.add('hidden');
            els.screenLogin.classList.remove('hidden');
        })
        .withFailureHandler(function(err) {
            console.error("GAS error:", err);
            els.overlay.classList.add('hidden');
            Swal.fire('錯誤', '無法連線至資料庫，請稍後再試。', 'error');
        })
        .getData(); // Call same getData API used by index.html
}

function loadCSVData() {
    Papa.parse(CSV_URL, {
        download: true,
        header: true,
        complete: function(results) {
            processData(results.data);
            els.overlay.classList.add('hidden');
            els.screenLogin.classList.remove('hidden');
        },
        error: function(err) {
            console.error(err);
            els.overlay.classList.add('hidden');
            Swal.fire('錯誤', '無法載入題庫，請檢查網路連線。', 'error');
        }
    });
}

// Process CSV Data
function processData(data) {
    // Filter out invalid items and empty rows
    const validData = data.filter(row => {
        if (!row['遷移文本內容'] || !row['AI_審查題目']) return false;
        const status = row['AI_綜合判定'] || '';
        return status.includes('通過'); 
    });

    groupedArticles = {};

    validData.forEach(row => {
        const textContent = row['遷移文本內容'].trim();
        const grade = row['年級'] || '未知';
        const originalTitle = row['原始案例標題'] ? row['原始案例標題'].trim() : '未分類原文章';
        
        // Extract transfer title: match 【...】 at the beginning
        let transferTitle = "未命名遷移文章";
        const titleMatch = textContent.match(/^【(.*?)】/);
        if (titleMatch) {
            transferTitle = titleMatch[1].trim();
        }

        // Grouping key for original article (e.g. "三年級_1. 王小明")
        const originalGroupKey = `${grade}_${originalTitle}`;

        if (!groupedArticles[originalGroupKey]) {
            groupedArticles[originalGroupKey] = {
                grade: grade,
                originalTitle: originalTitle,
                articles: {}
            };
        }

        // Inside original group, group by transfer title
        if (!groupedArticles[originalGroupKey].articles[transferTitle]) {
            groupedArticles[originalGroupKey].articles[transferTitle] = {
                id: `${originalGroupKey}_${transferTitle}`,
                title: transferTitle,
                originalTitle: originalTitle,
                content: textContent,
                questions: []
            };
        }

        groupedArticles[originalGroupKey].articles[transferTitle].questions.push({
            qText: row['AI_審查題目'],
            ans: row['AI_答案'],
            insight: row['認知歷程'],
            strategy: row['教學策略']
        });
    });

    populateArticleSelect();
}

function populateArticleSelect() {
    // Extract unique grades
    const grades = new Set();
    Object.values(groupedArticles).forEach(group => {
        if (group.grade && group.grade !== '未知') {
            grades.add(group.grade);
        }
    });

    // Sort grades by Chinese map
    const gradeWeight = { '一年級':1, '二年級':2, '三年級':3, '四年級':4, '五年級':5, '六年級':6 };
    const sortedGrades = Array.from(grades).sort((a, b) => {
        const wa = gradeWeight[a] || 99;
        const wb = gradeWeight[b] || 99;
        return wa - wb;
    });

    els.selectGrade.innerHTML = '<option value="" disabled selected>先選擇年級</option>';
    sortedGrades.forEach(grade => {
        const opt = document.createElement('option');
        opt.value = grade;
        opt.textContent = `${grade}`;
        els.selectGrade.appendChild(opt);
    });

    // When grade changes, update article options
    els.selectGrade.addEventListener('change', () => {
        const selectedGrade = els.selectGrade.value;
        
        // Reset article select
        els.selectArticle.innerHTML = '<option value="" disabled selected>請選擇老師指派的文章...</option>';
        els.selectArticle.disabled = false;
        els.selectArticle.classList.remove('bg-gray-200', 'cursor-not-allowed', 'text-gray-500');
        els.selectArticle.classList.add('bg-gray-50');

        // Filter and sort original groups for the selected grade
        const originalGroups = Object.values(groupedArticles)
            .filter(g => g.grade === selectedGrade)
            .sort((a, b) => a.originalTitle.localeCompare(b.originalTitle));

        originalGroups.forEach(group => {
            // Create an optgroup for each original article
            const optGroup = document.createElement('optgroup');
            optGroup.label = `原篇名：${group.originalTitle}`;
            
            // Sort transfer articles within this group
            const transferArticles = Object.values(group.articles).sort((a, b) => {
                return a.title.localeCompare(b.title);
            });

            transferArticles.forEach(art => {
                const opt = document.createElement('option');
                opt.value = art.id; // Combined id
                opt.textContent = `▶ 測驗：${art.title} (共 ${art.questions.length} 題)`;
                optGroup.appendChild(opt);
            });

            els.selectArticle.appendChild(optGroup);
        });
    });
}

// Format Article Text into Paragraphs
function formatArticleParagraphs(rawText) {
    let text = rawText.trim();
    // Remove title at the beginning
    text = text.replace(/^【.*?】\s*/, '');
    
    // Remove numbers like 1, 2, 3... at the end of sentences that might be sentence markers
    text = text.replace(/\d+$/gm, ''); 

    // Simple split by punctuations
    const sentences = text.split(/(?<=[。！？\.\!\?])\s*/);
    let paragraphs = [];
    let currentPara = [];

    sentences.forEach(s => {
        if (!s.trim()) return;
        currentPara.push(s.trim());
        if (currentPara.length >= 5) {
            paragraphs.push(currentPara.join(' '));
            currentPara = [];
        }
    });
    if (currentPara.length > 0) {
        paragraphs.push(currentPara.join(' '));
    }

    return paragraphs.map(p => `<p class="indent-8 text-neutral-700 font-medium">${p}</p>`).join('');
}

// Parse Options from Question text (A)(B)(C)(D) or (1)(2)(3)(4)
function parseOptions(qText) {
    let cleanedQText = qText.replace(/答案[:：].*?(?=\n|$)/g, '').replace(/解析[:：].*?(?=\n|$)/g, '').trim();
    
    // Matcher for (A)(B)(C)(D) or (1)(2)(3)(4) or ①②③④ using regex
    // This is a robust regex to split question body and options
    let regex = /[\(（]?[A-D1-4][\)）][^\(（A-D1-4]*/g;
    let options = cleanedQText.match(regex);
    let mainText = cleanedQText.replace(regex, '').trim();

    if (!options || options.length < 2) {
        // Fallback if formatting is weird, just return as is (essay format)
        return { mainText: cleanedQText, options: [] };
    }

    return {
        mainText: mainText,
        options: options.map(opt => {
            let labelMatch = opt.match(/[\(（]?([A-D1-4])[\)）]?\s*(.*)/);
            if (labelMatch) {
                return { key: labelMatch[1], text: labelMatch[2].trim() };
            }
            return { key: '', text: opt.trim() };
        })
    };
}

// Start Quiz Logic
els.btnStart.addEventListener('click', () => {
    const cls = els.inputClass.value.trim();
    const num = els.inputNumber.value.trim();
    const nam = els.inputName.value.trim();
    const artId = els.selectArticle.value;

    if (!cls || !num || !nam) {
        Swal.fire('提示', '請填寫完整的班級、座號與姓名資料。', 'warning');
        return;
    }
    if (!artId) {
        Swal.fire('提示', '請選擇要測驗的文章。', 'warning');
        return;
    }

    studentInfo = { class: cls, number: num, name: nam };
    
    // artId is now combined: `${originalGroupKey}_${transferTitle}`
    // We need to find it in the two-level structure
    currentDraft = null;
    for (const groupKey in groupedArticles) {
        for (const titleKey in groupedArticles[groupKey].articles) {
            const art = groupedArticles[groupKey].articles[titleKey];
            if (art.id === artId) {
                currentDraft = art;
                break;
            }
        }
        if (currentDraft) break;
    }

    if (!currentDraft) {
        Swal.fire('錯誤', '找不到該文章，請重新整理網頁。', 'error');
        return;
    }

    renderQuiz();
    
    els.screenLogin.classList.add('hidden');
    els.screenQuiz.classList.remove('hidden');
    window.scrollTo(0,0);
});

function renderQuiz() {
    const infoStr = `${studentInfo.class} ${studentInfo.number.padStart(2,'0')} ${studentInfo.name}`;
    els.quizStudentInfo.textContent = infoStr;
    els.quizArticleTitle.textContent = currentDraft.title;
    els.articleDisplayTitle.textContent = currentDraft.title;
    
    els.articleContent.innerHTML = formatArticleParagraphs(currentDraft.content);
    
    els.questionsContainer.innerHTML = '';
    
    currentDraft.questions.forEach((q, idx) => {
        const parsed = parseOptions(q.qText);
        
        const qBlock = document.createElement('div');
        qBlock.className = 'bg-white p-6 rounded-2xl shadow-sm border border-gray-100 transition-all hover:shadow-md';
        
        let html = `
            <div class="mb-4">
                <span class="inline-block bg-indigo-50 text-primary text-sm font-bold px-2 py-1 rounded mb-2">第 ${idx + 1} 題</span>
                <p class="text-lg font-bold text-gray-800 leading-relaxed">${parsed.mainText.replace(/\n/g, '<br>')}</p>
            </div>
            <div class="space-y-3 pl-2">
        `;

        if (parsed.options.length > 0) {
            parsed.options.forEach((opt) => {
                html += `
                <label class="option-label group block relative">
                    <input type="radio" name="q_${idx}" value="${opt.key}" class="option-radio">
                    <div class="flex items-center p-4 bg-gray-50 border border-gray-200 rounded-xl group-hover:bg-gray-100 group-hover:border-gray-300 transition-colors">
                        <div class="option-marker w-8 h-8 rounded-full border-2 border-gray-300 flex items-center justify-center font-bold text-gray-500 mr-4 transition-colors">
                            ${opt.key}
                        </div>
                        <span class="text-gray-700 text-lg group-hover:text-gray-900">${opt.text}</span>
                    </div>
                </label>`;
            });
        } else {
            html += `<textarea name="q_${idx}" rows="3" class="w-full border-gray-300 rounded-lg shadow-sm focus:border-primary focus:ring-primary p-3" placeholder="請填寫答案..."></textarea>`;
        }

        html += `</div>`;
        qBlock.innerHTML = html;
        els.questionsContainer.appendChild(qBlock);
    });
}

// Submit Quiz
els.btnSubmit.addEventListener('click', () => {
    let allAnswered = true;
    let scoreParams = [];

    currentDraft.questions.forEach((q, idx) => {
        const radios = document.getElementsByName(`q_${idx}`);
        let answered = false;
        let selectedValue = "";
        
        if (radios.length > 0 && radios[0].type === 'radio') {
            for (let r of radios) {
                if (r.checked) {
                    answered = true;
                    selectedValue = r.value;
                    break;
                }
            }
        } else if (radios.length > 0 && radios[0].type === 'textarea') {
            selectedValue = radios[0].value.trim();
            if (selectedValue) answered = true;
        }

        if (!answered) allAnswered = false;
        scoreParams.push({
            expected: q.ans,
            actual: selectedValue,
            qText: q.qText,
            parsed: parseOptions(q.qText)
        });
    });

    if (!allAnswered) {
        Swal.fire({
            title: '還有題目未作答！',
            text: '請檢查是否每題都已選擇答案才交卷喔。',
            icon: 'warning',
            confirmButtonText: '繼續作答'
        });
        return;
    }

    Swal.fire({
        title: '確定要交卷嗎？',
        text: "交卷後就能立刻看到成績囉！",
        icon: 'question',
        showCancelButton: true,
        confirmButtonColor: '#10B981',
        cancelButtonColor: '#d33',
        confirmButtonText: '確定交卷',
        cancelButtonText: '再檢查一下'
    }).then((result) => {
        if (result.isConfirmed) {
            calculateAndShowResult(scoreParams);
        }
    });
});

function calculateAndShowResult(scoreParams) {
    let correctCount = 0;
    els.resultDetails.innerHTML = '';

    scoreParams.forEach((sp, idx) => {
        // Simple Ans parse: usually Ans might be "A", "B", "1", "2" or exact string.
        let ansExtract = (sp.expected || "").match(/[A-D1-4]/);
        let expectedKey = ansExtract ? ansExtract[0] : sp.expected;
        
        let isCorrect = (sp.actual.toUpperCase() === expectedKey.toUpperCase());
        if (isCorrect) correctCount++;

        let actualText = sp.actual;
        let expectedText = expectedKey;
        // Map keys back to full text if possible for display
        const optActual = sp.parsed.options.find(o => o.key.toUpperCase() === sp.actual.toUpperCase());
        if (optActual) actualText = `(${optActual.key}) ${optActual.text}`;
        
        const optExpected = sp.parsed.options.find(o => o.key.toUpperCase() === expectedKey.toUpperCase());
        if (optExpected) expectedText = `(${optExpected.key}) ${optExpected.text}`;

        const div = document.createElement('div');
        div.className = `p-4 rounded-xl border ${isCorrect ? 'bg-green-50 border-green-100' : 'bg-red-50 border-red-100'}`;
        div.innerHTML = `
            <div class="flex items-start">
                <div class="mt-1 mr-3 flex-shrink-0">
                    ${isCorrect ? 
                        `<svg class="w-6 h-6 text-green-500" fill="currentColor" viewBox="0 0 20 20"><path fill-rule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z" clip-rule="evenodd"/></svg>` : 
                        `<svg class="w-6 h-6 text-red-500" fill="currentColor" viewBox="0 0 20 20"><path fill-rule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zM8.707 7.293a1 1 0 00-1.414 1.414L8.586 10l-1.293 1.293a1 1 0 101.414 1.414L10 11.414l1.293 1.293a1 1 0 001.414-1.414L11.414 10l1.293-1.293a1 1 0 00-1.414-1.414L10 8.586 8.707 7.293z" clip-rule="evenodd"/></svg>`
                    }
                </div>
                <div>
                    <p class="font-bold text-gray-800 mb-1">第 ${idx + 1} 題：${sp.parsed.mainText}</p>
                    <p class="text-sm ${isCorrect ? 'text-green-700' : 'text-red-700 mb-1'}">
                        您的作答：${actualText}
                    </p>
                    ${!isCorrect ? `<p class="text-sm text-gray-600 font-medium">正確答案：${expectedText}</p>` : ''}
                </div>
            </div>
        `;
        els.resultDetails.appendChild(div);
    });

    const finalScore = Math.round((correctCount / scoreParams.length) * 100);
    els.resultScore.innerHTML = `${finalScore}<span class="text-sm font-normal ml-1">分</span>`;
    els.resultStudent.textContent = `${studentInfo.class} ${studentInfo.number.padStart(2,'0')} ${studentInfo.name}`;
    els.resultArticleTitle.textContent = currentDraft.title;
    
    // Set current time
    const now = new Date();
    els.resultTime.textContent = now.toLocaleString('zh-TW', { year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit' });

    els.screenQuiz.classList.add('hidden');
    els.screenResult.classList.remove('hidden');
    window.scrollTo(0, 0);
    
    if (finalScore === 100) {
        Swal.fire({
            title: '太棒了！滿分！',
            text: '恭喜你全對！記得下載成績單給老師喔！',
            icon: 'success',
            confirmButtonText: '好棒'
        });
    }
}

// Download Screenshot
els.btnDownload.addEventListener('click', () => {
    // Show a loading state on button
    const ogText = els.btnDownload.innerHTML;
    els.btnDownload.innerHTML = '<span class="animate-pulse">正在產生圖檔...</span>';
    els.btnDownload.disabled = true;

    // Small delay to ensure rendering
    setTimeout(() => {
        html2canvas(els.reportCard, {
            scale: 2, // High res
            backgroundColor: '#ffffff',
            useCORS: true
        }).then(canvas => {
            const imgData = canvas.toDataURL('image/jpeg', 0.9);
            const link = document.createElement('a');
            link.download = `閱讀測驗成績_${studentInfo.name}_${currentDraft.title}.jpg`;
            link.href = imgData;
            link.click();
            
            els.btnDownload.innerHTML = ogText;
            els.btnDownload.disabled = false;
            
            Swal.fire({
                title: '下載完成',
                text: '成績圖片已儲存至您的裝置，請傳送給老師。',
                icon: 'success',
                timer: 3000,
                showConfirmButton: false
            });
        }).catch(err => {
            console.error(err);
            els.btnDownload.innerHTML = ogText;
            els.btnDownload.disabled = false;
            Swal.fire('錯誤', '產生圖片失敗，請嘗試使用裝置內建的螢幕截圖功能。', 'error');
        });
    }, 300);
});

// Back Home
els.btnBackHome.addEventListener('click', () => {
    Swal.fire({
        title: '確定要回首頁嗎？',
        text: '回到首頁將會清除目前的作答紀錄喔！',
        icon: 'warning',
        showCancelButton: true,
        confirmButtonText: '回首頁',
        cancelButtonText: '取消'
    }).then((result) => {
        if (result.isConfirmed) {
            // Reset state logically. Refreshing is deeply safest to clear cache states.
            window.location.reload();
        }
    });
});

// Teacher Guide
if (els.btnTeacherGuide) {
    els.btnTeacherGuide.addEventListener('click', () => {
        Swal.fire({
            title: '教師加題指南',
            html: `
                <div class="text-left text-sm space-y-3 mt-4 text-gray-700 h-64 overflow-y-auto custom-scrollbar pr-2">
                    <p>要在此系統新增考卷給學生考，請在您雲端的<strong>學習遷移題目 Excel</strong> 底下新增資料，並確保以下 <strong>6個欄位</strong> 有填寫：</p>
                    <hr>
                    <ol class="list-decimal pl-5 space-y-2">
                        <li><strong>年級：</strong>例如填寫「三年級」。這決定了文章出現的分類。</li>
                        <li><strong>原始案例標題：</strong>填寫任意篇名。決定了網頁下拉選單的大標題分類。</li>
                        <li><strong>遷移文本內容：</strong>最前面一定要加實心括號 <code>【文章的標題名稱】</code>，接著再貼上文章內容。這是網頁判斷子選項名稱的依據。</li>
                        <li><strong>AI_審查題目：</strong>選擇題的選項開頭請務必要打英文字母或數字，例如 <code>(A)</code>, <code>(B)</code> 或 <code>(1)</code>, <code>(2)</code>。</li>
                        <li><strong>AI_答案：</strong>請填寫正確解答（例如 A 或 (A)）。</li>
                        <li><strong>AI_綜合判定：</strong>請務必填寫<span class="text-green-600 font-bold">「通過」</span>兩個字。沒有「通過」的題目不會顯示在前端！</li>
                    </ol>
                    <hr>
                    <p class="text-xs text-gray-400 mt-2">填寫完畢儲存後，等待約 3 分鐘，學生的網頁選單就會自動更新您的新考卷囉！</p>
                </div>
            `,
            width: '600px',
            confirmButtonText: '我了解了',
            confirmButtonColor: '#4F46E5'
        });
    });
}

// Boot
window.addEventListener('DOMContentLoaded', initApp);
