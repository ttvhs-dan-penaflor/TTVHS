/** TTVHS Learning Hub – TEST recording backend. Paste into Apps Script attached to the Google Sheet.
 * No login/authentication. Do not use for official graded records without school-approved authentication.
 */
const SPREADSHEET_ID = '1Ggn1frnZruTljRvx2m2-jNY98oBpIHjet1LGYwAiW0Y';
const VALID_LESSONS = {g10lesson:true,g10culturelesson:true,g11lesson:true,g7lesson:true};
function doGet() {
  return HtmlService.createHtmlOutputFromFile('TeacherReports')
    .setTitle('TTVHS Teacher Reports')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function doPost(e) {
  try {
    const raw = e && e.postData && e.postData.contents || '';
    if (raw.length > 30000) throw new Error('Too much data');
    const d = JSON.parse(raw);
    if (d.type !== 'lesson_mastery' || !VALID_LESSONS[d.lessonId]) throw new Error('Invalid record');
    if (!Array.isArray(d.attempts) || d.attempts.length < 3 || d.attempts.length > 30) throw new Error('Invalid attempts');
    const completedIds = new Set(d.attempts.filter(a => a.correct === true).map(a => a.questionId));
    if (completedIds.size !== 3 || !['q1','q2','q3'].every(q => completedIds.has(q))) throw new Error('Incomplete quiz');
    const cell = x => { const s=String(x == null ? '' : x).slice(0,400); return /^[=+@\-\t\r]/.test(s)?"'"+s:s; };
    const must = x => { const s=cell(x).trim();if(!s)throw new Error('Missing field');return s; };
    const studentName=must(d.studentName),studentId=must(d.studentId),recordId=must(d.recordId);
    const teacher=must(d.teacher),grade=must(d.gradeLevel),term=must(d.term),subject=must(d.subject),section=must(d.section),lessonId=must(d.lessonId),title=must(d.lessonTitle);
    if (Number(d.questionsCorrect)!==3 || Number(d.questionsRequired)!==3) throw new Error('Invalid score');
    const date=x => {const t=new Date(x);if(isNaN(t.getTime()))throw new Error('Invalid date');return t;};
    const audioEnd=date(d.audioFinishedAt),completed=date(d.completedAt);
    if (completed.getTime()<audioEnd.getTime()) throw new Error('Invalid order');
    const ss=SpreadsheetApp.openById(SPREADSHEET_ID);
    const lock=LockService.getScriptLock();lock.waitLock(20000);
    try {
      const sheet=ss.getSheetByName('Lesson Completions'), attemptsSheet=ss.getSheetByName('Quiz Attempts');
      if(!sheet||!attemptsSheet)throw new Error('Missing tabs');
      // Idempotency: repeated mobile submits should not create duplicates.
      const last=sheet.getLastRow();
      if(last>1){const ids=sheet.getRange(2,1,last-1,1).getValues().flat();if(ids.includes(recordId))return json({ok:true,duplicate:true});}
      const attempts=d.attempts.map((a,i)=>[cell(recordId)+':'+i,studentId,studentName,teacher,grade,term,subject,section,lessonId,cell(a.questionId),cell(a.selectedAnswer),a.correct===true?'TRUE':'FALSE',a.explanationShown===true?'TRUE':'FALSE',date(a.attemptedAt),cell(d.sessionId)]);
      attemptsSheet.getRange(attemptsSheet.getLastRow()+1,1,attempts.length,15).setValues(attempts);
      sheet.appendRow([recordId,studentId,studentName,teacher,grade,term,subject,section,lessonId,title,audioEnd,3,3,Number(d.wrongAttempts)||0,completed,'Completed',cell(d.sessionId)]);
      const log=ss.getSheetByName('Activity Logs');
      if(log)log.appendRow([recordId+':complete',studentId,studentName,teacher,grade,term,subject,section,lessonId,'Lesson Mastered',completed,cell(d.sessionId),'3/3 questions answered correctly']);
      return json({ok:true});
    } finally {lock.releaseLock();}
  } catch(err) {console.error(err);return json({ok:false,error:String(err).slice(0,180)});}
}
function json(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);}

// Teacher-only reporting — set REPORT_ACCESS_KEY in Apps Script Project Settings > Script Properties.
// Use a long random teacher access passphrase, not a PIN or student name.
function getTeacherReport(accessKey, filters) {
  const configured = PropertiesService.getScriptProperties().getProperty('REPORT_ACCESS_KEY');
  if (!configured || configured.length < 12) throw new Error('Teacher report access is not configured. Ask the administrator to set REPORT_ACCESS_KEY (12+ characters).');
  if (String(accessKey || '') !== configured) throw new Error('Incorrect teacher access code.');
  filters = filters || {};
  const section = String(filters.section || 'All').slice(0,70);
  const grade = String(filters.grade || 'All').slice(0,20);
  const teacher = String(filters.teacher || 'All').slice(0,80);
  const subject = String(filters.subject || 'All').slice(0,100);
  const name = String(filters.name || '').trim().toLowerCase().slice(0,100);
  const day = String(filters.date || '').slice(0,10);
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Lesson Completions');
  if (!sheet) throw new Error('Lesson Completions sheet not found');
  const last = sheet.getLastRow();
  const zone = 'Asia/Manila';
  const rows = last > 1 ? sheet.getRange(2, 1, Math.min(last - 1, 10000), 17).getValues() : [];
  const records = rows.filter(r => String(r[15] || '').toLowerCase() === 'completed').map(r => ({
    student: String(r[2]||''), teacher: String(r[3]||''), grade: String(r[4]||''),
    term: String(r[5]||''),subject:String(r[6]||''),section:String(r[7]||''),
    lesson: String(r[9]||''), wrong: Number(r[13])||0,
    date: r[14] instanceof Date ? Utilities.formatDate(r[14],zone,'yyyy-MM-dd') : String(r[14]).slice(0,10),
    completed: r[14] instanceof Date ? Utilities.formatDate(r[14],zone,'yyyy-MM-dd HH:mm:ss') : String(r[14]),
  })).filter(r =>
     (section==='All'||r.section===section) && (grade==='All'||r.grade===grade) &&
     (teacher==='All'||r.teacher===teacher) && (subject==='All'||r.subject===subject) &&
     (!name||r.student.toLowerCase().includes(name))
  );
  records.sort((a,b)=>b.completed.localeCompare(a.completed));
  const daily = records.filter(r=>!day || r.date===day);
  const uniqueDaily = [...new Set(daily.map(r=>r.grade+'|'+r.section+'|'+r.student))].length;
  const uniqueProgress = [...new Set(records.map(r=>r.grade+'|'+r.section+'|'+r.student))].length;
  return {date:day, daily:daily.slice(0,1000), progress:records.slice(0,1500),
    stats:{dailyCompletions:daily.length,dailyStudents:uniqueDaily,progressStudents:uniqueProgress},
    truncated: daily.length>1000||records.length>1500,generated:Utilities.formatDate(new Date(),zone,'yyyy-MM-dd HH:mm:ss')};
}
