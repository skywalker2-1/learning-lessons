/**
 * أكاديمية فوزية — استبانة أثر الدرس (الدرس 4-2)
 * Run createSurvey() ONCE. It creates: the Google Form, a linked Sheet,
 * an "التحليل" sheet with averages + charts. (Email alerts are set by the owner in the Sheet: Tools > Notification settings.)
 */
var TITLE = 'استبانة أثر الدرس: تسمية الجزيئات';

var LIKERT = [
  'وضّح الدرس أهداف التعلم المقصودة بشكل واضح.',
  'كانت الاستراتيجيات التدريسية المستخدمة مناسبة وفعالة.',
  'أشرك «تحدّي التسمية» الطلاب في التعلم النشط والعمل الجماعي.',
  'أضاف نشاط الواقع المعزز قيمة لفهم تركيب الجزيئات وعدد ذراتها.',
  'نمّت أسئلة التفكير العليا و«المدقق الذكي» مهارات التحليل والنقد وحل المشكلات.',
  'ربطت تذكرة الدخول وتذكرة الخروج التقويم بأهداف الدرس وأتاحت قياس تقدم الطلاب.',
  'أظهر الطلاب فهمًا لتسمية المركبات الجزيئية الثنائية والأحماض وكتابة صيغها.',
  'عزّز الربط بالرياضيات والهندسة والفيزياء وبالحياة الواقعية معنى الدرس.',
  'قدّم الدرس أفكارًا عملية يمكنني تطبيقها في صفي.',
  'بشكل عام، كان للدرس أثر إيجابي في التعلم.'
];
var OPEN = [
  'أكثر ما لاحظته قيمة في الدرس.',
  'الفكرة التي أفكر في تطبيقها في صفي.',
  'مقترحات لتطوير الدروس القادمة.'
];

function createSurvey() {
  var form = FormApp.create(TITLE);
  form.setDescription('استبانة قصيرة (دقيقتان) لقياس أثر الدرس. أكاديمية فوزية · fawziaacademy.com\nالأسئلة من 4 إلى 13 إجبارية، والبقية اختيارية.');
  form.setConfirmationMessage('شكرًا لمشاركتك. تم استلام إجاباتك.');
  form.setShowLinkToRespondAgain(false);
  form.setCollectEmail(false);

  form.addTextItem().setTitle('الاسم (اختياري)');
  form.addMultipleChoiceItem().setTitle('التخصص (اختياري)')
      .setChoiceValues(['كيمياء','فيزياء','أحياء','علوم','تخصص آخر']);
  form.addMultipleChoiceItem().setTitle('سنوات الخبرة في التدريس (اختياري)')
      .setChoiceValues(['أقل من 5','5–10','11–20','أكثر من 20']);

  LIKERT.forEach(function (q, i) {
    form.addScaleItem().setTitle((i + 1) + '. ' + q)
        .setBounds(1, 5).setLabels('لا أوافق بشدة', 'أوافق بشدة').setRequired(true);
  });
  OPEN.forEach(function (q, i) {
    form.addParagraphTextItem().setTitle((i + 14) + '. ' + q);
  });

  // Linked spreadsheet
  var ss = SpreadsheetApp.create(TITLE + ' (النتائج)');
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  var resp = null;
  for (var t = 0; t < 12 && !resp; t++) {
    Utilities.sleep(1500); SpreadsheetApp.flush();
    ss.getSheets().forEach(function (s) { if (/Form Responses|إجابات النموذج/.test(s.getName())) resp = s; });
  }
  if (!resp) throw new Error('لم تظهر ورقة الإجابات. شغّل buildAnalysis() بعد دقيقة.');
  resp.setName('Responses');
  buildAnalysis_(ss);

  Logger.log('رابط الاستبانة للمشاركين:\n' + form.getPublishedUrl());
  Logger.log('رابط تعديل الاستبانة:\n' + form.getEditUrl());
  Logger.log('رابط ملف النتائج والتحليل:\n' + ss.getUrl());
}

function buildAnalysis_(ss) {
  var sh = ss.getSheetByName('التحليل') || ss.insertSheet('التحليل');
  sh.clear(); sh.setRightToLeft(true);
  sh.getRange('A1:D1').setValues([['العبارة', 'المتوسط (من 5)', 'نسبة الموافقة (4–5)', 'عدد الإجابات']])
    .setFontWeight('bold').setBackground('#0B2E6B').setFontColor('#FFFFFF');
  var cols = 'EFGHIJKLMN'.split('');          // Likert answers are columns E..N in Responses
  LIKERT.forEach(function (q, i) {
    var r = i + 2, rng = 'Responses!' + cols[i] + '2:' + cols[i];
    sh.getRange(r, 1).setValue((i + 1) + '. ' + q);
    sh.getRange(r, 2).setFormula('=IFERROR(ROUND(AVERAGE(' + rng + '),2),0)');
    sh.getRange(r, 3).setFormula('=IFERROR(COUNTIF(' + rng + ',">=4")/COUNT(' + rng + '),0)').setNumberFormat('0%');
    sh.getRange(r, 4).setFormula('=COUNT(' + rng + ')');
  });
  var last = LIKERT.length + 1;
  sh.getRange(last + 2, 1).setValue('المتوسط العام').setFontWeight('bold');
  sh.getRange(last + 2, 2).setFormula('=IFERROR(ROUND(AVERAGE(B2:B' + last + '),2),0)').setFontWeight('bold');
  sh.setColumnWidth(1, 520); sh.setColumnWidth(2, 120); sh.setColumnWidth(3, 150); sh.setColumnWidth(4, 110);
  sh.getCharts().forEach(function (c) { sh.removeChart(c); });
  sh.insertChart(sh.newChart().setChartType(Charts.ChartType.BAR)
    .addRange(sh.getRange('A1:B' + last)).setPosition(last + 5, 1, 0, 0)
    .setOption('title', 'متوسط التقييم لكل عبارة (من 5)').setOption('hAxis', {minValue: 0, maxValue: 5})
    .setOption('legend', {position: 'none'}).setOption('width', 760).setOption('height', 420).build());
  sh.insertChart(sh.newChart().setChartType(Charts.ChartType.COLUMN)
    .addRange(sh.getRange('A1:A' + last)).addRange(sh.getRange('C1:C' + last)).setPosition(last + 28, 1, 0, 0)
    .setOption('title', 'نسبة الموافقة (4–5) لكل عبارة').setOption('vAxis', {minValue: 0, maxValue: 1, format: '0%'})
    .setOption('legend', {position: 'none'}).setOption('width', 760).setOption('height', 380).build());
  ss.setActiveSheet(sh);
}
