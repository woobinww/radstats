
const ExcelJS = require('exceljs');
const sqlite3 = require('sqlite3').verbose();
const path = require('path');
const fs = require('fs');

// 입력할 대상 월
// const targetMonth = '2025-04';

// DB 및 템플릿 파일 경로
const dbPath = path.join(__dirname, '../db/database.sqlite');
const templatePath = path.join(__dirname, '../templates/stat_form.xlsx');
// 저장할 파일 경로 없으면 자동 생성
const reportsDir = path.join(__dirname, '../reports');
if (!fs.existsSync(reportsDir)) fs.mkdirSync(reportsDir);


// 조건별 셀 매핑 (검사명 키워드는 없으면 null, 있으면 OR조건으로 사용)
const sheetMaps = {
  '총 통계': {
    'CR|기세린|외래': 'B7',
    'CR|기세린|기타': 'C7',
    'CR|강진우|외래': 'D7',
    'CR|강진우|기타': 'E7',
    'CR|유재중|외래': 'F7',
    'CR|유재중|기타': 'G7',
    'CR|김태겸|외래': 'H7',
    'CR|김태겸|기타': 'I7',
    'CR|김건우|외래': 'J7',
    'CR|김건우|기타': 'K7',
    'CR|*|*': 'L7',
    'US|기세린|외래|복부|abdo': 'B8',
    'US|기세린|기타|복부|abdo': 'C8',
    'US|강진우|외래|복부|abdo': 'D8',
    'US|강진우|기타|복부|abdo': 'E8',
    'US|유재중|외래|복부|abdo': 'F8',
    'US|유재중|기타|복부|abdo': 'G8',
    'US|김태겸|외래|복부|abdo': 'H8',
    'US|김태겸|기타|복부|abdo': 'I8',
    'US|김건우|외래|복부|abdo': 'J8',
    'US|김건우|기타|복부|abdo': 'K8',
    'US|*|*|복부|abdo': 'L8',
    'US|기세린|외래|Doppler|Dopper': 'B9',
    'US|기세린|기타|Doppler|Dopper': 'C9',
    'US|강진우|외래|Doppler|Dopper': 'D9',
    'US|강진우|기타|Doppler|Dopper': 'E9',
    'US|유재중|외래|Doppler|Dopper': 'F9',
    'US|유재중|기타|Doppler|Dopper': 'G9',
    'US|김태겸|외래|Doppler|Dopper': 'H9',
    'US|김태겸|기타|Doppler|Dopper': 'I9',
    'US|김건우|외래|Doppler|Dopper': 'J9',
    'US|김건우|기타|Doppler|Dopper': 'K9',
    'US|*|*|Doppler|Dopper': 'L9',
    'US|기세린|외래|!복부|!abdo|!Doppler|!dopper': 'B10',
    'US|기세린|기타|!복부|!abdo|!Doppler|!dopper': 'C10',
    'US|강진우|외래|!복부|!abdo|!Doppler|!dopper': 'D10',
    'US|강진우|기타|!복부|!abdo|!Doppler|!dopper': 'E10',
    'US|유재중|외래|!복부|!abdo|!Doppler|!dopper': 'F10',
    'US|유재중|기타|!복부|!abdo|!Doppler|!dopper': 'G10',
    'US|김태겸|외래|!복부|!abdo|!Doppler|!dopper': 'H10',
    'US|김태겸|기타|!복부|!abdo|!Doppler|!dopper': 'I10',
    'US|김건우|외래|!복부|!abdo|!Doppler|!dopper': 'J10',
    'US|김건우|기타|!복부|!abdo|!Doppler|!dopper': 'K10',
    'US|*|*|!복부|!abdo|!Doppler|!dopper': 'L10',
    'CT|기세린|외래': 'B11',
    'CT|기세린|기타': 'C11',
    'CT|강진우|외래': 'D11',
    'CT|강진우|기타': 'E11',
    'CT|유재중|외래': 'F11',
    'CT|유재중|기타': 'G11',
    'CT|김태겸|외래': 'H11',
    'CT|김태겸|기타': 'I11',
    'CT|김건우|외래': 'J11',
    'CT|김건우|기타': 'K11',
    'CT|*|*': 'L11',
    'MR|기세린|외래': 'B12',
    'MR|기세린|기타': 'C12',
    'MR|강진우|외래': 'D12',
    'MR|강진우|기타': 'E12',
    'MR|유재중|외래': 'F12',
    'MR|유재중|기타': 'G12',
    'MR|김태겸|외래': 'H12',
    'MR|김태겸|기타': 'I12',
    'MR|김건우|외래': 'J12',
    'MR|김건우|기타': 'K12',
    'MR|*|*': 'L12',
    'RF|기세린|외래': 'B13',
    'RF|기세린|기타': 'C13',
    'RF|강진우|외래': 'D13',
    'RF|강진우|기타': 'E13',
    'RF|유재중|외래': 'F13',
    'RF|유재중|기타': 'G13',
    'RF|김태겸|외래': 'H13',
    'RF|김태겸|기타': 'I13',
    'RF|김건우|외래': 'J13',
    'RF|김건우|기타': 'K13',
    'RF|*|*': 'L13',
    'CR|기세린|외래|age': 'B14',
    'CR|기세린|기타|age': 'C14',
    'CR|강진우|외래|age': 'D14',
    'CR|강진우|기타|age': 'E14',
    'CR|유재중|외래|age': 'F14',
    'CR|유재중|기타|age': 'G14',
    'CR|김태겸|외래|age': 'H14',
    'CR|김태겸|기타|age': 'I14',
    'CR|김건우|외래|age': 'J14',
    'CR|김건우|기타|age': 'K14',
    'CR|*|*|age': 'L14'
  },
  'MRI': {
    'MR|기세린|*|brain': 'B4',
    'MR|기세린|*|c-spine': 'B5',
    'MR|기세린|*|t-spine': 'B6',
    'MR|기세린|*|l-spine': 'B7',
    'MR|기세린|*|shoulder': 'B8',
    'MR|기세린|*|elbow': 'B9',
    'MR|기세린|*|wrist': 'B10',
    'MR|기세린|*|hand': 'B11',
    'MR|기세린|*|hip': 'B12',
    'MR|기세린|*|knee': 'B13',
    'MR|기세린|*|ankle': 'B14',
    'MR|기세린|*|foot': 'B15',
    'MR|기세린|*|enhance': 'B16',
    'MR|기세린|*|!brain|!c-spine|!t-spine|!l-spine|!shoulder|!elbow|!wrist|!hand|!hip|!knee|!ankle|!foot|!enhance': 'B19',
    'MR|기세린|*': 'B20',
    'MR|강진우|*|brain': 'C4',
    'MR|강진우|*|c-spine': 'C5',
    'MR|강진우|*|t-spine': 'C6',
    'MR|강진우|*|l-spine': 'C7',
    'MR|강진우|*|shoulder': 'C8',
    'MR|강진우|*|elbow': 'C9',
    'MR|강진우|*|wrist': 'C10',
    'MR|강진우|*|hand': 'C11',
    'MR|강진우|*|hip': 'C12',
    'MR|강진우|*|knee': 'C13',
    'MR|강진우|*|ankle': 'C14',
    'MR|강진우|*|foot': 'C15',
    'MR|강진우|*|enhance': 'C16',
    'MR|강진우|*|!brain|!c-spine|!t-spine|!l-spine|!shoulder|!elbow|!wrist|!hand|!hip|!knee|!ankle|!foot|!enhance': 'C19',
    'MR|강진우|*': 'C20',
    'MR|유재중|*|brain': 'D4',
    'MR|유재중|*|c-spine': 'D5',
    'MR|유재중|*|t-spine': 'D6',
    'MR|유재중|*|l-spine': 'D7',
    'MR|유재중|*|shoulder': 'D8',
    'MR|유재중|*|elbow': 'D9',
    'MR|유재중|*|wrist': 'D10',
    'MR|유재중|*|hand': 'D11',
    'MR|유재중|*|hip': 'D12',
    'MR|유재중|*|knee': 'D13',
    'MR|유재중|*|ankle': 'D14',
    'MR|유재중|*|foot': 'D15',
    'MR|유재중|*|enhance': 'D16',
    'MR|유재중|*|!brain|!c-spine|!t-spine|!l-spine|!shoulder|!elbow|!wrist|!hand|!hip|!knee|!ankle|!foot|!enhance': 'D19',
    'MR|유재중|*': 'D20',
    'MR|김태겸|*|brain': 'E4',
    'MR|김태겸|*|c-spine': 'E5',
    'MR|김태겸|*|t-spine': 'E6',
    'MR|김태겸|*|l-spine': 'E7',
    'MR|김태겸|*|shoulder': 'E8',
    'MR|김태겸|*|elbow': 'E9',
    'MR|김태겸|*|wrist': 'E10',
    'MR|김태겸|*|hand': 'E11',
    'MR|김태겸|*|hip': 'E12',
    'MR|김태겸|*|knee': 'E13',
    'MR|김태겸|*|ankle': 'E14',
    'MR|김태겸|*|foot': 'E15',
    'MR|김태겸|*|enhance': 'E16',
    'MR|김태겸|*|!brain|!c-spine|!t-spine|!l-spine|!shoulder|!elbow|!wrist|!hand|!hip|!knee|!ankle|!foot|!enhance': 'E19',
    'MR|김태겸|*': 'E20',
    'MR|김건우|*|brain': 'F4',
    'MR|김건우|*|c-spine': 'F5',
    'MR|김건우|*|t-spine': 'F6',
    'MR|김건우|*|l-spine': 'F7',
    'MR|김건우|*|shoulder': 'F8',
    'MR|김건우|*|elbow': 'F9',
    'MR|김건우|*|wrist': 'F10',
    'MR|김건우|*|hand': 'F11',
    'MR|김건우|*|hip': 'F12',
    'MR|김건우|*|knee': 'F13',
    'MR|김건우|*|ankle': 'F14',
    'MR|김건우|*|foot': 'F15',
    'MR|김건우|*|enhance': 'F16',
    'MR|김건우|*|!brain|!c-spine|!t-spine|!l-spine|!shoulder|!elbow|!wrist|!hand|!hip|!knee|!ankle|!foot|!enhance': 'F19',
    'MR|김건우|*': 'F20',
  },
  'CT': {
    'CT|기세린|*|brain': 'C4',
    'CT|기세린|*|pns': 'C5',
    'CT|기세린|*|facial': 'C6',
    'CT|기세린|*|brain|pns|facial': 'C7',
    'CT|기세린|*|c-spine|!3d': 'C8',
    'CT|기세린|*|t-spine|!3d': 'C9',
    'CT|기세린|*|l-spine|!3d': 'C10',
    'CT|기세린|*|tl-spine|!3d': 'C11',
    'CT|기세린|*|spine+3d': 'C12',
    'CT|기세린|*|spine': 'C13',
    'CT|기세린|*|pelvis|hip|!3d': 'C14',
    'CT|기세린|*|pelvis cbct 3d|hip cbct 3d': 'C15',
    'CT|기세린|*|pelvis|hip': 'C16',
    'CT|기세린|*|foot|!3d': 'C17',
    'CT|기세린|*|ankle|!3d': 'C18',
    'CT|기세린|*|knee|!3d': 'C19',
    'CT|기세린|*|tibia|!3d': 'C20',
    'CT|기세린|*|thigh|femur|!3d': 'C21',
    'CT|기세린|*|foot|ankle|knee|tibia|thigh|femur|!3d': 'C22',
    'CT|기세린|*|foot+3d': 'C23',
    'CT|기세린|*|ankle+3d': 'C24',
    'CT|기세린|*|knee+3d': 'C25',
    'CT|기세린|*|tibia+3d': 'C26',
    'CT|기세린|*|thigh cbct 3d|femur cbct 3d': 'C27',
    'CT|기세린|*|foot cbct 3d|ankle cbct 3d|knee cbct 3d|tibia cbct 3d|thigh cbct 3d|femur cbct 3d': 'C28',
    'CT|기세린|*|hand|!3d': 'C29',
    'CT|기세린|*|wrist|!3d': 'C30',
    'CT|기세린|*|forearm|!3d': 'C31',
    'CT|기세린|*|elbow|!3d': 'C32',
    'CT|기세린|*|shoulder|!3d': 'C33',
    'CT|기세린|*|humerus|!3d': 'C34',
    'CT|기세린|*|hand|wrist|forearm|elbow|shoulder|humerus|!3d': 'C35',
    'CT|기세린|*|hand+3d': 'C36',
    'CT|기세린|*|wrist+3d': 'C37',
    'CT|기세린|*|forearm+3d': 'C38',
    'CT|기세린|*|elbow+3d': 'C39',
    'CT|기세린|*|shoulder+3d': 'C40',
    'CT|기세린|*|humerus+3d': 'C41',
    'CT|기세린|*|hand cbct 3d|wrist cbct 3d|forearm cbct 3d|elbow cbct 3d|shoulder cbct 3d|humerus cbct 3d': 'C42',
    'CT|기세린|*|!brain|!pns|!facial|!spine|!pelvis|!hip|!foot|!ankle|!knee|!tibia|!thigh|!femur|!hand|!wrist|!forearm|!elbow|!shoulder|!humerus': 'C43',
    'CT|기세린|*': 'C44',
    'CT|강진우|*|brain': 'D4',
    'CT|강진우|*|pns': 'D5',
    'CT|강진우|*|facial': 'D6',
    'CT|강진우|*|brain|pns|facial': 'D7',
    'CT|강진우|*|c-spine|!3d': 'D8',
    'CT|강진우|*|t-spine|!3d': 'D9',
    'CT|강진우|*|l-spine|!3d': 'D10',
    'CT|강진우|*|tl-spine|!3d': 'D11',
    'CT|강진우|*|spine+3d': 'D12',
    'CT|강진우|*|spine': 'D13',
    'CT|강진우|*|pelvis|hip|!3d': 'D14',
    'CT|강진우|*|pelvis cbct 3d|hip cbct 3d': 'D15',
    'CT|강진우|*|pelvis|hip': 'D16',
    'CT|강진우|*|foot|!3d': 'D17',
    'CT|강진우|*|ankle|!3d': 'D18',
    'CT|강진우|*|knee|!3d': 'D19',
    'CT|강진우|*|tibia|!3d': 'D20',
    'CT|강진우|*|thigh|femur|!3d': 'D21',
    'CT|강진우|*|foot|ankle|knee|tibia|thigh|femur|!3d': 'D22',
    'CT|강진우|*|foot+3d': 'D23',
    'CT|강진우|*|ankle+3d': 'D24',
    'CT|강진우|*|knee+3d': 'D25',
    'CT|강진우|*|tibia+3d': 'D26',
    'CT|강진우|*|thigh cbct 3d|femur cbct 3d': 'D27',
    'CT|강진우|*|foot cbct 3d|ankle cbct 3d|knee cbct 3d|tibia cbct 3d|thigh cbct 3d|femur cbct 3d': 'D28',
    'CT|강진우|*|hand|!3d': 'D29',
    'CT|강진우|*|wrist|!3d': 'D30',
    'CT|강진우|*|forearm|!3d': 'D31',
    'CT|강진우|*|elbow|!3d': 'D32',
    'CT|강진우|*|shoulder|!3d': 'D33',
    'CT|강진우|*|humerus|!3d': 'D34',
    'CT|강진우|*|hand|wrist|forearm|elbow|shoulder|humerus|!3d': 'D35',
    'CT|강진우|*|hand+3d': 'D36',
    'CT|강진우|*|wrist+3d': 'D37',
    'CT|강진우|*|forearm+3d': 'D38',
    'CT|강진우|*|elbow+3d': 'D39',
    'CT|강진우|*|shoulder+3d': 'D40',
    'CT|강진우|*|humerus+3d': 'D41',
    'CT|강진우|*|hand cbct 3d|wrist cbct 3d|forearm cbct 3d|elbow cbct 3d|shoulder cbct 3d|humerus cbct 3d': 'D42',
    'CT|강진우|*|!brain|!pns|!facial|!spine|!pelvis|!hip|!foot|!ankle|!knee|!tibia|!thigh|!femur|!hand|!wrist|!forearm|!elbow|!shoulder|!humerus': 'D43',
    'CT|강진우|*': 'D44',
    'CT|유재중|*|brain': 'E4',
    'CT|유재중|*|pns': 'E5',
    'CT|유재중|*|facial': 'E6',
    'CT|유재중|*|brain|pns|facial': 'E7',
    'CT|유재중|*|c-spine|!3d': 'E8',
    'CT|유재중|*|t-spine|!3d': 'E9',
    'CT|유재중|*|l-spine|!3d': 'E10',
    'CT|유재중|*|tl-spine|!3d': 'E11',
    'CT|유재중|*|spine+3d': 'E12',
    'CT|유재중|*|spine': 'E13',
    'CT|유재중|*|pelvis|hip|!3d': 'E14',
    'CT|유재중|*|pelvis cbct 3d|hip cbct 3d': 'E15',
    'CT|유재중|*|pelvis|hip': 'E16',
    'CT|유재중|*|foot|!3d': 'E17',
    'CT|유재중|*|ankle|!3d': 'E18',
    'CT|유재중|*|knee|!3d': 'E19',
    'CT|유재중|*|tibia|!3d': 'E20',
    'CT|유재중|*|thigh|femur|!3d': 'E21',
    'CT|유재중|*|foot|ankle|knee|tibia|thigh|femur|!3d': 'E22',
    'CT|유재중|*|foot+3d': 'E23',
    'CT|유재중|*|ankle+3d': 'E24',
    'CT|유재중|*|knee+3d': 'E25',
    'CT|유재중|*|tibia+3d': 'E26',
    'CT|유재중|*|thigh cbct 3d|femur cbct 3d': 'E27',
    'CT|유재중|*|foot cbct 3d|ankle cbct 3d|knee cbct 3d|tibia cbct 3d|thigh cbct 3d|femur cbct 3d': 'E28',
    'CT|유재중|*|hand|!3d': 'E29',
    'CT|유재중|*|wrist|!3d': 'E30',
    'CT|유재중|*|forearm|!3d': 'E31',
    'CT|유재중|*|elbow|!3d': 'E32',
    'CT|유재중|*|shoulder|!3d': 'E33',
    'CT|유재중|*|humerus|!3d': 'E34',
    'CT|유재중|*|hand|wrist|forearm|elbow|shoulder|humerus|!3d': 'E35',
    'CT|유재중|*|hand+3d': 'E36',
    'CT|유재중|*|wrist+3d': 'E37',
    'CT|유재중|*|forearm+3d': 'E38',
    'CT|유재중|*|elbow+3d': 'E39',
    'CT|유재중|*|shoulder+3d': 'E40',
    'CT|유재중|*|humerus+3d': 'E41',
    'CT|유재중|*|hand cbct 3d|wrist cbct 3d|forearm cbct 3d|elbow cbct 3d|shoulder cbct 3d|humerus cbct 3d': 'E42',
    'CT|유재중|*|!brain|!pns|!facial|!spine|!pelvis|!hip|!foot|!ankle|!knee|!tibia|!thigh|!femur|!hand|!wrist|!forearm|!elbow|!shoulder|!humerus': 'E43',
    'CT|유재중|*': 'E44',
    'CT|김태겸|*|brain': 'F4',
    'CT|김태겸|*|pns': 'F5',
    'CT|김태겸|*|facial': 'F6',
    'CT|김태겸|*|brain|pns|facial': 'F7',
    'CT|김태겸|*|c-spine|!3d': 'F8',
    'CT|김태겸|*|t-spine|!3d': 'F9',
    'CT|김태겸|*|l-spine|!3d': 'F10',
    'CT|김태겸|*|tl-spine|!3d': 'F11',
    'CT|김태겸|*|spine+3d': 'F12',
    'CT|김태겸|*|spine': 'F13',
    'CT|김태겸|*|pelvis|hip|!3d': 'F14',
    'CT|김태겸|*|pelvis cbct 3d|hip cbct 3d': 'F15',
    'CT|김태겸|*|pelvis|hip': 'F16',
    'CT|김태겸|*|foot|!3d': 'F17',
    'CT|김태겸|*|ankle|!3d': 'F18',
    'CT|김태겸|*|knee|!3d': 'F19',
    'CT|김태겸|*|tibia|!3d': 'F20',
    'CT|김태겸|*|thigh|femur|!3d': 'F21',
    'CT|김태겸|*|foot|ankle|knee|tibia|thigh|femur|!3d': 'F22',
    'CT|김태겸|*|foot+3d': 'F23',
    'CT|김태겸|*|ankle+3d': 'F24',
    'CT|김태겸|*|knee+3d': 'F25',
    'CT|김태겸|*|tibia+3d': 'F26',
    'CT|김태겸|*|thigh cbct 3d|femur cbct 3d': 'F27',
    'CT|김태겸|*|foot cbct 3d|ankle cbct 3d|knee cbct 3d|tibia cbct 3d|thigh cbct 3d|femur cbct 3d': 'F28',
    'CT|김태겸|*|hand|!3d': 'F29',
    'CT|김태겸|*|wrist|!3d': 'F30',
    'CT|김태겸|*|forearm|!3d': 'F31',
    'CT|김태겸|*|elbow|!3d': 'F32',
    'CT|김태겸|*|shoulder|!3d': 'F33',
    'CT|김태겸|*|humerus|!3d': 'F34',
    'CT|김태겸|*|hand|wrist|forearm|elbow|shoulder|humerus|!3d': 'F35',
    'CT|김태겸|*|hand+3d': 'F36',
    'CT|김태겸|*|wrist+3d': 'F37',
    'CT|김태겸|*|forearm+3d': 'F38',
    'CT|김태겸|*|elbow+3d': 'F39',
    'CT|김태겸|*|shoulder+3d': 'F40',
    'CT|김태겸|*|humerus+3d': 'F41',
    'CT|김태겸|*|hand cbct 3d|wrist cbct 3d|forearm cbct 3d|elbow cbct 3d|shoulder cbct 3d|humerus cbct 3d': 'F42',
    'CT|김태겸|*|!brain|!pns|!facial|!spine|!pelvis|!hip|!foot|!ankle|!knee|!tibia|!thigh|!femur|!hand|!wrist|!forearm|!elbow|!shoulder|!humerus': 'F43',
    'CT|김태겸|*': 'F44',
    'CT|김건우|*|brain': 'G4',
    'CT|김건우|*|pns': 'G5',
    'CT|김건우|*|facial': 'G6',
    'CT|김건우|*|brain|pns|facial': 'G7',
    'CT|김건우|*|c-spine|!3d': 'G8',
    'CT|김건우|*|t-spine|!3d': 'G9',
    'CT|김건우|*|l-spine|!3d': 'G10',
    'CT|김건우|*|tl-spine|!3d': 'G11',
    'CT|김건우|*|spine+3d': 'G12',
    'CT|김건우|*|spine': 'G13',
    'CT|김건우|*|pelvis|hip|!3d': 'G14',
    'CT|김건우|*|pelvis cbct 3d|hip cbct 3d': 'G15',
    'CT|김건우|*|pelvis|hip': 'G16',
    'CT|김건우|*|foot|!3d': 'G17',
    'CT|김건우|*|ankle|!3d': 'G18',
    'CT|김건우|*|knee|!3d': 'G19',
    'CT|김건우|*|tibia|!3d': 'G20',
    'CT|김건우|*|thigh|femur|!3d': 'G21',
    'CT|김건우|*|foot|ankle|knee|tibia|thigh|femur|!3d': 'G22',
    'CT|김건우|*|foot+3d': 'G23',
    'CT|김건우|*|ankle+3d': 'G24',
    'CT|김건우|*|knee+3d': 'G25',
    'CT|김건우|*|tibia+3d': 'G26',
    'CT|김건우|*|thigh cbct 3d|femur cbct 3d': 'G27',
    'CT|김건우|*|foot cbct 3d|ankle cbct 3d|knee cbct 3d|tibia cbct 3d|thigh cbct 3d|femur cbct 3d': 'G28',
    'CT|김건우|*|hand|!3d': 'G29',
    'CT|김건우|*|wrist|!3d': 'G30',
    'CT|김건우|*|forearm|!3d': 'G31',
    'CT|김건우|*|elbow|!3d': 'G32',
    'CT|김건우|*|shoulder|!3d': 'G33',
    'CT|김건우|*|humerus|!3d': 'G34',
    'CT|김건우|*|hand|wrist|forearm|elbow|shoulder|humerus|!3d': 'G35',
    'CT|김건우|*|hand+3d': 'G36',
    'CT|김건우|*|wrist+3d': 'G37',
    'CT|김건우|*|forearm+3d': 'G38',
    'CT|김건우|*|elbow+3d': 'G39',
    'CT|김건우|*|shoulder+3d': 'G40',
    'CT|김건우|*|humerus+3d': 'G41',
    'CT|김건우|*|hand cbct 3d|wrist cbct 3d|forearm cbct 3d|elbow cbct 3d|shoulder cbct 3d|humerus cbct 3d': 'G42',
    'CT|김건우|*|!brain|!pns|!facial|!spine|!pelvis|!hip|!foot|!ankle|!knee|!tibia|!thigh|!femur|!hand|!wrist|!forearm|!elbow|!shoulder|!humerus': 'G43',
    'CT|김건우|*': 'G44',
  }
};

// "항목별 통계" 시트 타겟 월 구성
function generateLastYearToCurrentMonths(targetMonth) {
  const [targetYearStr, targetMonStr] = targetMonth.split('-').map(Number);
  const currentYear = parseInt(targetYearStr, 10);
  const lastYear = currentYear - 1;
  const months = [];

  for (let m = 1; m <= 12; m++) {
    months.push(`${lastYear}-${String(m).padStart(2, '0')}`);
  }
  for (let m = 1; m <= parseInt(targetMonStr, 10); m++) {
    months.push(`${currentYear}-${String(m).padStart(2, '0')}`);
  }
  return { months, lastYear, currentYear };
}

// "항목별 통계" 시트의 해당 자리 알파벳 인덱스
function getExcelColumnLetter(index) {
  const base = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  let result = '';
  while (index >= 0) {
    result = base[index % 26] + result;
    index = Math.floor(index / 26) - 1;
  }
  return result;
}

//  항목별 통계 시트용 rowMap 생성 함수
function createStatRowMap(lastYear, currentYear) {
  return {
    'CR': { rowMap: { [lastYear]: 4, [currentYear]: 5 }, startCol: 'B' },
    'CT': { rowMap: { [lastYear]: 10, [currentYear]: 11 }, startCol: 'B' },
    'MR': { rowMap: { [lastYear]: 16, [currentYear]: 17 }, startCol: 'B' },
    'US': { rowMap: { [lastYear]: 22, [currentYear]: 23 }, startCol: 'B' },
  };
}



// 메인 함수
async function writeStatistics(targetMonth) {
  const db = new sqlite3.Database(dbPath);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(templatePath);

  const [year, month] = targetMonth.split('-');
  const outputPath = path.join(
    __dirname,
    `../reports/${year}_${month}_영상의학과_월간통계.xlsx`
  );

  for (const sheetName in sheetMaps) {
    const sheet = workbook.getWorksheet(sheetName);
    const cellMap = sheetMaps[sheetName];

    if (!sheet) continue; // sheet 존재하지 않는 경우 넘어감. 오류 방지

    // "총 통계" 시트일 때만 A4에 날짜 표시
    if (sheetName === "총 통계") {
      sheet.getCell('A4').value = `${year}년 ${month}월`;
    }

    if (sheetName === "MRI") {
      sheet.getCell('A1').value = `의료진별 MRI 통계(${year}년 ${month}월)`;
    }
    if (sheetName === "CT") {
      sheet.getCell('A1').value = `진료과별 CT 통계(${year}년 ${month}월)`;
    }

    for (const key in cellMap) {
      const [장비, 처방의, 병동, ...검사명키워드] = key.split('|');
      const cell = cellMap[key];
  
      const whereClauses = [
        "REPLACE(substr(검사시간, 1, 7), '/', '-') = ?",
        "TRIM(장비) = ?", // TRIM : 문자열 앞뒤의 공백 제거
      ];
      const params = [targetMonth, 장비];
  
      if (처방의 && 처방의 !== '*') {
        whereClauses.push("TRIM(처방의) = ?");
        params.push(처방의);
      }
  
      if (병동 && 병동 !== '*') {
        if (병동 === '기타' || 병동.startsWith('!')) {
          const exclude = 병동 === '기타' ? '외래' : 병동.slice(1);
          whereClauses.push("TRIM(병동) != ?");
          params.push(exclude);
        } else {
          whereClauses.push("TRIM(병동) = ?");
          params.push(병동);
        }
      }
  
      // 기본 제외 조건
      whereClauses.push("NOT (장비 = 'MR' AND 검사명 LIKE '%외부%')");
      whereClauses.push("NOT (장비 = 'US' AND 검사명 LIKE '%상담%')");
      whereClauses.push("NOT (장비 = 'RF' AND 검사명 NOT LIKE '%외래%')");
  
      // 검사명 필터 처리
      if (검사명키워드.length > 0) {
        const includeRaw = 검사명키워드.filter(k => !k.startsWith('!'));
        const exclude = 검사명키워드.filter(k => k.startsWith('!')).map(k => k.slice(1));

        const orKeywords = [];
        const andKeywords = [];

        for (const word of includeRaw) {
          if (word.includes('+')) {
            andKeywords.push(...word.split('+'));
          } else {
            orKeywords.push(word);
          }
        }
  
        if (orKeywords.length > 0) {
          whereClauses.push(
            `(${orKeywords.map(() => '검사명 LIKE ? COLLATE NOCASE').join(' OR ')})`
          );
          params.push(...orKeywords.map(k => `%${k}%`));
        }

        if (andKeywords.length > 0) {
          whereClauses.push(
            `${andKeywords.map(() => '검사명 LIKE ? COLLATE NOCASE').join(' AND ')}`
          );
          params.push(...andKeywords.map(k => `%${k}%`));
        }

        if (exclude.length > 0) {
          whereClauses.push(
            `${exclude.map(() => '검사명 NOT LIKE ? COLLATE NOCASE').join(' AND ')}`
          );
          params.push(...exclude.map(k => `%${k}%`));
        }
      }
  
      const query = `
        SELECT COUNT(*) AS count
        FROM rad_exam
        WHERE ${whereClauses.join(' AND ')}
      `;
  
      const count = await new Promise((resolve, reject) => {
        db.get(query, params, (err, row) => {
          if (err) reject(err);
          else resolve(row.count);
        });
      });
  
      sheet.getCell(cell).value = count;
    }
  }

  // '항목별 통계' 시트 처리
  const { months: monthList, lastYear, currentYear } = generateLastYearToCurrentMonths(targetMonth);
  const 항목별Sheet = workbook.getWorksheet('항목별 통계');
  const 항목별통계Map = createStatRowMap(lastYear, currentYear)
  if (항목별Sheet) {
    for (const 장비 in 항목별통계Map) {
      for (const month of monthList) {
        const [y, m] = month.split('-');
        const col = getExcelColumnLetter(m - 1 + 1); // B 부터 시작
        const row = 항목별통계Map[장비].rowMap[y];
        if (!row) continue;
        const cell = `${col}${row}`;

        const count = await new Promise((resolve, reject) => {
          const where = [
            "substr(REPLACE(검사시간, '/', '-'), 1, 7) = ?",
            "TRIM(장비) = ?"
          ];
          const params = [month, 장비];

          if (장비 == 'MR') {
            where.push("검사명 NOT LIKE '%외부%'");
          }

          if (장비 == 'US') {
            where.push("검사명 NOT LIKE '%상담%'");
          }
          
          const query = `
            SELECT COUNT(*) AS count
            FROM rad_exam
            WHERE ${where.join(' AND ')}
          `;

          db.get(query, params, (err, row) => {
            if (err) reject(err);
            else resolve(row.count);
          });
        });

        항목별Sheet.getCell(cell).value = count;
      }
    }
  }

  await workbook.xlsx.writeFile(outputPath);
  db.close();
  console.log(`✅ ${targetMonth} 엑셀 저장 완료 → ${outputPath}`);
}

module.exports = { writeStatistics };


