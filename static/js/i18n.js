(function () {
  const LANG = window.APP_LANG || localStorage.getItem('app_lang') || 'th';

  const translations = {
    en: {
      "แผงกดคะแนน": "Score keypad",
      "สถานีนี้": "This station",
      "รวมรอบ": "Round total",
      "แตะตัวเลขเพื่อบันทึกทันที · ครบ 4 ระยะแล้วจะไปสถานีถัดไปเอง": "Tap a number to save instantly · moves to the next station after all 4 distances",
      "ใบนี้กำลังถูกคีย์โดย": "This scorecard is being entered by",
      "เพื่อกันคะแนนทับกัน หน้านี้จึงดูได้อย่างเดียว ถ้าเครื่องเดิมเลิกใช้แล้วให้กดรับช่วงคีย์ต่อ": "To avoid overwriting scores this page is read-only. If the other device is no longer in use, take over.",
      "รับช่วงคีย์ต่อ": "Take over",
      "← คิวสนาม": "← Court queue",
      "คิวสนาม": "Court queue",
      "โหมดจอแสดงผล": "Display mode",
      "เครื่องมือ ▾": "Tools ▾",
      "หน้ารวมทุกสนาม": "All courts overview",
      "ตีจบแล้ว": "Finished",
      "รอคิว": "Waiting",
      "คนถัดไป": "Next",
      "รับรองผลแล้ว": "Approved",
      "คะแนน": "points",
      "ดู / แก้ไข": "View / edit",
      "คีย์ต่อ": "Continue",
      "เริ่มคีย์": "Start scoring",
      "หน้านี้อัปเดตเองทุก 15 วินาที": "This page refreshes every 15 seconds",
      "ยังไม่มีนักกีฬาในสนามนี้": "No athletes on this court yet",
      "ยังไม่มีรายชื่อรอบ 2 ในสนามนี้ (รายชื่อจะขึ้นหลังรอบ 1 จบครบทุกสนาม)": "No round 2 list for this court yet (it appears after round 1 finishes on every court)",
      "บันทึกเป็น JPG": "Save as JPG",
      "พร้อมบันทึก": "Ready",
      "กำลังบันทึก…": "Saving…",
      "บันทึกแล้ว": "Saved",
      "Court ID ของอีเวนต์นี้": "Court IDs for this event",
      "ยังใช้รหัสตั้งต้น – เข้าระบบไม่ได้จนกว่าจะตั้งใหม่": "Still on a default password – blocked until reset",
      "ตั้งรหัสผ่านใหม่": "Set new password",
      "ตั้งรหัส": "Set",
      "รหัสใหม่": "New password",
      "รหัสผ่าน (อย่างน้อย 8 ตัว)": "Password (min. 8 characters)",
      "ผลรอบนี้ได้รับการรับรองแล้ว": "This round result has been approved",
      "กรรมการยกคะแนน": "Scoring referee",
      'หน้าหลัก': 'Home',
      'สร้างอีเวนต์': 'Create event',
      'ธีม': 'Theme', 'ธีมของระบบ': 'Site themes', 'สร้างธีมใหม่': 'New theme', 'แก้ไขธีม': 'Edit theme', 'เปิดใช้ทั้งระบบ': 'Use site-wide', 'กำลังใช้': 'Active',
      'สร้างหลายอีเวนต์': 'Bulk create events', 'ศูนย์จับสลาก': 'Draw center', '🎲 ศูนย์จับสลาก': '🎲 Draw center', '+ สร้างหลายอีเวนต์': '+ Bulk create events', 'สร้างหลายอีเวนต์พร้อมกัน': 'Create multiple events', 'ผลการจับสลาก': 'Draw results', '🎲 จับสลากพร้อมกัน': '🎲 Draw all selected', 'ตรวจสอบก่อนนำเข้า': 'Review before import',
      'สร้างอีเวนต์ Shooting': 'Create Shooting Event',
      'แก้ไขอีเวนต์': 'Edit event',
      'ผู้ใช้': 'Users',
      'ออกจากระบบ': 'Logout',
      'เข้าสู่ระบบ': 'Login',
      'ชื่อผู้ใช้': 'Username',
      'รหัสผ่าน': 'Password',
      'สิทธิ์': 'Role',
      'สร้างผู้ใช้ใหม่': 'Create new user',
      'รายการผู้ใช้': 'User list',
      'บันทึก': 'Save',
      'บันทึกการแก้ไข': 'Save changes',
      'แดชบอร์ด': 'Dashboard',
      'ภาพรวมของรายการแข่งขันและนักกีฬาในระบบ': 'Overview of events and athletes in the system',
      'อีเวนต์': 'Events',
      'รายการแข่งขัน': 'Competition events',
      'เลือกดูหน้ารวม สถิติ และจัดการนักกีฬาได้จากตารางนี้': 'Open overview, statistics, and athlete management from this table',
      'คะแนนสูงสุดในระบบ': 'Highest score in the system',
      'สถิติสังกัดที่ทำคะแนนดีที่สุด': 'Best-scoring affiliation statistics',
      'คะแนนสูงสุด': 'Highest score',
      'คะแนนรวม': 'Total',
      'คะแนน': 'points',
      'ชื่องาน': 'Event name',
      'รุ่น': 'Group',
      'รุ่นแข่งขัน': 'Competition group',
      'ประเภท': 'Category',
      'วันแข่งขัน': 'Competition date',
      'สนาม': 'Lane',
      'จำนวนสนาม': 'Number of lanes',
      'จัดการ': 'Manage',
      'หน้ารวม': 'Overview',
      'สถิติ': 'Statistics',
      'นักกีฬา': 'Athletes',
      'แก้ไข': 'Edit',
      'ลบ': 'Delete',
      'ยังไม่มีอีเวนต์': 'No events yet',
      'ยังไม่มีข้อมูล': 'No data yet',
      'ทั่วไป': 'Open',
      'อาวุโส': 'Senior',
      'เยาวชน': 'Youth',
      'ชาย': 'Men',
      'หญิง': 'Women',
      'ผสม': 'Mixed',
      'มี': 'Yes',
      'ไม่มี': 'No',
      'จากรอบแรกเข้ารอบถัดไปทันที กี่คน': 'Direct qualifiers from round 1',
      'มีรอบ 2 หรือไม่': 'Enable round 2?',
      'ถ้ามีรอบ 2 ให้ตีถึงลำดับที่เท่าไหร่': 'Round 2 cutoff rank',
      'จากรอบ 2 คัดเข้ารอบต่อกี่คน': 'Round 2 advancers',
      'รอบถัดไปคือรอบอะไร': 'Next round label',
      'รอบที่กำลังดู': 'Viewing round',
      'รอบ 1': 'Qualification Round 1',
      'รอบ 2': 'Qualification Round 2',
      'รอบที่ 1': 'Qualification 1',
      'รอบที่ 2': 'Qualification 2',
      'รอบ 8 คน': 'Quarterfinal',
      'รอบ 4 คน': 'Semifinal',
      'รอบรองชนะเลิศ': 'Semifinal',
      'รอบชิงชนะเลิศ': 'Final',
      'ตารางประกบ': 'Bracket',
      'หน้าประกบคู่': 'Bracket page',
      'พิมพ์ Scorecard หลายคน': 'Print multiple scorecards',
      'บันทึกเป็นรูปภาพ': 'Save as image',
      'พิมพ์ตาราง': 'Print table',
      'ย่อ/ขยาย': 'Zoom',
      'ความกว้างคู่': 'Match width',
      'ระยะระหว่างรอบ': 'Round spacing',
      'เลือกรอบที่จะเลื่อน': 'Round to move',
      'เลื่อนขึ้น–ลง': 'Move up/down',
      'พอดีหน้าจอ': 'Fit to screen',
      'คืนค่า': 'Reset',
      'รอคิว': 'Waiting',
      'กำลังตี': 'In progress',
      'ตีเสร็จแล้ว': 'Finished',
      'เข้ารอบตรง': 'Direct qualifier',
      'ผ่านจากรอบ 2': 'Advanced from round 2',
      'มีสิทธิ์รอบ 2': 'Eligible for round 2',
      'ตกรอบ': 'Eliminated',
      'ลำดับการตี': 'Shooting order',
      'ลำดับในสนาม': 'Lane order',
      'สถานะ': 'Status',
      'หมายเลข': 'no.',
      'ชื่อ': 'Name',
      'ชื่อ:': 'Name:',
      'สังกัด': 'Affiliation',
      'สังกัด:': 'Affiliation:',
      'สนาม:': 'Lane:',
      'ลำดับการตี:': 'Shooting order:',
      'เพิ่มนักกีฬาและจัดลำดับยิง': 'Add athletes and set shooting order',
      'ไปหน้ารวม': 'Go to overview',
      'หมายเลขและลำดับรวม ระบบสร้างให้อัตโนมัติ': 'Bib number and overall order are generated automatically',
      'เพิ่มนักกีฬา': 'Add athlete',
      'การกระจายสนามอัตโนมัติ': 'Automatic lane distribution',
      'สนาม = วนตามจำนวนสนาม เช่น 1, 2, 3, 4 แล้วกลับไป 1': 'Lane assignment cycles by lane count, e.g. 1, 2, 3, 4, then back to 1',
      'สุ่มลำดับทั้งหมด': 'Randomize all orders',
      'นำเข้ารายชื่อนักกีฬาจาก Excel': 'Import athlete list from Excel',
      'ไฟล์ต้องเป็น .xlsx และต้องมี 2 คอลัมน์คือ': 'File must be .xlsx with 2 columns:',
      'เลือกไฟล์ Excel': 'Choose Excel file',
      'อัปโหลด Excel': 'Upload Excel',
      'ดาวน์โหลดไฟล์ตัวอย่าง': 'Download template file',
      'ยังไม่มีนักกีฬา': 'No athletes yet',
      'ประเภทสุดยอดความแม่นยำ (SHOOTING)': 'Precision Shooting',
      'กลับหน้ารวม': 'Back to overview',
      'พิมพ์แบบฟอร์ม': 'Print form',
      'รอบที่คีย์': 'Entry round',
      'สถานีที่': 'Station',
      'ยิงลูกเดี่ยว': 'Single ball shot',
      'ยิงลูกเหนือเป้า': 'Target-ball jump shot',
      'ยิงลูกกลางขวาง': 'Middle obstacle shot',
      'ยิงลูกขาวเหนือลูกดำ': 'White over black ball shot',
      'ยิงลูกเป้า': 'Target ball shot',
      'รอบแข่งขัน': 'Round',
      'รวม': 'Total',
      'แดง': 'Red',
      'กรรมการตัดสิน': 'Umpire',
      'กรรมการบันทึกคะแนน': 'Recorder',
      'ลายเซ็นกรรมการตัดสิน': 'Referee signature',
      'ลายเซ็นกรรมการบันทึกคะแนน': 'Recorder signature',
      'ลายเซ็นนักกีฬา': 'Athlete signature',
      'พิมพ์ชื่อ ถ้าไม่เซ็นระบบจะใช้ชื่อนี้': 'Type name; used if no signature is drawn',
      'เซ็นชื่อเต็มหน้าจอ': 'Full-screen signature',
      'ล้างลายเซ็น': 'Clear signature',
      'แตะปุ่มเพื่อเปิดพื้นที่เซ็นเต็มหน้าจอ': 'Tap the button to open the full-screen signing area',
      'รหัสข้ามการลงชื่อ (admin)': 'Signature bypass code (admin)',
      'กรอกเมื่อจำเป็น': 'Enter only when needed',
      'จบการตี': 'Finish attempt',
      'ลงลายเซ็น': 'Sign',
      'เซ็นเต็มหน้าจอ แล้วกดบันทึกเพื่อกลับไปยังฟอร์ม': 'Sign in full screen, then save to return to the form',
      'ปิด': 'Close',
      'ล้าง': 'Clear',
      'พิมพ์': 'Print',
      'ลำดับ': 'Rank',
      'ลายเซ็น / ผู้เกี่ยวข้องแต่ละรอบ': 'Signatures / officials by round',
      'รอบ': 'Round',
      'กรรมการ': 'Umpire',
      'ผู้บันทึก': 'Recorder',
      'กลับหน้าเลือกนักกีฬา': 'Back to athlete selection',
      'พิมพ์ทั้งหมดที่แสดง': 'Print all displayed',
      'จำนวน': 'Count',
      'คน': 'people',
      'ยังไม่มีนักกีฬาสำหรับพิมพ์ในรอบนี้': 'No athletes to print in this round',
      'เลือกพิมพ์ Scorecard': 'Select scorecards to print',
      'เลือกรอบ': 'Select round',
      'เปลี่ยนรอบ': 'Change round',
      'เลือกทั้งหมด': 'Select all',
      'ยกเลิกทั้งหมด': 'Clear all',
      'ยังไม่มีนักกีฬาในรอบนี้': 'No athletes in this round',
      'พิมพ์คนที่เลือก': 'Print selected athletes',
      'พิมพ์ทั้งหมดในรอบนี้': 'Print all in this round',
      'คะแนนสูงสุดของประเภทนี้': 'Highest score in this category',
      'อันดับ': 'Rank',
      'ครั้ง': 'times',
      'ตีเที่ยวพิเศษ 7 เมตรทุกสถานี': 'Special tie-break: 7 meters at every station',
      'เที่ยวพิเศษ': 'Tie-break',
      'บันทึกเที่ยวพิเศษ': 'Save tie-break',
      'กดที่ชื่อในแต่ละ match box เพื่อเปิด scorecard ของรอบนั้น': 'Click a name in each match box to open that round scorecard',
      'ยังไม่มีคู่ Quarter Final': 'No Quarter Final matches yet',
      'ยังไม่มีคู่ Semi Final': 'No Semi Final matches yet',
      'ยังไม่มีคู่ Final': 'No Final matches yet',
      'ยืนยันลบ': 'Confirm delete',
      'ลบอีเวนต์นี้?': 'Delete this event?',

      /* Navigation and event pages */
      'และ': 'and',
      '+ สร้างอีเวนต์': '+ Create event',
      'กลับหน้าอีเวนต์': 'Back to event',
      'รายการทั้งหมด': 'All entries',
      'ประวัติรอบคัดเลือก': 'Qualification history',
      'ผู้แข่งขัน': 'Competitors',
      '🖨 พิมพ์ Score Sheet รอบนี้': '🖨 Print score sheets for this round',
      'ยังไม่มีคู่รอบ 16 คน': 'No Round of 16 matches yet',
      'บันทึกภาพไม่สำเร็จ ให้ใช้ปุ่มพิมพ์แล้วเลือก Save as PDF': 'The image could not be saved. Use Print and select Save as PDF instead.',
      'ไม่สามารถเปิดเครื่องมือบันทึกภาพได้ กรุณารีเฟรชหน้าแล้วลองใหม่': 'The image export tool is unavailable. Refresh the page and try again.',
      'กำลังสร้างภาพ…': 'Creating image…',
      'รอแข่งขัน': 'Waiting to compete',
      'ระบบใช้ตามรอบถัดไปอัตโนมัติ: รอบ 8 = 4, รอบ 16 = 8': 'The system sets this automatically from the next round: Quarterfinal = 4, Round of 16 = 8.',
      'รอบ 16 คน': 'Round of 16',

      /* Overview and ranking */
      'ตั้งค่า Results': 'Results settings',
      '📊 ดูสถิติ 5/3': '📊 View 5/3 statistics',
      'ยังไม่ตี': 'Not started',
      'กำลังตี / มีการคีย์': 'In progress / scores entered',
      'ตีแล้ว': 'Done',
      'ผ่านไปตีรอบ 2': 'Advanced to Round 2',
      'เข้า Knockout': 'Advanced to the knockout stage',
      'กำลังตี': 'Shooting',
      'เข้ารอบ': 'Qualified',
      'จบ': 'Done',
      'รอ': 'Waiting',
      'KO': 'KO',
      'สถิติประกอบการจัดอันดับ 5/3': '5/3 ranking statistics',
      'เรียงตาม TOTAL → จำนวน 5 → จำนวน 3 ใช้ดูเหตุผลการจัดลำดับ ไม่รวมคะแนนรอบพิเศษในตารางหลัก': 'Ranked by TOTAL, then the number of 5-point shots, then the number of 3-point shots. This explains the ranking; tie-break scores are not included in the main table.',
      'กำลังโหลด...': 'Loading...',
      'โหลดข้อมูลไม่ได้': 'Unable to load data',
      'ไม่มีข้อมูล': 'No data available',
      'เส้นตัดเข้ารอบ': 'Qualification cut line',
      'CUT LINE / เส้นตัดเข้ารอบ': 'QUALIFICATION CUT LINE',
      'แก้': 'Edit',
      'ต้องตี Sudden death': 'Shoot-off required',
      'คะแนนเสมอ': 'Tied score',
      'ตีรอบ Sudden death': 'Shoot-off',
      'พร้อมบันทึก': 'Ready to save',
      'กำลังบันทึก…': 'Saving…',
      'บันทึกแล้ว': 'Saved',
      'บันทึกไม่สำเร็จ กรุณาลองอีกครั้ง': 'Save failed. Please try again.',
      'เซิร์ฟเวอร์ตอบกลับไม่ถูกต้อง': 'The server returned an invalid response.',

      /* Results Approved */
      '← กลับ Overview': '← Back to Overview',
      '⚙ ตั้งค่า Results Approved / กรรมการ': '⚙ Results Approved / Officials settings',
      'ดาวน์โหลด DOCX': 'Download DOCX',
      'พิมพ์ / Save PDF': 'Print / Save as PDF',
      'ตั้งค่า Results Approved': 'Results Approved settings',
      'ตั้งค่า Results Approved SEA Games Style': 'Configure Results Approved in SEA Games style',
      'ตั้งค่าข้อความหัวเอกสารและรายชื่อกรรมการสำหรับ': 'Configure the document header and officials for',
      'ชื่อการแข่งขัน / Competition Title': 'Competition title',
      'บรรทัดประเทศ/เจ้าภาพ': 'Country / host line',
      'วันที่บนเอกสาร': 'Document date',
      'สถานที่ / Location': 'Location',
      'หัวคอลัมน์สังกัด': 'Affiliation column heading',
      'ข้อความ APPROVED': 'APPROVED text',
      'โลโก้เอกสาร': 'Document logos',
      'อัปโหลดไฟล์ png, jpg, jpeg, webp หรือ gif ได้เอง ถ้าไม่อัปโหลดระบบจะใช้โลโก้ค่าเริ่มต้นตามเอกสารตัวอย่าง': 'Upload a PNG, JPG, JPEG, WEBP, or GIF file. If no file is uploaded, the default logo from the sample document will be used.',
      'ใช้ค่าเริ่มต้น': 'Use default',
      'แสดงหน้ากรรมการ': 'Show the officials page',
      'แสดง': 'Show',
      'ไม่แสดง': 'Hide',
      'ตำแหน่งประธาน / President Title': 'President title',
      'ชื่อประธาน / President Name': 'President name',
      'ตำแหน่ง Technical Delegate': 'Technical Delegate title',
      'ชื่อ Technical Delegate': 'Technical Delegate name',
      'รายชื่อกรรมการ Umpire': 'Umpire list',
      'ใส่บรรทัดละ 1 คน รูปแบบ: ชื่อ | สังกัด เช่น MR. SOMCHAI JAIDEE | ABSC': 'Enter one person per line in this format: Name | Affiliation, for example MR. SOMCHAI JAIDEE | ABSC.',
      'บันทึกตั้งค่า': 'Save settings',
      'กลับ Results Approved': 'Back to Results Approved',
      'กลับ Overview': 'Back to Overview',

      /* Scorecard and signatures */
      'กรอกทีละสถานี': 'Enter scores one station at a time',
      'สถานี': 'Station',
      'มือถือ/แท็บเล็ตแนวตั้งจะแสดงเฉพาะสถานีที่เลือก เพื่อให้ช่องใหญ่ขึ้นและกดง่ายขึ้น': 'On phones and portrait tablets, only the selected station is shown so the score boxes are larger and easier to tap.',
      'ยิงลูกเดี่ยว ยิงลูกเหนือเป้า ยิงลูกกลางขวาง ยิงลูกขาวเหนือลูกดำ ยิงลูกเป้า': 'Single ball shot, target-ball jump shot, middle obstacle shot, white-over-black ball shot, and target ball shot',
      'ลายเซ็นเฉพาะ': 'Signatures for',
      'บันทึกลายเซ็น': 'Save signature',
      'เลือกสถานีสำหรับหน้าจอแนวตั้ง': 'Select a station for portrait screens',
      'รอบที่': 'Round',
      'ลายเซ็นกรรมการตัดสิน -': 'Umpire signature -',
      'ลายเซ็นกรรมการบันทึกคะแนน -': 'Recorder signature -',
      'ลายเซ็นนักกีฬา -': 'Athlete signature -',
      '← กลับหน้าเลือกนักกีฬา': '← Back to athlete selection',
      '← กลับ scorecard': '← Back to scorecard',
      'ลายเซ็นกรรมการ': 'Umpire signature',
      'ลายเซ็นผู้บันทึก': 'Recorder signature',
      'ใบแดงครบ 2 ครั้ง: ยุติการยิงรอบนี้และคงคะแนนที่ทำได้ไว้': 'Two red cards: this shooting round has ended. All points already scored are retained.',

      /* Statistics and tie-break */
      'สถิติ —': 'Statistics —',
      '5 คะแนน': '5 points',
      '3 คะแนน': '3 points',
      'รอบ · เลือกคะแนนแบบ 0 / 1 / 3 / 5': 'round · Select a score of 0, 1, 3, or 5',
      'ต้องบันทึกพร้อมกันทุกคนที่ติด Sudden death หากยังเท่ากันให้กลับมาเลือกกลุ่มเดิมแล้วบันทึกเพิ่มอีกเที่ยว': 'All athletes involved in the shoot-off must be saved together. If they are still tied, select the same group and record another shoot-off.',
      'สะสมเดิม:': 'Previous total:',
      'บันทึกเที่ยวพิเศษพร้อมกัน': 'Save the shoot-off for all selected athletes',

      /* Messages and validation */
      'คุณไม่มีสิทธิ์ทำรายการนี้': 'You do not have permission to perform this action.',
      'รอคู่แข่งขัน': 'Waiting for opponent',
      'แข่งเสร็จ': 'Match completed',
      'กำลังแข่งขัน': 'Match in progress',
      'รอบันทึกผู้ชนะ': 'Waiting for the winner to be recorded',
      'เข้าสู่ระบบสำเร็จ': 'You have logged in successfully.',
      'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง': 'The username or password is incorrect.',
      'ออกจากระบบแล้ว': 'You have logged out.',
      'สร้างผู้ใช้สำเร็จ': 'The user was created successfully.',
      'สร้างผู้ใช้ไม่สำเร็จ กรุณาตรวจสอบข้อมูล': 'The user could not be created. Please check the information and try again.',
      'สร้างอีเวนต์สำเร็จ': 'The event was created successfully.',
      'แก้ไขอีเวนต์สำเร็จ': 'The event was updated successfully.',
      'ลบอีเวนต์แล้ว': 'The event was deleted.',
      'เพิ่มนักกีฬาสำเร็จ': 'The athlete was added successfully.',
      'ลบรายการ': 'Deleted',
      'เรียบร้อย': 'successfully.',
      'กรุณาเลือกไฟล์ Excel': 'Please select an Excel file.',
      'รองรับเฉพาะไฟล์ .xlsx': 'Only .xlsx files are supported.',
      'นำเข้านักกีฬาสำเร็จ': 'Athletes imported successfully:',
      'สุ่มลำดับใหม่แล้ว': 'The shooting order was randomized successfully.',
      'นักกีฬาคนนี้ไม่มีสิทธิ์ตีรอบ 2': 'This athlete is not eligible to compete in Round 2.',
      'รอบแข่งขันไม่ถูกต้อง': 'The competition round is invalid.',
      'จบการตีเรียบร้อย': 'The attempt was completed successfully.',
      'ต้องลงชื่ออย่างใดอย่างหนึ่ง (พิมพ์ชื่อหรือเขียน) ให้ครบทั้ง 3 ฝ่าย หรือใช้สิทธิ์ข้าม': 'All three parties must provide a signature, either by typing a name or drawing a signature, unless an authorized signature bypass is used.',
      'กรุณาเลือกนักกีฬาอย่างน้อย 1 คน': 'Please select at least one athlete.',
      'บันทึกผลเที่ยวพิเศษแล้ว ถ้ายังเท่ากันให้บันทึกเที่ยวพิเศษเพิ่มอีกครั้ง': 'The tie-break result was saved. If the athletes are still tied, record another tie-break.',
      'กรุณาเลือกนักกีฬาที่ต้องตี Sudden death อย่างน้อย 2 คน': 'Please select at least two athletes for the shoot-off.',
      'บันทึก Sudden death พร้อมกันแล้ว ถ้ายังเท่ากันให้เลือกกลุ่มเดิมแล้วบันทึกเพิ่ม': 'The shoot-off was saved for all selected athletes. If they are still tied, select the same group and record another shoot-off.',
      'ผู้ชนะไม่ถูกต้อง': 'The selected winner is invalid.',
      'บันทึกผู้ชนะแล้ว': 'The winner was saved successfully.',
      'ต้องเป็น png, jpg, jpeg, webp หรือ gif': 'must be a PNG, JPG, JPEG, WEBP, or GIF file.',
      'บันทึกตั้งค่า Results Approved แล้ว': 'The Results Approved settings were saved successfully.',
      'ยังไม่ได้ติดตั้ง python-docx ให้รัน: python -m pip install -r requirements.txt': 'python-docx is not installed. Run: python -m pip install -r requirements.txt',
      'ไฟล์ Excel ต้องมีหัวคอลัมน์ชื่อ และ สังกัด': 'The Excel file must contain Name and Affiliation columns.',
      'ต้องมีทั้งชื่อและสังกัด': 'must include both a name and an affiliation.',
      'ไม่พบข้อมูลนักกีฬาในไฟล์ Excel': 'No athlete data was found in the Excel file.'
    },
    fr: {
      "แผงกดคะแนน": "Pavé de score",
      "สถานีนี้": "Cet atelier",
      "รวมรอบ": "Total du tour",
      "รับช่วงคีย์ต่อ": "Reprendre la saisie",
      "← คิวสนาม": "← File du terrain",
      "คิวสนาม": "File du terrain",
      "โหมดจอแสดงผล": "Mode affichage",
      "เครื่องมือ ▾": "Outils ▾",
      "ตีจบแล้ว": "Terminé",
      "รอคิว": "En attente",
      "คนถัดไป": "Suivant",
      "รับรองผลแล้ว": "Validé",
      "คะแนน": "points",
      "เริ่มคีย์": "Commencer",
      "คีย์ต่อ": "Continuer",
      "ดู / แก้ไข": "Voir / modifier",
      'หน้าหลัก': 'Accueil', 'สร้างอีเวนต์': 'Créer un événement', 'สร้างอีเวนต์ Shooting': 'Créer un événement de tir', 'แก้ไขอีเวนต์': 'Modifier l’événement', 'ผู้ใช้': 'Utilisateurs', 'ออกจากระบบ': 'Déconnexion', 'เข้าสู่ระบบ': 'Connexion', 'ชื่อผู้ใช้': 'Nom d’utilisateur', 'รหัสผ่าน': 'Mot de passe', 'สิทธิ์': 'Rôle', 'สร้างผู้ใช้ใหม่': 'Créer un utilisateur', 'รายการผู้ใช้': 'Liste des utilisateurs', 'บันทึก': 'Enregistrer', 'บันทึกการแก้ไข': 'Enregistrer les modifications',
      'แดชบอร์ด': 'Tableau de bord', 'ภาพรวมของรายการแข่งขันและนักกีฬาในระบบ': 'Vue d’ensemble des compétitions et des athlètes', 'อีเวนต์': 'Événements', 'รายการแข่งขัน': 'Compétitions', 'เลือกดูหน้ารวม สถิติ และจัดการนักกีฬาได้จากตารางนี้': 'Ouvrir l’aperçu, les statistiques et la gestion des athlètes depuis ce tableau', 'คะแนนสูงสุดในระบบ': 'Meilleur score du système', 'สถิติสังกัดที่ทำคะแนนดีที่สุด': 'Statistiques des clubs/affiliations les mieux notés', 'คะแนนสูงสุด': 'Meilleur score', 'คะแนนรวม': 'Score total', 'คะแนน': 'points',
      'ชื่องาน': 'Nom de l’événement', 'รุ่น': 'Groupe', 'รุ่นแข่งขัน': 'Groupe de compétition', 'ประเภท': 'Catégorie', 'วันแข่งขัน': 'Date de compétition', 'สนาม': 'Terrain', 'จำนวนสนาม': 'Nombre de terrains', 'จัดการ': 'Gérer', 'หน้ารวม': 'Aperçu', 'สถิติ': 'Statistiques', 'นักกีฬา': 'Athlètes', 'แก้ไข': 'Modifier', 'ลบ': 'Supprimer', 'ยังไม่มีอีเวนต์': 'Aucun événement', 'ยังไม่มีข้อมูล': 'Aucune donnée',
      'ทั่วไป': 'Open', 'อาวุโส': 'Senior', 'เยาวชน': 'Jeunes', 'ชาย': 'Hommes', 'หญิง': 'Femmes', 'ผสม': 'Mixte', 'มี': 'Oui', 'ไม่มี': 'Non', 'จากรอบแรกเข้ารอบถัดไปทันที กี่คน': 'Qualifiés directs depuis le 1er tour', 'มีรอบ 2 หรือไม่': 'Activer le 2e tour ?', 'ถ้ามีรอบ 2 ให้ตีถึงลำดับที่เท่าไหร่': 'Rang limite du 2e tour', 'จากรอบ 2 คัดเข้ารอบต่อกี่คน': 'Qualifiés depuis le 2e tour', 'รอบถัดไปคือรอบอะไร': 'Nom du tour suivant',
      'รอบที่กำลังดู': 'Tour affiché', 'รอบ 1': 'Tour 1', 'รอบ 2': 'Tour 2', 'รอบที่ 1': 'Tour 1', 'รอบที่ 2': 'Tour 2', 'รอบ 8 คน': 'Quart de finale', 'รอบ 4 คน': 'Demi-finale à 4', 'รอบรองชนะเลิศ': 'Demi-finale', 'รอบชิงชนะเลิศ': 'Finale', 'ตารางประกบ': 'Tableau', 'หน้าประกบคู่': 'Page du tableau', 'พิมพ์ Scorecard หลายคน': 'Imprimer plusieurs fiches', 'บันทึกเป็นรูปภาพ': 'Enregistrer comme image', 'พิมพ์ตาราง': 'Imprimer le tableau', 'ย่อ/ขยาย': 'Zoom', 'ความกว้างคู่': 'Largeur du match', 'ระยะระหว่างรอบ': 'Espacement des tours', 'เลือกรอบที่จะเลื่อน': 'Tour à déplacer', 'เลื่อนขึ้น–ลง': 'Déplacer haut/bas', 'พอดีหน้าจอ': 'Ajuster à l’écran', 'คืนค่า': 'Réinitialiser',
      'รอคิว': 'En attente', 'กำลังตี': 'En cours', 'ตีเสร็จแล้ว': 'Terminé', 'เข้ารอบตรง': 'Qualifié direct', 'ผ่านจากรอบ 2': 'Qualifié du 2e tour', 'มีสิทธิ์รอบ 2': 'Éligible au 2e tour', 'ตกรอบ': 'Éliminé', 'ลำดับการตี': 'Ordre de tir', 'ลำดับในสนาม': 'Ordre sur terrain', 'สถานะ': 'Statut', 'หมายเลข': 'Dossard', 'ชื่อ': 'Nom', 'สังกัด': 'Affiliation',
      'เพิ่มนักกีฬาและจัดลำดับยิง': 'Ajouter des athlètes et définir l’ordre de tir', 'ไปหน้ารวม': 'Aller à l’aperçu', 'หมายเลขและลำดับรวม ระบบสร้างให้อัตโนมัติ': 'Le dossard et l’ordre général sont générés automatiquement', 'เพิ่มนักกีฬา': 'Ajouter un athlète', 'การกระจายสนามอัตโนมัติ': 'Répartition automatique des terrains', 'สนาม = วนตามจำนวนสนาม เช่น 1, 2, 3, 4 แล้วกลับไป 1': 'Les terrains tournent selon le nombre disponible, ex. 1, 2, 3, 4 puis retour à 1', 'สุ่มลำดับทั้งหมด': 'Mélanger tous les ordres', 'นำเข้ารายชื่อนักกีฬาจาก Excel': 'Importer une liste depuis Excel', 'ไฟล์ต้องเป็น .xlsx และต้องมี 2 คอลัมน์คือ': 'Le fichier doit être .xlsx avec 2 colonnes :', 'เลือกไฟล์ Excel': 'Choisir un fichier Excel', 'อัปโหลด Excel': 'Importer Excel', 'ดาวน์โหลดไฟล์ตัวอย่าง': 'Télécharger le modèle', 'ยังไม่มีนักกีฬา': 'Aucun athlète',
      'ประเภทสุดยอดความแม่นยำ (SHOOTING)': 'Tir de précision', 'กลับหน้ารวม': 'Retour à l’aperçu', 'พิมพ์แบบฟอร์ม': 'Imprimer le formulaire', 'รอบที่คีย์': 'Tour saisi', 'สถานีที่': 'Atelier', 'ยิงลูกเดี่ยว': 'Tir boule seule', 'ยิงลูกเหนือเป้า': 'Tir par-dessus la cible', 'ยิงลูกกลางขวาง': 'Tir avec obstacle central', 'ยิงลูกขาวเหนือลูกดำ': 'Boule blanche par-dessus boule noire', 'ยิงลูกเป้า': 'Tir sur la cible', 'รอบแข่งขัน': 'Tour de compétition', 'รวม': 'Total', 'แดง': 'Rouge',
      'กรรมการตัดสิน': 'Arbitre', 'กรรมการบันทึกคะแนน': 'Marqueur', 'ลายเซ็นกรรมการตัดสิน': 'Signature de l’arbitre', 'ลายเซ็นกรรมการบันทึกคะแนน': 'Signature du marqueur', 'ลายเซ็นนักกีฬา': 'Signature de l’athlète', 'พิมพ์ชื่อ ถ้าไม่เซ็นระบบจะใช้ชื่อนี้': 'Saisir le nom ; utilisé si aucune signature n’est dessinée', 'เซ็นชื่อเต็มหน้าจอ': 'Signature plein écran', 'ล้างลายเซ็น': 'Effacer la signature', 'แตะปุ่มเพื่อเปิดพื้นที่เซ็นเต็มหน้าจอ': 'Touchez le bouton pour ouvrir la zone de signature plein écran', 'รหัสข้ามการลงชื่อ (admin)': 'Code de contournement signature (admin)', 'กรอกเมื่อจำเป็น': 'Saisir si nécessaire', 'จบการตี': 'Terminer le tir', 'ลงลายเซ็น': 'Signer', 'เซ็นเต็มหน้าจอ แล้วกดบันทึกเพื่อกลับไปยังฟอร์ม': 'Signer en plein écran puis enregistrer pour revenir au formulaire', 'ปิด': 'Fermer', 'ล้าง': 'Effacer', 'พิมพ์': 'Imprimer',
      'ลำดับ': 'Rang', 'ลายเซ็น / ผู้เกี่ยวข้องแต่ละรอบ': 'Signatures / officiels par tour', 'รอบ': 'Tour', 'กรรมการ': 'Arbitre', 'ผู้บันทึก': 'Marqueur', 'กลับหน้าเลือกนักกีฬา': 'Retour à la sélection', 'พิมพ์ทั้งหมดที่แสดง': 'Imprimer tout ce qui est affiché', 'จำนวน': 'Nombre', 'คน': 'personnes', 'ยังไม่มีนักกีฬาสำหรับพิมพ์ในรอบนี้': 'Aucun athlète à imprimer pour ce tour', 'เลือกพิมพ์ Scorecard': 'Choisir les fiches à imprimer', 'เลือกรอบ': 'Choisir le tour', 'เปลี่ยนรอบ': 'Changer de tour', 'เลือกทั้งหมด': 'Tout sélectionner', 'ยกเลิกทั้งหมด': 'Tout désélectionner', 'ยังไม่มีนักกีฬาในรอบนี้': 'Aucun athlète dans ce tour', 'พิมพ์คนที่เลือก': 'Imprimer la sélection', 'พิมพ์ทั้งหมดในรอบนี้': 'Imprimer tout ce tour',
      'คะแนนสูงสุดของประเภทนี้': 'Meilleur score de cette catégorie', 'อันดับ': 'Classement', 'ครั้ง': 'fois', 'ตีเที่ยวพิเศษ 7 เมตรทุกสถานี': 'Départage spécial : 7 m à chaque atelier', 'เที่ยวพิเศษ': 'Départage', 'บันทึกเที่ยวพิเศษ': 'Enregistrer le départage', 'กดที่ชื่อในแต่ละ match box เพื่อเปิด scorecard ของรอบนั้น': 'Cliquez sur un nom dans chaque case pour ouvrir la fiche du tour', 'ยังไม่มีคู่ Quarter Final': 'Aucun quart de finale', 'ยังไม่มีคู่ Semi Final': 'Aucune demi-finale', 'ยังไม่มีคู่ Final': 'Aucune finale', 'ยืนยันลบ': 'Confirmer la suppression', 'ลบอีเวนต์นี้?': 'Supprimer cet événement ?'
    },
    zh: {
      "แผงกดคะแนน": "录分键盘",
      "สถานีนี้": "本站",
      "รวมรอบ": "本轮合计",
      "แตะตัวเลขเพื่อบันทึกทันที · ครบ 4 ระยะแล้วจะไปสถานีถัดไปเอง": "点击数字即时保存 · 4 个距离完成后自动进入下一站",
      "ใบนี้กำลังถูกคีย์โดย": "此记分卡正在由以下用户录入：",
      "เพื่อกันคะแนนทับกัน หน้านี้จึงดูได้อย่างเดียว ถ้าเครื่องเดิมเลิกใช้แล้วให้กดรับช่วงคีย์ต่อ": "为防止比分被覆盖，本页只读。如原设备已不再使用，请接管录入。",
      "รับช่วงคีย์ต่อ": "接管录入",
      "← คิวสนาม": "← 场地队列",
      "คิวสนาม": "场地队列",
      "โหมดจอแสดงผล": "大屏模式",
      "เครื่องมือ ▾": "工具 ▾",
      "หน้ารวมทุกสนาม": "所有场地总览",
      "ตีจบแล้ว": "已完成",
      "รอคิว": "等待中",
      "คนถัดไป": "下一位",
      "รับรองผลแล้ว": "已确认",
      "คะแนน": "分",
      "ดู / แก้ไข": "查看 / 编辑",
      "คีย์ต่อ": "继续录入",
      "เริ่มคีย์": "开始录入",
      "หน้านี้อัปเดตเองทุก 15 วินาที": "本页每 15 秒自动刷新",
      "ยังไม่มีนักกีฬาในสนามนี้": "该场地暂无运动员",
      "ยังไม่มีรายชื่อรอบ 2 ในสนามนี้ (รายชื่อจะขึ้นหลังรอบ 1 จบครบทุกสนาม)": "该场地暂无第2轮名单（所有场地第1轮结束后显示）",
      "บันทึกเป็น JPG": "保存为 JPG",
      "พร้อมบันทึก": "就绪",
      "กำลังบันทึก…": "保存中…",
      "บันทึกแล้ว": "已保存",
      "Court ID ของอีเวนต์นี้": "本赛事场地账号",
      "ยังใช้รหัสตั้งต้น – เข้าระบบไม่ได้จนกว่าจะตั้งใหม่": "仍在使用默认密码 – 重设前无法登录",
      "ตั้งรหัสผ่านใหม่": "设置新密码",
      "ตั้งรหัส": "设置",
      "รหัสใหม่": "新密码",
      "รหัสผ่าน (อย่างน้อย 8 ตัว)": "密码（至少 8 位）",
      "ผลรอบนี้ได้รับการรับรองแล้ว": "本轮成绩已确认",
      "กรรมการยกคะแนน": "计分裁判",
      "📊 ดูสถิติ 5/3": "📊 查看 5/3 统计",
      "สถิติประกอบการจัดอันดับ 5/3": "排名参考统计 5/3",
      "เรียงตาม TOTAL → จำนวน 5 → จำนวน 3 ใช้ดูเหตุผลการจัดลำดับ ไม่รวมคะแนนรอบพิเศษในตารางหลัก": "按 TOTAL → 5分次数 → 3分次数排序，用于查看排名依据，主表不含加赛分",
      "ใบแดงครบ 2 ครั้ง: ยุติการยิงรอบนี้และคงคะแนนที่ทำได้ไว้": "红牌满 2 次：本轮停止射击并保留已得分",
      "มือถือ/แท็บเล็ตแนวตั้งจะแสดงเฉพาะสถานีที่เลือก เพื่อให้ช่องใหญ่ขึ้นและกดง่ายขึ้น": "手机/平板竖屏只显示所选站点，格子更大更易点击",
      'หน้าหลัก': '首页', 'สร้างอีเวนต์': '创建赛事', 'สร้างอีเวนต์ Shooting': '创建射击赛事', 'แก้ไขอีเวนต์': '编辑赛事', 'ผู้ใช้': '用户', 'ออกจากระบบ': '退出登录', 'เข้าสู่ระบบ': '登录', 'ชื่อผู้ใช้': '用户名', 'รหัสผ่าน': '密码', 'สิทธิ์': '权限', 'สร้างผู้ใช้ใหม่': '创建新用户', 'รายการผู้ใช้': '用户列表', 'บันทึก': '保存', 'บันทึกการแก้ไข': '保存修改',
      'แดชบอร์ด': '仪表盘', 'ภาพรวมของรายการแข่งขันและนักกีฬาในระบบ': '系统中的赛事和运动员概览', 'อีเวนต์': '赛事', 'รายการแข่งขัน': '比赛列表', 'เลือกดูหน้ารวม สถิติ และจัดการนักกีฬาได้จากตารางนี้': '可从此表进入总览、统计和运动员管理', 'คะแนนสูงสุดในระบบ': '系统最高分', 'สถิติสังกัดที่ทำคะแนนดีที่สุด': '最佳单位/队伍统计', 'คะแนนสูงสุด': '最高分', 'คะแนนรวม': '总分', 'คะแนน': '分',
      'ชื่องาน': '赛事名称', 'รุ่น': '组别', 'รุ่นแข่งขัน': '比赛组别', 'ประเภท': '类别', 'วันแข่งขัน': '比赛日期', 'สนาม': '场地', 'จำนวนสนาม': '场地数量', 'จัดการ': '管理', 'หน้ารวม': '总览', 'สถิติ': '统计', 'นักกีฬา': '运动员', 'แก้ไข': '编辑', 'ลบ': '删除', 'ยังไม่มีอีเวนต์': '暂无赛事', 'ยังไม่มีข้อมูล': '暂无数据',
      'ทั่วไป': '公开组', 'อาวุโส': '长青组', 'เยาวชน': '青年组', 'ชาย': '男子', 'หญิง': '女子', 'ผสม': '混合', 'มี': '有', 'ไม่มี': '无', 'จากรอบแรกเข้ารอบถัดไปทันที กี่คน': '第一轮直接晋级人数', 'มีรอบ 2 หรือไม่': '是否有第二轮？', 'ถ้ามีรอบ 2 ให้ตีถึงลำดับที่เท่าไหร่': '第二轮截止名次', 'จากรอบ 2 คัดเข้ารอบต่อกี่คน': '第二轮晋级人数', 'รอบถัดไปคือรอบอะไร': '下一轮名称',
      'รอบที่กำลังดู': '当前查看轮次', 'รอบ 1': '第1轮', 'รอบ 2': '第2轮', 'รอบที่ 1': '第1轮', 'รอบที่ 2': '第2轮', 'รอบ 8 คน': '八强赛', 'รอบ 4 คน': '四强赛', 'รอบรองชนะเลิศ': '半决赛', 'รอบชิงชนะเลิศ': '决赛', 'ตารางประกบ': '对阵表', 'หน้าประกบคู่': '对阵页面', 'พิมพ์ Scorecard หลายคน': '打印多张记分卡', 'บันทึกเป็นรูปภาพ': '保存为图片', 'พิมพ์ตาราง': '打印表格', 'ย่อ/ขยาย': '缩放', 'ความกว้างคู่': '对阵框宽度', 'ระยะระหว่างรอบ': '轮次间距', 'เลือกรอบที่จะเลื่อน': '选择要移动的轮次', 'เลื่อนขึ้น–ลง': '上下移动', 'พอดีหน้าจอ': '适合屏幕', 'คืนค่า': '重置',
      'รอคิว': '等待中', 'กำลังตี': '进行中', 'ตีเสร็จแล้ว': '已完成', 'เข้ารอบตรง': '直接晋级', 'ผ่านจากรอบ 2': '第二轮晋级', 'มีสิทธิ์รอบ 2': '可进入第二轮', 'ตกรอบ': '淘汰', 'ลำดับการตี': '击球顺序', 'ลำดับในสนาม': '场地内顺序', 'สถานะ': '状态', 'หมายเลข': '号码', 'ชื่อ': '姓名', 'สังกัด': '所属单位',
      'เพิ่มนักกีฬาและจัดลำดับยิง': '添加运动员并安排击球顺序', 'ไปหน้ารวม': '前往总览', 'หมายเลขและลำดับรวม ระบบสร้างให้อัตโนมัติ': '号码和总顺序由系统自动生成', 'เพิ่มนักกีฬา': '添加运动员', 'การกระจายสนามอัตโนมัติ': '自动分配场地', 'สนาม = วนตามจำนวนสนาม เช่น 1, 2, 3, 4 แล้วกลับไป 1': '场地按数量循环分配，例如 1、2、3、4 后回到 1', 'สุ่มลำดับทั้งหมด': '随机全部顺序', 'นำเข้ารายชื่อนักกีฬาจาก Excel': '从 Excel 导入运动员名单', 'ไฟล์ต้องเป็น .xlsx และต้องมี 2 คอลัมน์คือ': '文件必须为 .xlsx，且包含两列：', 'เลือกไฟล์ Excel': '选择 Excel 文件', 'อัปโหลด Excel': '上传 Excel', 'ดาวน์โหลดไฟล์ตัวอย่าง': '下载模板文件', 'ยังไม่มีนักกีฬา': '暂无运动员',
      'ประเภทสุดยอดความแม่นยำ (SHOOTING)': '精准射击', 'กลับหน้ารวม': '返回总览', 'พิมพ์แบบฟอร์ม': '打印表单', 'รอบที่คีย์': '录入轮次', 'สถานีที่': '站点', 'ยิงลูกเดี่ยว': '单球射击', 'ยิงลูกเหนือเป้า': '越过目标球射击', 'ยิงลูกกลางขวาง': '中间障碍射击', 'ยิงลูกขาวเหนือลูกดำ': '白球越过黑球射击', 'ยิงลูกเป้า': '目标球射击', 'รอบแข่งขัน': '比赛轮次', 'รวม': '合计', 'แดง': '红牌',
      'กรรมการตัดสิน': '裁判', 'กรรมการบันทึกคะแนน': '记录员', 'ลายเซ็นกรรมการตัดสิน': '裁判签名', 'ลายเซ็นกรรมการบันทึกคะแนน': '记录员签名', 'ลายเซ็นนักกีฬา': '运动员签名', 'พิมพ์ชื่อ ถ้าไม่เซ็นระบบจะใช้ชื่อนี้': '输入姓名；未签名时使用此姓名', 'เซ็นชื่อเต็มหน้าจอ': '全屏签名', 'ล้างลายเซ็น': '清除签名', 'แตะปุ่มเพื่อเปิดพื้นที่เซ็นเต็มหน้าจอ': '点击按钮打开全屏签名区域', 'รหัสข้ามการลงชื่อ (admin)': '跳过签名代码（管理员）', 'กรอกเมื่อจำเป็น': '必要时填写', 'จบการตี': '结束击球', 'ลงลายเซ็น': '签名', 'เซ็นเต็มหน้าจอ แล้วกดบันทึกเพื่อกลับไปยังฟอร์ม': '全屏签名后点击保存返回表单', 'ปิด': '关闭', 'ล้าง': '清除', 'พิมพ์': '打印',
      'ลำดับ': '名次', 'ลายเซ็น / ผู้เกี่ยวข้องแต่ละรอบ': '每轮签名 / 相关人员', 'รอบ': '轮次', 'กรรมการ': '裁判', 'ผู้บันทึก': '记录员', 'กลับหน้าเลือกนักกีฬา': '返回运动员选择页', 'พิมพ์ทั้งหมดที่แสดง': '打印当前显示全部', 'จำนวน': '数量', 'คน': '人', 'ยังไม่มีนักกีฬาสำหรับพิมพ์ในรอบนี้': '本轮暂无可打印运动员', 'เลือกพิมพ์ Scorecard': '选择打印记分卡', 'เลือกรอบ': '选择轮次', 'เปลี่ยนรอบ': '切换轮次', 'เลือกทั้งหมด': '全选', 'ยกเลิกทั้งหมด': '取消全选', 'ยังไม่มีนักกีฬาในรอบนี้': '本轮暂无运动员', 'พิมพ์คนที่เลือก': '打印所选', 'พิมพ์ทั้งหมดในรอบนี้': '打印本轮全部',
      'คะแนนสูงสุดของประเภทนี้': '该类别最高分', 'อันดับ': '排名', 'ครั้ง': '次', 'ตีเที่ยวพิเศษ 7 เมตรทุกสถานี': '加赛：每个站点 7 米', 'เที่ยวพิเศษ': '加赛', 'บันทึกเที่ยวพิเศษ': '保存加赛', 'กดที่ชื่อในแต่ละ match box เพื่อเปิด scorecard ของรอบนั้น': '点击对阵框中的姓名打开该轮记分卡', 'ยังไม่มีคู่ Quarter Final': '暂无四分之一决赛对阵', 'ยังไม่มีคู่ Semi Final': '暂无半决赛对阵', 'ยังไม่มีคู่ Final': '暂无决赛对阵', 'ยืนยันลบ': '确认删除', 'ลบอีเวนต์นี้?': '删除此赛事？'
    }
  };


  /* เพิ่มคำแปลหน้าใหม่ (Results Approved, ซ่อนอีเวนต์, ผังรอบ 2, คีย์คู่, ธีม ฯลฯ) — ใส่เฉพาะคำที่ยังไม่มี */
  const EXTRA_TRANSLATIONS = {
 "en": {
  "ชื่อผู้ใช้ / สิทธิ์": "Username / role",
  "บัญชีที่กำลังใช้อยู่": "Current account",
  "ลบบัญชีที่กำลังใช้อยู่ไม่ได้": "You cannot delete the account you are using",
  "ต้องมี superadmin อย่างน้อย 1 บัญชี": "At least 1 superadmin account is required",
  "กรุณาใส่ชื่อผู้ใช้": "Please enter a username",
  "เปลี่ยนสิทธิ์ของบัญชีที่กำลังใช้อยู่ไม่ได้": "You cannot change the role of the account you are using",
  "ไม่พบผู้ใช้ (บัญชีสนามให้จัดการจากหน้าอีเวนต์)": "User not found (manage court accounts from the event page)",
  "รูปแบบการคีย์": "Entry layout",
  "📋 ตาราง (แบบเดิม)": "📋 Table (classic)",
  "👆 ปุ่มแตะ (แบบรอบน็อค)": "👆 Tap buttons (knockout style)",
  "รวมรอบนี้": "Round total",
  "สถานีสุดท้าย": "Last station",
  "สถานีถัดไป →": "Next station →",
  "ปิดการเรียงสดแล้ว · แถวอยู่ตามลำดับการตี": "Live sorting off · rows follow shooting order",
  "กรุณาระบุเหตุผลในการล้างการคีย์": "Please give a reason for clearing the entries",
  "กรุณาเพิ่มอย่างน้อย 1 อีเวนต์": "Please add at least 1 event",
  "กรุณาเลือกอีเวนต์": "Please select events",
  "กรุณาเลือกอีเวนต์ที่จะจับสลาก": "Please select events to draw",
  "กรุณาเลือกอีเวนต์อย่างน้อย 1 รายการ": "Please select at least 1 event",
  "กรุณาเลือกไฟล์": "Please choose a file",
  "ข้ามอีเวนต์ที่เริ่มกรอกคะแนนแล้ว:": "Skipped events that already have scores:",
  "คะแนนเท่ากัน · ตีเที่ยวพิเศษ 7 เมตรทุกสถานี": "Tie · extra round at 7 m on every station",
  "คู่นี้ยังไม่ครบ 2 คน": "This match does not have 2 athletes yet",
  "จบการตีคู่นี้แล้ว": "Match finished",
  "ธีมที่มากับระบบลบไม่ได้ (แก้ไขหรือคัดลอกได้)": "Built-in themes cannot be deleted (you can edit or copy them)",
  "บัญชีนี้บันทึกบทลงโทษได้เฉพาะนักกีฬาในสนามของตัวเอง": "This account can only record penalties for athletes on its own court",
  "บัญชีนี้ยังใช้รหัสผ่านตั้งต้นที่เปิดเผยในโค้ด จึงถูกระงับ ให้ superadmin ตั้งรหัสใหม่ที่หน้าผู้ใช้": "This account still uses a default password from the code and is blocked; a superadmin must set a new one on the Users page",
  "บัญชีสนามกรอกเที่ยวพิเศษได้เฉพาะคู่ Knockout ในสนามตัวเอง": "Court accounts can only enter extra rounds for knockout matches on their own court",
  "บัญชีสนามนี้คีย์ได้เฉพาะคู่ในสนามตัวเอง": "This court account can only enter matches on its own court",
  "บัญชีสนามให้สร้างจากหน้านักกีฬาของอีเวนต์": "Create court accounts from the event’s athletes page",
  "บันทึก Sudden death แล้ว": "Sudden death saved",
  "พิมพ์คำว่า ล้าง ในช่องยืนยันก่อน": "Type ล้าง in the confirmation box first",
  "มี court ID นี้มากกว่า 1 อีเวนต์ที่ใช้รหัสเดียวกัน กรุณาเปลี่ยนรหัสของอีเวนต์ให้ไม่ซ้ำ": "This court ID uses the same password in more than one event; please make them different",
  "มีอีเวนต์ที่เริ่มกรอกคะแนนแล้ว การจัดสนามใหม่จะย้ายสนามของนักกีฬา · ติ๊กยืนยันก่อนบันทึก": "Some events already have scores; reassigning moves athletes between courts · tick confirm before saving",
  "ยังยิงไม่ครบทุกสถานี จึงยัง Approve ไม่ได้": "Not all stations are complete, so it cannot be approved yet",
  "ยังไม่ได้ติ๊กหัวข้อการตั้งค่าที่ต้องการใช้ร่วมกัน": "No shared setting was ticked",
  "ยังไม่ได้เลือกประเภทที่จะสร้าง": "No category selected to create",
  "รหัสข้ามลายเซ็นไม่ถูกต้อง · หรือไปลงลายเซ็นที่ Scorecard รายคน": "Invalid signature bypass code · or sign on each scorecard",
  "รองรับไฟล์ .xlsx และ .xls": "Supports .xlsx and .xls files",
  "ระดับบทลงโทษไม่ถูกต้อง": "Invalid penalty level",
  "ลบธีมที่กำลังใช้งานไม่ได้ ให้เปิดใช้ธีมอื่นก่อน": "The active theme cannot be deleted; activate another theme first",
  "ลบอีเวนต์ไม่สำเร็จ กรุณาลองใหม่หรือแจ้งผู้ดูแลระบบ": "Could not delete the event; try again or contact the administrator",
  "ล้าง Sudden death ของกลุ่มนี้แล้ว": "Sudden death cleared for this group",
  "สิทธิ์ไม่ถูกต้อง": "Invalid role",
  "อีเวนต์นี้มีผล Knockout แล้ว · ตรวจสาย Knockout อีกครั้งหลังคีย์ใหม่": "This event already has knockout results · recheck the bracket after re-entering",
  "อีเวนต์นี้อยู่ในสายอื่นแล้ว:": "This event is already in another chain:",
  "อีเวนต์นี้ไม่ได้เปิดใช้รอบ 2": "Round 2 is not enabled for this event",
  "อ่านไฟล์ไม่สำเร็จ:": "Could not read the file:",
  "เปิดการเรียงตามอันดับสดแล้ว": "Live rank sorting turned on",
  "ปิดการเรียงตามอันดับสดแล้ว": "Live rank sorting turned off",
  "เลือกอย่างน้อย 2 อีเวนต์เพื่อต่อคิวกัน": "Select at least 2 events to chain",
  "แก้ไข/ล้าง Sudden death ได้เฉพาะ superadmin": "Only a superadmin can edit/clear sudden death",
  "แสดงเส้นตัดแล้ว": "Cut lines shown",
  "ซ่อนเส้นตัดแล้ว": "Cut lines hidden",
  "ไม่พบผู้ใช้ (บัญชีสนามให้เปลี่ยนรหัสจากหน้าอีเวนต์)": "User not found (change court account passwords from the event page)",
  "ไม่พบหัวข้อประเภทที่มีรายชื่อในไฟล์": "No category heading with entries was found in the file",
  "ไม่มีสิทธิ์ดูประวัติคะแนนของอีเวนต์หรือสนามอื่น": "You cannot view score history of other events or courts",
  "บันทึกตั้งค่า Results Approved แล้ว": "Results Approved settings saved",
  "ล้างการคีย์": "Clear entries:",
  "(จางๆ หลังแบนเนอร์)": "(faded behind the banner)",
  "(สร้างอีเวนต์พร้อมรายชื่อ) หรือ": "(create events with entry lists) or",
  "(แนวนอน กว้าง)": "(landscape, wide)",
  "+ สร้างธีมใหม่จากภาพ": "+ New theme from image",
  "+ สร้างสายคิวใหม่": "+ New court chain",
  "+ สร้างหลายอีเวนต์": "+ Bulk create events",
  "+ เพิ่ม 1 แถว": "+ Add 1 row",
  "+ เพิ่มแถวตามที่เลือก": "+ Add selected rows",
  "(ยังไม่ครบ)": "(incomplete)",
  "1 · โยนภาพของงาน ให้ระบบสร้างสีธีม": "1 · Drop the event image to generate theme colors",
  "อัปเดตหน้าอัตโนมัติทุก 15 วินาที": "auto-refreshes every 15 seconds",
  "1) เลือกอีเวนต์ (เรียงตามลำดับที่จะตี) · 2) ดูผัง สนาม × ลำดับคิว · 3) กด \"ลงสนามตามผังนี้\" กรรมการสนามจะเห็นคิวรอบ 2 ตามผังทันที · ลำดับในรอบ 2: คะแนนรอบ 1 น้อยสุดตีก่อน เริ่มจากสนาม 1": "1) Select events (in playing order) · 2) Review the court × queue plan · 3) Press \"Apply this plan\" and court referees see the round 2 queue immediately · Round 2 order: lowest round 1 score shoots first, starting on court 1",
  "2 · สี": "2 · Colors",
  "3 · ภาพ": "3 · Images",
  "4 · ข้อความ": "4 · Text",
  "Court ID ของอีเวนต์นี้": "Court IDs for this event",
  "ID สำหรับล็อกอิน": "Login ID",
  "ผ่านตรงไม่ต้องตีรอบ 2:": "Direct qualifiers (no round 2):",
  "คีย์คู่": "Match entry",
  "กดเพื่อซ่อนเส้นตัด QUARTERFINALS / QUALIFIED FOR ROUND 2": "Click to hide the QUARTERFINALS / QUALIFIED FOR ROUND 2 cut lines",
  "กดเพื่อปิด: แถวจะอยู่ตามลำดับการตี ไม่สลับตามคะแนน": "Click to turn off: rows stay in shooting order instead of re-sorting by score",
  "กรรมการยกคะแนน": "Scoring referee",
  "กรอกตาราง": "Fill in table",
  "กรอกทีละสถานี": "Enter one station at a time",
  "กรอกหลายอีเวนต์ในตาราง": "Enter several events in a table",
  "กรองวันแข่งขัน": "Filter by competition date",
  "กลับ": "Back",
  "กลับ Overview": "Back to overview",
  "กลับ Results Approved": "Back to Results Approved",
  "กลับ Scorecard": "Back to scorecard",
  "กลับหน้ารายการ": "Back to list",
  "กลับหน้าอีเวนต์": "Back to event",
  "กลับไปหน้าธีม": "Back to themes",
  "การจัดลำดับ": "Ordering",
  "กำลังแข่งขัน": "In progress",
  "กำลังโหลด...": "Loading...",
  "กำลังใช้": "Active",
  "กำหนด / เปลี่ยนรหัส": "Set / change password",
  "ขึ้น": "Up",
  "ข้อความ APPROVED": "APPROVED text",
  "ข้อความใหญ่บนปก (เว้นว่าง = ใช้ชื่องานหน้าเครื่องหมาย ·)": "Large cover text (blank = event name before the · sign)",
  "คงลำดับการตี": "Keep shooting order",
  "คนต่อสนามไม่เกิน": "Max athletes per court",
  "คลิกที่ทีมเพื่อเลือก/ไม่เลือก แล้วกดบันทึก · ผู้ผ่านตรงไม่อยู่ในรายการนี้": "Click a team to select/deselect, then save · Direct qualifiers are not in this list",
  "คอลัมน์ NAME เดียว (ชื่อเต็มตามที่ลงทะเบียน)": "Single NAME column (full registered name)",
  "คะแนนเท่ากัน: ตี 7 เมตรทุกสถานีอีก 1 เที่ยว บันทึกพร้อมกันทุกคน · ถ้ายังเท่ากันระบบจะให้ตี Sudden death ต่อ": "Tie: one extra round at 7 m on every station, entered for everyone together · If still tied, the system continues with sudden death",
  "คัดลอก": "Copy",
  "คำนวณสีเข้ม/พื้นอ่อน/เส้นกรอบ/ตัวอักษร ให้เข้ากับสีหลัก และให้ตัวหนังสือสีขาวบนปุ่มยังอ่านง่าย": "Derives dark / soft background / border / text colors from the main color, keeping white button text readable",
  "คำนำหน้าชื่องาน": "Event name prefix",
  "คีย์คู่นี้": "Enter this match",
  "ค่า N": "Value N",
  "ค้นหาชื่องาน / รุ่น / ประเภท": "Search event / group / category",
  "ค้นหาอีเวนต์": "Search events",
  "จบการตีทั้งคู่": "Finish both",
  "จากรอบ 2 คัดเข้ารอบต่อ (คน)": "Qualifiers from round 2 (athletes)",
  "จำนวนคะแนน 3 รวมรอบ 1 และรอบ 2": "Number of 3s, rounds 1 and 2 combined",
  "จำนวนคะแนน 5 รวมรอบ 1 และรอบ 2": "Number of 5s, rounds 1 and 2 combined",
  "จำนวนผู้ผ่านเข้ารอบถัดไปทันทีจากรอบแรก": "Direct qualifiers from round 1",
  "จำนวนสนาม (ค่าคงที่)": "Number of courts (fixed)",
  "จำนวนสนาม / สนามแรก": "Courts / first court",
  "จำนวนสนามเริ่มต้นต่อแถว": "Default courts per row",
  "ชื่อ Technical Delegate": "Technical Delegate name",
  "ชื่อ:": "Name:",
  "ชื่อการแข่งขัน / Competition Title": "Competition title",
  "ชื่อธีม (เห็นเฉพาะผู้ดูแล)": "Theme name (admins only)",
  "ชื่อนักกีฬา (ใส่ระบบ)": "Athlete name (to import)",
  "ชื่อประธาน / President Name": "President name",
  "ชื่อประเภท": "Category name",
  "ชื่อย่องาน": "Event short name",
  "ชื่อรอบถัดไป": "Next round name",
  "ชื่อสนาม": "Court labels",
  "ชื่อสาย": "Chain name",
  "ชื่อใหญ่บนแบนเนอร์": "Banner title",
  "ซ่อน": "Hide",
  "ซ่อนเส้นแบ่ง": "Hide cut lines",
  "ดาวน์โหลด DOCX": "Download DOCX",
  "ดูตัวอย่างด้านขวาก่อน แล้วติ๊ก “เปิดใช้ทั้งระบบ” เมื่อพร้อม": "Check the preview on the right, then tick “Use site-wide” when ready",
  "ดูผลหลายอีเวนต์": "View several events",
  "ดูเต็มจอ →": "Full screen →",
  "ด้านล่าง ระบบจะพาไปหน้าตีเที่ยวพิเศษ 7 เมตรทันที": "below; the system opens the 7 m extra-round page immediately",
  "ตัวหนา = ผู้ชนะ": "Bold = winner",
  "ตัวอย่าง": "Preview",
  "ตัวอย่างคิวสนาม": "Court queue preview",
  "ตัวอักษร (1=A … 8=H)": "Letters (1=A … 8=H)",
  "ตัวอักษร (1=A, 2=B … 8=H)": "Letters (1=A, 2=B … 8=H)",
  "ตัวเลข": "Numbers",
  "ตัวเลข (1, 2, 3 …)": "Numbers (1, 2, 3 …)",
  "ตั้งค่า Results": "Results settings",
  "ตั้งค่า Results Approved": "Results Approved settings",
  "ตั้งค่าข้อความหัวเอกสารและรายชื่อกรรมการสำหรับ": "Document header text and officials for",
  "ตั้งค่าร่วม (ใช้กับทุกอีเวนต์ในชุดนี้)": "Shared settings (apply to every event in this set)",
  "ตั้งค่าหลายอีเวนต์": "Bulk event settings",
  "ตั้งค่าหลายอีเวนต์พร้อมกัน": "Configure several events at once",
  "ตั้งรหัส": "Set",
  "ตั้งรหัสผ่านใหม่": "Set new password",
  "ตามด้วยตาราง": "followed by a table",
  "ตำแหน่ง Technical Delegate": "Technical Delegate title",
  "ตำแหน่งประธาน / President Title": "President title",
  "ติ๊กเฉพาะหัวข้อที่ต้องการใช้เหมือนกัน ค่าอื่นของแต่ละอีเวนต์จะไม่ถูกเปลี่ยน": "Tick only the settings to share; other values of each event stay unchanged",
  "ตี 7 เมตร →": "7 m extra round →",
  "ตีครบ 5 สถานีแล้วคะแนนเท่ากัน · กด": "Tied after all 5 stations · press",
  "ตีแล้ว": "Scored",
  "ต่อคิวสนามข้ามรุ่น": "Court chains across groups",
  "ต่อจากผู้ผ่านตรงอีก N ลำดับ": "N more places after the direct qualifiers",
  "ต่อจากผู้ผ่านตรงอีก N ลำดับ (เท่ากันที่ลำดับสุดท้าย ตี Sudden death)": "N more places after the direct qualifiers (tie at the last place → sudden death)",
  "ต้องการสร้างหลายรายการพร้อมกัน?": "Need to create several events at once?",
  "ถอนตัว": "Withdrawn",
  "ถึงลำดับ (Class) ที่ N": "Up to place (Class) N",
  "ที่ · ชื่ออปท. · อำเภอ · จังหวัด": "No. · Organization · District · Province",
  "ทุกวันแข่งขัน": "All dates",
  "ท้ายภาพผลคะแนน · ชื่องาน": "Results image footer · event name",
  "ท้ายภาพผลคะแนน · สโลแกน": "Results image footer · slogan",
  "ธีม": "Theme",
  "ธีมของระบบ": "Site themes",
  "นักกีฬา/ทีม": "Athletes/teams",
  "นำเข้าจากไฟล์ใบสมัคร": "Import from entry file",
  "นำเข้าจากไฟล์ใบสมัคร (.xls / .xlsx)": "Import from entry file (.xls / .xlsx)",
  "นำเข้าชื่อ": "Import names",
  "นำเข้าชื่อนักกีฬา": "Import athlete names",
  "บรรทัดประเทศ/เจ้าภาพ": "Country / host line",
  "บรรทัดรอง": "Subtitle line",
  "บรรทัดเล็กด้านบน": "Small top line",
  "บังคับจับใหม่แม้เริ่มตีแล้ว (ล้างสาย Knockout)": "Force a new draw even after scoring started (clears the knockout bracket)",
  "บัญชีกรรมการสนาม (court01, court02 …): ตั้งรหัสที่อีเวนต์ใดก็ได้ในสาย กรรมการคนเดิมจะเห็นคิวของสนามตัวเองต่อกันทุกรุ่นในสาย": "Court referee accounts (court01, court02 …): set the password in any event of the chain; the same referee sees their court’s queue across every group in the chain",
  "บันทึกตั้งค่า": "Save settings",
  "บันทึกทีมที่เลือก": "Save selected teams",
  "บันทึกพร้อมกัน": "Save together",
  "บันทึกลายเซ็น": "Save signature",
  "บันทึกสายคิว": "Save chain",
  "บันทึกเป็น JPG": "Save as JPG",
  "บันทึกเป็น JPG ทั้งหมด": "Save all as JPG",
  "บันทึกแล้วระบบจัดสนามให้ทุกอีเวนต์ในสายทันที": "On save, courts are assigned for every event in the chain immediately",
  "บันทึกให้ 1 อีเวนต์": "Save for 1 event",
  "ประวัติการแก้คะแนน": "Score edit history",
  "ประวัติรอบคัดเลือก": "Qualification history",
  "ปิด (คงลำดับการตี)": "Off (keep shooting order)",
  "ปิด — คงลำดับการตี": "Off — keep shooting order",
  "ปิดรอบ 2": "Round 2 off",
  "ปิดเมื่อรายการไม่ต้องการให้แถวสลับไปมา · อันดับ (Class) ยังแสดงตามปกติ": "Turn off when rows should not move around · Class is still shown",
  "ปุ่มรอง": "Secondary button",
  "ผลคะแนนรอบนี้ได้รับการรับรองแล้ว": "This round’s scores have been approved",
  "ผลรวมหลายอีเวนต์": "Combined results",
  "ผังรอบ 2": "Round 2 plan",
  "ผังลงสนามรอบ 2": "Round 2 court plan",
  "ผู้ผ่านจากรอบ 2": "Round 2 qualifiers",
  "ผู้ผ่านตรง": "Direct qualifiers",
  "ผู้แก้": "Edited by",
  "ผู้แข่งขัน": "Competitors",
  "พร้อมคีย์": "Ready to score",
  "พร้อมบันทึก": "Ready",
  "พิมพ์คำว่า": "Type",
  "พิมพ์ผัง": "Print plan",
  "ภาพพื้นหลังแบนเนอร์": "Banner background image",
  "มากับระบบ": "Built-in",
  "มาตรฐาน (ทั่วไป)": "Standard (general)",
  "มาตรฐาน (ใบเดิมของระบบ)": "Standard (original sheet)",
  "มีรอบ 2": "Round 2 on",
  "มีอีเวนต์ที่เริ่มกรอกคะแนนแล้ว การบันทึกจะย้ายสนามของนักกีฬา (คะแนนไม่หาย)": "Some events already have scores; saving moves athletes between courts (scores are kept)",
  "มือถือ/แท็บเล็ตแนวตั้งจะแสดงเฉพาะสถานีที่เลือก เพื่อให้ช่องใหญ่ขึ้นและกดง่ายขึ้น": "Phones/tablets in portrait show only the selected station, for bigger, easier buttons",
  "ยังเสมอ": "Still tied",
  "ยังไม่ตั้งรหัส": "No password yet",
  "ยังไม่มีการแก้คะแนนหลังลงลายเซ็น": "No score edits after signing",
  "ยังไม่มีสายคิว ·": "No court chains yet ·",
  "ยังไม่ได้ลงลายเซ็นผู้แก้ไข": "Editor has not signed yet",
  "ยังไม่ได้เลือกอีเวนต์": "No events selected",
  "ยิงลูกกลางขาว": "Middle white ball shot",
  "ยิงลูกสีขาวเหนือลูกดำ": "White ball over black ball shot",
  "ยิงลูกเดี่ยว": "Single ball shot",
  "ยิงลูกเป้า": "Target ball shot",
  "ยิงลูกเหนือลูกเป้า": "Ball over target ball shot",
  "ยืนยันจัดสนามใหม่": "Confirm court reassignment",
  "รวมเที่ยวนี้": "This round total",
  "รหัสของ": "Password for",
  "รหัสผ่าน (อย่างน้อย 8 ตัว)": "Password (min. 8 characters)",
  "รหัสใหม่": "New password",
  "รอคู่แข่งขัน": "Waiting for opponent",
  "รอตี": "Waiting",
  "เลือกคะแนนแบบ 0 / 1 / 3 / 5": "Select a score of 0, 1, 3 or 5",
  "รอบ 1 และรอบ 2": "Rounds 1 and 2",
  "รอบ 16 คน": "Round of 16",
  "รอบการแข่งขัน": "Competition round",
  "รอบที่จะล้าง": "Round to clear",
  "รอแข่งขัน": "Waiting to compete",
  "ระบบคัดอัตโนมัติ": "Automatic selection",
  "ระบบจะแยกเป็น 1 อีเวนต์ต่อ 1 หัวข้อ และให้ตรวจสอบก่อนสร้างจริง": "Each heading becomes one event; you can review before they are created",
  "ระบบจะใส่ชื่อให้นักกีฬาตาม id\n    ไม่เปลี่ยนสังกัด ลำดับตี สนาม หรือคะแนน · แถวที่ชื่อว่างจะข้าม · อัปโหลดซ้ำได้": "Names are filled in by athlete id. Affiliation, order, court and scores are not changed · Rows with an empty name are skipped · You can upload again",
  "ระบบบันทึกคะแนนและจัดอันดับเปตอง Shooting": "Pétanque shooting scoring and ranking system",
  "ระบบเตรียม ID ตามจำนวนสนาม": "IDs are prepared automatically for",
  "สนามให้อัตโนมัติ · กำหนดรหัสเมื่อพร้อมใช้งาน": "courts · set passwords when ready",
  "ระบบเลือกสีหลักและสีเน้นให้แล้ว · กด “หลัก” หรือ “เน้น” ใต้สีเพื่อเปลี่ยน": "Main and accent colors were picked automatically · Press “main” or “accent” under a color to change",
  "ระบบใช้ตามรอบถัดไปอัตโนมัติ: รอบ 8 = 4, รอบ 16 = 8": "Set automatically from the next round: round of 8 = 4, round of 16 = 8",
  "ระยะ": "Distance",
  "รายการทั้งหมด": "All events",
  "รายการอีเวนต์": "Events",
  "รายชื่อ": "Entries",
  "รายชื่อกรรมการ Umpire": "Umpires",
  "รูปแบบชื่อในเอกสาร": "Name format in documents",
  "รูปแบบใบ:": "Sheet layout:",
  "รูปแบบใบบันทึกคะแนน": "Scorecard layout",
  "รูปแบบใบบันทึกคะแนน (พิมพ์)": "Scorecard layout (print)",
  "ลง": "Down",
  "ลงลายเซ็นผู้แก้ไข": "Editor signature",
  "ลงสนามตามผังนี้": "Apply this plan",
  "ลบแถว": "Delete row",
  "ลากภาพโปสเตอร์ / โลโก้งานมาวางที่นี่": "Drag the poster / event logo here",
  "ลายเซ็น": "Signature",
  "ลายเซ็นเฉพาะ": "Signatures for",
  "ล้างการคีย์ทั้งอีเวนต์": "Clear all entries in this event",
  "ล้างการคีย์ทั้งอีเวนต์ (เช่น ทดลองคีย์ก่อนแข่งจริง)": "Clear all entries in this event (e.g. test entries before the real competition)",
  "ล้างการคีย์รอบนี้ (กรณีคีย์ก่อนเวลาหรือคีย์ผิด)": "Clear this round’s entries (entered too early or by mistake)",
  "ล้างคะแนน ใบแดง ลายเซ็น และ Sudden death ทุกคนในรอบที่เลือก กลับเป็นสถานะรอตี · ล้างสาย Knockout ของอีเวนต์นี้ · รายชื่อ ลำดับยิง และสนามไม่เปลี่ยน · ช่องที่เคยคีย์ถูกบันทึกในประวัติการแก้": "Clears scores, red cards, signatures and sudden death for everyone in the selected round, back to waiting · Clears this event’s knockout bracket · Entries, order and courts stay · Entered cells are logged in the edit history",
  "ล้างคะแนนทุกช่อง ใบแดง ลายเซ็น/APPROVED และ Sudden death ของรอบนี้ · ช่องที่เคยคีย์ถูกบันทึกในประวัติการแก้พร้อมเหตุผล": "Clears every score, red card, signature/APPROVED and sudden death of this round · Entered cells are logged in the edit history with the reason",
  "ล้างทั้งหมด": "Clear all",
  "วันที่บนเอกสาร": "Date on documents",
  "วิธีคัดผู้มีสิทธิ์ตีรอบ 2": "Round 2 eligibility",
  "วิธีคัดรอบ 2": "Round 2 selection",
  "วิธีคัดเข้ารอบ 2": "Round 2 selection",
  "ศูนย์จับสลาก": "Draw center",
  "สถานที่": "Venue",
  "สถานที่ / Location": "Location",
  "สถานี": "Station",
  "สถิติประกอบการจัดอันดับ 5/3": "Ranking statistics 5/3",
  "สนาม เริ่มที่สนาม": "courts, starting at court",
  "เริ่มที่สนาม": "starting at court",
  "สนาม:": "Court:",
  "สนามแรกของอีเวนต์": "First court of the event",
  "สร้างธีม": "Create theme",
  "สร้างธีมใหม่": "New theme",
  "สร้างสายคิวใหม่": "New court chain",
  "สร้างสีรองจากสีหลักใหม่": "Regenerate secondary colors from the main color",
  "สร้างหลายอีเวนต์": "Bulk create events",
  "สร้างหลายอีเวนต์ / นำเข้าจากไฟล์": "Bulk create events / import from file",
  "สร้างหลายอีเวนต์พร้อมกัน": "Create multiple events",
  "สร้างอีเวนต์ทั้งหมด": "Create all events",
  "สร้างแถวอัตโนมัติ": "Generate rows",
  "สะสมเดิม:": "Previous total:",
  "สิทธิ์รอบ 2": "Round 2 eligible",
  "สีตัวอักษร": "Text color",
  "สีพื้นหน้า": "Page background",
  "สีพื้นอ่อน": "Soft background",
  "สีหลัก (ปุ่ม/แบนเนอร์)": "Main color (buttons/banner)",
  "สีหลักเข้ม (หัวข้อ)": "Dark main color (headings)",
  "สีเน้น": "Accent color",
  "สีเน้น (เส้นตัด)": "Accent color (cut lines)",
  "สีเส้นกรอบ": "Border color",
  "หน้านี้อัปเดตเองทุก 15 วินาที": "This page refreshes every 15 seconds",
  "หน้ารวม: เรียงแถวตามอันดับสด": "Overview: sort rows by live rank",
  "หน้ารวมทุกสนาม": "All courts overview",
  "หมายเหตุ (ไม่บังคับ)": "Note (optional)",
  "หรือคลิกเพื่อเลือกไฟล์ · PNG, JPG, WEBP · ภาพนี้จะถูกใช้เป็นพื้นหลังแบนเนอร์ด้วย": "or click to choose a file · PNG, JPG, WEBP · This image is also used as the banner background",
  "หลายรายการแล้วกดสร้างครั้งเดียว · เมื่อสร้างเสร็จจะไปหน้า": "several rows, then create them in one go · afterwards you go to the",
  "หัวข้อ": "Heading",
  "หัวคอลัมน์สังกัด": "Affiliation column header",
  "หัวภาพผลคะแนน · ชื่องาน": "Results image header · event name",
  "หัวภาพผลคะแนน · บรรทัดเล็ก": "Results image header · small line",
  "หัวภาพผลคะแนน · สถานที่/หน่วยงาน": "Results image header · venue/organization",
  "หากต้องแก้คะแนน ระบบจะเก็บค่าเดิม/ค่าใหม่ ผู้แก้ เวลา และต้องมีลายเซ็นผู้แก้ไข": "Score edits keep the old/new value, editor and time, and need the editor’s signature",
  "หากแก้จำนวนสนามในอีเวนต์ ระบบจะเพิ่มหรือลด Court ID ให้ตรงกับจำนวนสนามโดยอัตโนมัติ": "If you change the number of courts, court IDs are added or removed automatically",
  "อนุญาตตามระบบ": "System default",
  "อัปโหลดไฟล์ Excel ที่มีคอลัมน์": "Upload an Excel file with the columns",
  "อัปโหลดไฟล์ png, jpg, jpeg, webp หรือ gif ได้เอง ถ้าไม่อัปโหลดระบบจะใช้โลโก้ค่าเริ่มต้นตามเอกสารตัวอย่าง": "Upload png, jpg, jpeg, webp or gif files; otherwise the default logos are used",
  "อีเวนต์ที่เลือก ·": "Selected events ·",
  "อีเวนต์ในสาย": "Events in chain",
  "อ่านไฟล์และแสดงตัวอย่าง": "Read file and preview",
  "เกิน 10 นาที ตัดออก (DQ)": "Over 10 min — disqualified (DQ)",
  "เกิน 5 นาที −5": "Over 5 min −5",
  "เข้ารอบทันทีจากรอบแรก (คน)": "Direct qualifiers from round 1",
  "เฉพาะหัวข้อที่มีคำว่า": "Only headings containing",
  "เช่น 3 สนาม เริ่มสนาม 5 = ใช้สนาม 5-7 · ลำดับตีไม่เปลี่ยน": "e.g. 3 courts starting at court 5 = courts 5–7 · Shooting order is unchanged",
  "เช่น": "e.g.",
  "เดิม": "Old",
  "เปิด": "On",
  "เปิด — แถวสลับตามคะแนน": "On — rows re-sort by score",
  "เปิดรอบ 2": "Round 2 on",
  "เปิดใช้ทั้งระบบ": "Use site-wide",
  "เพื่อกันคะแนนทับกัน หน้านี้จึงดูได้อย่างเดียว ถ้าเครื่องเดิมเลิกใช้แล้วให้กดรับช่วงคีย์ต่อ": "To avoid overwriting scores this page is read-only. If the other device is no longer in use, take over.",
  "เพื่อยืนยัน": "to confirm",
  "เพื่อสั่งจับพร้อมกัน": "to draw them together",
  "เพื่อให้หลายรุ่นตีต่อกันบนสนามชุดเดียว": "so several groups play back-to-back on one set of courts",
  "เริ่มจากสนาม 1 · คะแนนรอบ 1 น้อยสุดตีก่อน (เท่ากันดูจำนวน 5 แล้วจำนวน 3 น้อยกว่าตีก่อน)": "Starts on court 1 · Lowest round 1 score shoots first (ties: fewer 5s, then fewer 3s first)",
  "เริ่มตีแล้ว": "Started",
  "เริ่มตีแล้ว · ล็อก": "Started · locked",
  "เรียงตาม TOTAL → จำนวน 5 → จำนวน 3 ใช้ดูเหตุผลการจัดลำดับ ไม่รวมคะแนนรอบพิเศษในตารางหลัก": "Sorted by TOTAL → number of 5s → number of 3s to explain the ranking; extra rounds are not included in the main table",
  "เรียงตามชื่อ (12→18, ชาย→หญิง)": "Sort by name (12→18, men→women)",
  "เรียงตามอันดับ: เปิด": "Sort by rank: on",
  "เรียงตามอันดับสด": "Sort by live rank",
  "เลือก": "Select",
  "เลือกทั้งหมดที่แสดง": "Select all shown",
  "เลือกทีมเข้ารอบ 2": "Select round 2 teams",
  "เลือกที่พร้อมจับ (ที่แสดง)": "Select drawable (shown)",
  "เลือกทุกอีเวนต์": "Select all events",
  "เลือกภาพของงาน": "Choose event image",
  "เลือกรายการใหม่": "Choose other events",
  "เลือกสถานีสำหรับหน้าจอแนวตั้ง": "Choose station for portrait screens",
  "เลือกสนาม": "Select court",
  "เลือกหลายอีเวนต์แล้วกดจับสลากครั้งเดียว ระบบสุ่มลำดับยิงและกระจายสนามใหม่ให้ทุกอีเวนต์พร้อมกัน · อีเวนต์ที่เริ่มกรอกคะแนนแล้วจะถูกล็อกไว้": "Select several events and draw once: shooting order and courts are randomized for every event together · Events that already have scores are locked",
  "เลือกอีเวนต์ที่ใช้สนามชุดเดียวกัน แล้วเรียงลำดับรุ่นที่จะตีก่อน–หลัง · ลำดับยิงในแต่ละรุ่นมาจากการจับสลากของรุ่นนั้น": "Select events sharing the same courts, then order the groups · Shooting order within each group comes from its own draw",
  "เลือกแล้ว": "Selected",
  "เลือกได้ 2 วิธี:": "Two ways:",
  "เลื่อนขึ้น": "Move up",
  "เลื่อนลง": "Move down",
  "เวลา": "Time",
  "เว้นว่าง = ทุกประเภท": "Blank = all categories",
  "เส้นตัด: เปิด": "Cut lines: on",
  "เส้นแบ่งเข้ารอบ": "Qualification cut lines",
  "เหตุผล": "Reason",
  "เหตุผล เช่น คีย์ก่อนนักกีฬาตีจริง": "Reason, e.g. entered before the athlete actually shot",
  "เหตุผลในการแก้ (ถ้ามี)": "Reason for the change (optional)",
  "แตะคะแนนของแต่ละระยะ · บันทึกอัตโนมัติทุกครั้งที่แตะ": "Tap the score for each distance · Saved automatically on every tap",
  "แถบโลโก้ผู้สนับสนุน": "Sponsor logo strip",
  "แยก FAMILY / GIVEN — คำสุดท้ายเป็นนามสกุล (แบบไทย)": "Split FAMILY / GIVEN — last word is the family name (Thai style)",
  "แยก FAMILY / GIVEN — คำแรกเป็นนามสกุล (แบบสากล)": "Split FAMILY / GIVEN — first word is the family name (international)",
  "และ": "and",
  "แสดง": "Show",
  "แสดงผัง": "Show plan",
  "แสดงหน้ากรรมการ": "Show officials page",
  "แสดงเส้นแบ่ง": "Show cut lines",
  "แสดงแบนเนอร์ด้านบนของเว็บ": "Show the banner at the top of the site",
  "โลโก้งาน": "Event logo",
  "โลโก้มุมซ้ายในหน้าผล": "Left corner logo on results pages",
  "โลโก้ล่างปก 1": "Cover bottom logo 1",
  "โลโก้ล่างปก 2": "Cover bottom logo 2",
  "โลโก้ล่างปก 3": "Cover bottom logo 3",
  "โลโก้หลักบนปก": "Main cover logo",
  "โลโก้หัวกระดาษ 1": "Header logo 1",
  "โลโก้หัวกระดาษ 2": "Header logo 2",
  "โลโก้หัวกระดาษ 3": "Header logo 3",
  "โลโก้หัวกระดาษ 4": "Header logo 4",
  "โลโก้เอกสาร": "Document logos",
  "ใช้คำนวณจำนวนสนาม (0 = ใช้ค่าคงที่)": "Used to compute the number of courts (0 = fixed value)",
  "ใช้ค่าเริ่มต้น": "Use default",
  "ใช้ได้ครั้งละ 1 ธีมทั้งเว็บ (แบนเนอร์ สีปุ่ม/เมนู และหัว–ท้ายภาพผลคะแนนที่บันทึกจากหน้ารวม) · สร้างธีมใหม่จากภาพของงานได้ ระบบดึงสีจากภาพให้อัตโนมัติ": "One theme applies site-wide at a time (banner, button/menu colors, results image header and footer) · Create a new theme from an event image; colors are extracted automatically",
  "ใบนี้กำลังถูกคีย์โดย": "This scorecard is being entered by",
  "ใบแดง": "Red card",
  "ใบแดงครบ 2 ครั้ง: ยุติการยิงรอบนี้และคงคะแนนที่ทำได้ไว้": "2 red cards: shooting stops for this round and the score so far is kept",
  "ใส่ชื่อจากไฟล์": "Fill names from file",
  "ใส่บรรทัดละ 1 คน รูปแบบ: ชื่อ | สังกัด เช่น MR. SOMCHAI JAIDEE | ABSC": "One person per line: name | federation, e.g. MR. SOMCHAI JAIDEE | ABSC",
  "ใหม่": "New",
  "ให้หลายอีเวนต์ตีต่อกันบนสนามชุดเดียว เมื่อรุ่นหนึ่งเหลือคนไม่เต็มสนาม รุ่นถัดไปเข้าตีสนามที่ว่างได้ทันที · อันดับและสิทธิ์เข้ารอบยังแยกต่ออีเวนต์ · ใช้กับรอบ 1 · อีเวนต์ที่ไม่ได้เลือกยังจัดสนามแยกเหมือนเดิม": "Lets several events play back-to-back on one set of courts: when a group no longer fills every court, the next group takes the free courts · Rankings and qualification stay per event · Round 1 only · Unselected events keep their own courts",
  "ไฟล์ค่าล้วน ไม่มีสูตร แก้ไขได้ทุกช่อง": "Values only, no formulas — every cell is editable",
  "ไฟล์ที่มีหัวข้อแต่ละประเภท เช่น": "A file with a heading per category, e.g.",
  "ไฟล์ใบสมัคร": "Entry file",
  "ไม่บังคับ": "Optional",
  "ไม่มาทำการแข่งขันตามเวลา": "Did not report on time",
  "ไม่มีบทลงโทษ": "No penalty",
  "ไม่มีรอบ 2": "No round 2",
  "ไม่ล้างคนที่ APPROVED แล้ว": "Keep athletes already APPROVED",
  "ไม่เลือก": "Not selected",
  "ไม่แสดง": "Hide",
  "← กลับ Overview": "← Back to overview",
  "← กลับ scorecard": "← Back to scorecard",
  "← กลับหน้าประกบคู่": "← Back to bracket",
  "← ก่อนหน้า": "← Previous",
  "↩ เลิกซ่อนที่เลือก": "↩ Unhide selected",
  "⚙ ตั้งค่า Results Approved / กรรมการ": "⚙ Results Approved / officials settings",
  "⚙️ ตั้งค่าร่วมกัน": "⚙️ Shared settings",
  "⚠️ ผลรอบนี้ลงลายเซ็นยืนยันแล้ว": "⚠️ This round has been signed and confirmed",
  "✓ ระบบลงสนามตามผังนี้อยู่แล้ว": "✓ Courts already follow this plan",
  "🎲 จับสลากพร้อมกัน": "🎲 Draw all selected",
  "🎲 ศูนย์จับสลาก": "🎲 Draw center",
  "👀 เปิดดูรวมกัน": "👀 View together",
  "👁 แสดงเฉพาะที่เลือก": "👁 Show only selected",
  "📄 Excel ไม่มีสูตร": "📄 Excel (no formulas)",
  "📄 QF / Final ไม่มีสูตร": "📄 QF / Final (no formulas)",
  "📊 ดูสถิติ 5/3": "📊 5/3 statistics",
  "🔗 ต่อคิวสนาม": "🔗 Court chains",
  "🔗 ต่อคิวสนามข้ามรุ่น": "🔗 Court chains across groups",
  "🖨 พิมพ์ Score Sheet รอบนี้": "🖨 Print this round’s score sheets",
  "🗂 ผังเข้ารอบ 2": "🗂 Round 2 plan",
  "🙈 ซ่อนที่เลือก": "🙈 Hide selected",
  "ซ่อนอยู่": "Hidden",
  "ซ่อนไว้": "Hidden:",
  "แสดงที่ซ่อนไว้": "Show hidden",
  "ซ่อนจากรายการ": "Hide from list"
 },
 "fr": {
  "ชื่อผู้ใช้ / สิทธิ์": "Utilisateur / rôle",
  "บัญชีที่กำลังใช้อยู่": "Compte actuel",
  "ลบบัญชีที่กำลังใช้อยู่ไม่ได้": "Impossible de supprimer le compte utilisé",
  "ต้องมี superadmin อย่างน้อย 1 บัญชี": "Au moins 1 compte superadmin est requis",
  "กรุณาใส่ชื่อผู้ใช้": "Veuillez saisir un nom d’utilisateur",
  "เปลี่ยนสิทธิ์ของบัญชีที่กำลังใช้อยู่ไม่ได้": "Impossible de changer le rôle du compte utilisé",
  "ไม่พบผู้ใช้ (บัญชีสนามให้จัดการจากหน้าอีเวนต์)": "Utilisateur introuvable (comptes terrain depuis la page de l’événement)",
  "รูปแบบการคีย์": "Mode de saisie",
  "📋 ตาราง (แบบเดิม)": "📋 Tableau (classique)",
  "👆 ปุ่มแตะ (แบบรอบน็อค)": "👆 Boutons tactiles (style KO)",
  "รวมรอบนี้": "Total du tour",
  "สถานีสุดท้าย": "Dernier atelier",
  "สถานีถัดไป →": "Atelier suivant →",
  "กรุณาเลือกนักกีฬาที่ต้องตี Sudden death อย่างน้อย 2 คน": "Sélectionnez au moins 2 athlètes pour la mort subite",
  "กรุณาเลือกนักกีฬาอย่างน้อย 1 คน": "Sélectionnez au moins 1 athlète",
  "คุณไม่มีสิทธิ์ทำรายการนี้": "Vous n’avez pas le droit d’effectuer cette action",
  "จบการตีเรียบร้อย": "Tir terminé",
  "ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง": "Nom d’utilisateur ou mot de passe incorrect",
  "ต้องลงชื่ออย่างใดอย่างหนึ่ง (พิมพ์ชื่อหรือเขียน) ให้ครบทั้ง 3 ฝ่าย หรือใช้สิทธิ์ข้าม": "Les 3 parties doivent signer (nom tapé ou signature), ou utiliser le contournement",
  "นักกีฬาคนนี้ไม่มีสิทธิ์ตีรอบ 2": "Cet athlète n’est pas qualifié pour le tour 2",
  "บันทึกผลเที่ยวพิเศษแล้ว ถ้ายังเท่ากันให้บันทึกเที่ยวพิเศษเพิ่มอีกครั้ง": "Tour supplémentaire enregistré ; en cas d’égalité, saisissez un nouveau tour",
  "บันทึกผู้ชนะแล้ว": "Vainqueur enregistré",
  "ผู้ชนะไม่ถูกต้อง": "Vainqueur invalide",
  "ยังไม่ได้ติดตั้ง python-docx ให้รัน: python -m pip install -r requirements.txt": "python-docx n’est pas installé ; lancez : python -m pip install -r requirements.txt",
  "รองรับเฉพาะไฟล์ .xlsx": "Fichiers .xlsx uniquement",
  "รอบแข่งขันไม่ถูกต้อง": "Tour invalide",
  "ลบอีเวนต์แล้ว": "Événement supprimé",
  "สร้างผู้ใช้สำเร็จ": "Utilisateur créé",
  "สร้างผู้ใช้ไม่สำเร็จ กรุณาตรวจสอบข้อมูล": "Échec de la création de l’utilisateur, vérifiez les données",
  "สุ่มลำดับใหม่แล้ว": "Nouvel ordre tiré au sort",
  "ออกจากระบบแล้ว": "Déconnecté",
  "เข้าสู่ระบบสำเร็จ": "Connexion réussie",
  "เพิ่มนักกีฬาสำเร็จ": "Athlète ajouté",
  "แก้ไขอีเวนต์สำเร็จ": "Événement modifié",
  "ปิดการเรียงสดแล้ว · แถวอยู่ตามลำดับการตี": "Tri en direct désactivé · ordre de tir",
  "กรุณาระบุเหตุผลในการล้างการคีย์": "Indiquez un motif pour effacer les saisies",
  "กรุณาเพิ่มอย่างน้อย 1 อีเวนต์": "Ajoutez au moins 1 événement",
  "กรุณาเลือกอีเวนต์": "Veuillez choisir des événements",
  "กรุณาเลือกอีเวนต์ที่จะจับสลาก": "Choisissez les événements à tirer",
  "กรุณาเลือกอีเวนต์อย่างน้อย 1 รายการ": "Choisissez au moins 1 événement",
  "กรุณาเลือกไฟล์": "Veuillez choisir un fichier",
  "ข้ามอีเวนต์ที่เริ่มกรอกคะแนนแล้ว:": "Événements déjà commencés ignorés :",
  "คะแนนเท่ากัน · ตีเที่ยวพิเศษ 7 เมตรทุกสถานี": "Égalité · tour supplémentaire à 7 m sur chaque atelier",
  "คู่นี้ยังไม่ครบ 2 คน": "Ce match n’a pas encore 2 athlètes",
  "จบการตีคู่นี้แล้ว": "Match terminé",
  "ธีมที่มากับระบบลบไม่ได้ (แก้ไขหรือคัดลอกได้)": "Les thèmes intégrés ne peuvent pas être supprimés (modifiables ou copiables)",
  "บัญชีนี้บันทึกบทลงโทษได้เฉพาะนักกีฬาในสนามของตัวเอง": "Ce compte ne peut pénaliser que les athlètes de son terrain",
  "บัญชีนี้ยังใช้รหัสผ่านตั้งต้นที่เปิดเผยในโค้ด จึงถูกระงับ ให้ superadmin ตั้งรหัสใหม่ที่หน้าผู้ใช้": "Ce compte utilise encore un mot de passe par défaut et est bloqué ; un superadmin doit en définir un nouveau",
  "บัญชีสนามกรอกเที่ยวพิเศษได้เฉพาะคู่ Knockout ในสนามตัวเอง": "Les comptes terrain ne saisissent les tours supplémentaires que pour les matchs KO de leur terrain",
  "บัญชีสนามนี้คีย์ได้เฉพาะคู่ในสนามตัวเอง": "Ce compte terrain ne saisit que les matchs de son terrain",
  "บัญชีสนามให้สร้างจากหน้านักกีฬาของอีเวนต์": "Créez les comptes terrain depuis la page des athlètes",
  "บันทึก Sudden death แล้ว": "Mort subite enregistrée",
  "พิมพ์คำว่า ล้าง ในช่องยืนยันก่อน": "Tapez ล้าง dans la case de confirmation",
  "มี court ID นี้มากกว่า 1 อีเวนต์ที่ใช้รหัสเดียวกัน กรุณาเปลี่ยนรหัสของอีเวนต์ให้ไม่ซ้ำ": "Cet identifiant terrain a le même mot de passe dans plusieurs événements ; changez-les",
  "มีอีเวนต์ที่เริ่มกรอกคะแนนแล้ว การจัดสนามใหม่จะย้ายสนามของนักกีฬา · ติ๊กยืนยันก่อนบันทึก": "Certains événements ont déjà des scores ; la réattribution déplace les athlètes · cochez la confirmation",
  "ยังยิงไม่ครบทุกสถานี จึงยัง Approve ไม่ได้": "Tous les ateliers ne sont pas terminés : validation impossible",
  "ยังไม่ได้ติ๊กหัวข้อการตั้งค่าที่ต้องการใช้ร่วมกัน": "Aucun paramètre commun coché",
  "ยังไม่ได้เลือกประเภทที่จะสร้าง": "Aucune catégorie choisie",
  "รหัสข้ามลายเซ็นไม่ถูกต้อง · หรือไปลงลายเซ็นที่ Scorecard รายคน": "Code de contournement invalide · ou signez chaque feuille",
  "รองรับไฟล์ .xlsx และ .xls": "Fichiers .xlsx et .xls acceptés",
  "ระดับบทลงโทษไม่ถูกต้อง": "Niveau de pénalité invalide",
  "ลบธีมที่กำลังใช้งานไม่ได้ ให้เปิดใช้ธีมอื่นก่อน": "Impossible de supprimer le thème actif ; activez-en un autre d’abord",
  "ลบอีเวนต์ไม่สำเร็จ กรุณาลองใหม่หรือแจ้งผู้ดูแลระบบ": "Échec de la suppression ; réessayez ou contactez l’administrateur",
  "ล้าง Sudden death ของกลุ่มนี้แล้ว": "Mort subite effacée pour ce groupe",
  "สิทธิ์ไม่ถูกต้อง": "Rôle invalide",
  "อีเวนต์นี้มีผล Knockout แล้ว · ตรวจสาย Knockout อีกครั้งหลังคีย์ใหม่": "Cet événement a déjà des résultats KO · revérifiez le tableau après la nouvelle saisie",
  "อีเวนต์นี้อยู่ในสายอื่นแล้ว:": "Cet événement est déjà dans une autre chaîne :",
  "อีเวนต์นี้ไม่ได้เปิดใช้รอบ 2": "Le tour 2 n’est pas activé pour cet événement",
  "อ่านไฟล์ไม่สำเร็จ:": "Lecture du fichier impossible :",
  "เปิดการเรียงตามอันดับสดแล้ว": "Tri par classement en direct activé",
  "ปิดการเรียงตามอันดับสดแล้ว": "Tri par classement en direct désactivé",
  "เลือกอย่างน้อย 2 อีเวนต์เพื่อต่อคิวกัน": "Choisissez au moins 2 événements à enchaîner",
  "แก้ไข/ล้าง Sudden death ได้เฉพาะ superadmin": "Seul un superadmin peut modifier/effacer la mort subite",
  "แสดงเส้นตัดแล้ว": "Lignes affichées",
  "ซ่อนเส้นตัดแล้ว": "Lignes masquées",
  "ไม่พบผู้ใช้ (บัญชีสนามให้เปลี่ยนรหัสจากหน้าอีเวนต์)": "Utilisateur introuvable (mots de passe terrain depuis la page de l’événement)",
  "ไม่พบหัวข้อประเภทที่มีรายชื่อในไฟล์": "Aucun titre de catégorie avec des inscrits dans le fichier",
  "ไม่มีสิทธิ์ดูประวัติคะแนนของอีเวนต์หรือสนามอื่น": "Vous ne pouvez pas voir l’historique d’autres événements ou terrains",
  "บันทึกตั้งค่า Results Approved แล้ว": "Paramètres des résultats validés enregistrés",
  "ล้างการคีย์": "Effacer les saisies :",
  "(จางๆ หลังแบนเนอร์)": "(estompé derrière la bannière)",
  "(สร้างอีเวนต์พร้อมรายชื่อ) หรือ": "(créer les événements avec les listes) ou",
  "(แนวนอน กว้าง)": "(paysage, large)",
  "+ สร้างธีมใหม่จากภาพ": "+ Nouveau thème depuis une image",
  "+ สร้างสายคิวใหม่": "+ Nouvelle chaîne de terrains",
  "+ สร้างหลายอีเวนต์": "+ Créer plusieurs événements",
  "+ เพิ่ม 1 แถว": "+ Ajouter 1 ligne",
  "+ เพิ่มแถวตามที่เลือก": "+ Ajouter les lignes choisies",
  "(ยังไม่ครบ)": "(incomplet)",
  "1 · โยนภาพของงาน ให้ระบบสร้างสีธีม": "1 · Déposez l’image de l’événement pour générer les couleurs",
  "อัปเดตหน้าอัตโนมัติทุก 15 วินาที": "actualisation automatique toutes les 15 secondes",
  "1) เลือกอีเวนต์ (เรียงตามลำดับที่จะตี) · 2) ดูผัง สนาม × ลำดับคิว · 3) กด \"ลงสนามตามผังนี้\" กรรมการสนามจะเห็นคิวรอบ 2 ตามผังทันที · ลำดับในรอบ 2: คะแนนรอบ 1 น้อยสุดตีก่อน เริ่มจากสนาม 1": "1) Choisissez les événements (dans l’ordre de jeu) · 2) Vérifiez le plan terrain × ordre · 3) Cliquez « Appliquer ce plan » : les arbitres voient aussitôt la file du tour 2 · Ordre du tour 2 : le plus petit score du tour 1 tire en premier, en commençant par le terrain 1",
  "2 · สี": "2 · Couleurs",
  "3 · ภาพ": "3 · Images",
  "4 · ข้อความ": "4 · Textes",
  "Court ID ของอีเวนต์นี้": "Identifiants terrain de cet événement",
  "ID สำหรับล็อกอิน": "Identifiant de connexion",
  "ผ่านตรงไม่ต้องตีรอบ 2:": "Qualifiés directs (sans tour 2) :",
  "คีย์คู่": "Saisie du match",
  "กดเพื่อซ่อนเส้นตัด QUARTERFINALS / QUALIFIED FOR ROUND 2": "Cliquez pour masquer les lignes QUARTERFINALS / QUALIFIED FOR ROUND 2",
  "กดเพื่อปิด: แถวจะอยู่ตามลำดับการตี ไม่สลับตามคะแนน": "Cliquez pour désactiver : les lignes restent dans l’ordre de tir",
  "กรรมการยกคะแนน": "Arbitre marqueur",
  "กรอกตาราง": "Remplir le tableau",
  "กรอกทีละสถานี": "Saisie atelier par atelier",
  "กรอกหลายอีเวนต์ในตาราง": "Saisir plusieurs événements dans un tableau",
  "กรองวันแข่งขัน": "Filtrer par date",
  "กลับ": "Retour",
  "กลับ Overview": "Retour à l’aperçu",
  "กลับ Results Approved": "Retour aux résultats validés",
  "กลับ Scorecard": "Retour à la feuille",
  "กลับหน้ารายการ": "Retour à la liste",
  "กลับหน้าอีเวนต์": "Retour à l’événement",
  "กลับไปหน้าธีม": "Retour aux thèmes",
  "การจัดลำดับ": "Classement",
  "กำลังแข่งขัน": "En cours",
  "กำลังโหลด...": "Chargement...",
  "กำลังใช้": "Actif",
  "กำหนด / เปลี่ยนรหัส": "Définir / changer le mot de passe",
  "ขึ้น": "Monter",
  "ข้อความ APPROVED": "Texte APPROVED",
  "ข้อความใหญ่บนปก (เว้นว่าง = ใช้ชื่องานหน้าเครื่องหมาย ·)": "Grand texte de couverture (vide = nom avant le signe ·)",
  "คงลำดับการตี": "Garder l’ordre de tir",
  "คนต่อสนามไม่เกิน": "Max. athlètes par terrain",
  "คลิกที่ทีมเพื่อเลือก/ไม่เลือก แล้วกดบันทึก · ผู้ผ่านตรงไม่อยู่ในรายการนี้": "Cliquez sur une équipe pour la (dé)sélectionner, puis enregistrez · Les qualifiés directs ne figurent pas ici",
  "คอลัมน์ NAME เดียว (ชื่อเต็มตามที่ลงทะเบียน)": "Colonne NAME unique (nom complet inscrit)",
  "คะแนนเท่ากัน: ตี 7 เมตรทุกสถานีอีก 1 เที่ยว บันทึกพร้อมกันทุกคน · ถ้ายังเท่ากันระบบจะให้ตี Sudden death ต่อ": "Égalité : un tour supplémentaire à 7 m sur chaque atelier, saisi pour tous · En cas de nouvelle égalité, mort subite",
  "คัดลอก": "Copier",
  "คำนวณสีเข้ม/พื้นอ่อน/เส้นกรอบ/ตัวอักษร ให้เข้ากับสีหลัก และให้ตัวหนังสือสีขาวบนปุ่มยังอ่านง่าย": "Calcule les teintes foncée / fond clair / bordure / texte à partir de la couleur principale, en gardant le texte blanc lisible",
  "คำนำหน้าชื่องาน": "Préfixe du nom",
  "คีย์คู่นี้": "Saisir ce match",
  "ค่า N": "Valeur N",
  "ค้นหาชื่องาน / รุ่น / ประเภท": "Rechercher événement / groupe / catégorie",
  "ค้นหาอีเวนต์": "Rechercher un événement",
  "จบการตีทั้งคู่": "Terminer les deux",
  "จากรอบ 2 คัดเข้ารอบต่อ (คน)": "Qualifiés du tour 2 (athlètes)",
  "จำนวนคะแนน 3 รวมรอบ 1 และรอบ 2": "Nombre de 3, tours 1 et 2 cumulés",
  "จำนวนคะแนน 5 รวมรอบ 1 และรอบ 2": "Nombre de 5, tours 1 et 2 cumulés",
  "จำนวนผู้ผ่านเข้ารอบถัดไปทันทีจากรอบแรก": "Qualifiés directs du tour 1",
  "จำนวนสนาม (ค่าคงที่)": "Nombre de terrains (fixe)",
  "จำนวนสนาม / สนามแรก": "Terrains / premier terrain",
  "จำนวนสนามเริ่มต้นต่อแถว": "Terrains par défaut par ligne",
  "ชื่อ Technical Delegate": "Nom du délégué technique",
  "ชื่อ:": "Nom :",
  "ชื่อการแข่งขัน / Competition Title": "Titre de la compétition",
  "ชื่อธีม (เห็นเฉพาะผู้ดูแล)": "Nom du thème (admins uniquement)",
  "ชื่อนักกีฬา (ใส่ระบบ)": "Nom de l’athlète (à importer)",
  "ชื่อประธาน / President Name": "Nom du président",
  "ชื่อประเภท": "Nom de la catégorie",
  "ชื่อย่องาน": "Nom court de l’événement",
  "ชื่อรอบถัดไป": "Nom du tour suivant",
  "ชื่อสนาม": "Noms des terrains",
  "ชื่อสาย": "Nom de la chaîne",
  "ชื่อใหญ่บนแบนเนอร์": "Titre de la bannière",
  "ซ่อน": "Masquer",
  "ซ่อนเส้นแบ่ง": "Masquer les lignes",
  "ดาวน์โหลด DOCX": "Télécharger DOCX",
  "ดูตัวอย่างด้านขวาก่อน แล้วติ๊ก “เปิดใช้ทั้งระบบ” เมื่อพร้อม": "Vérifiez l’aperçu à droite puis cochez « Utiliser partout » quand c’est prêt",
  "ดูผลหลายอีเวนต์": "Voir plusieurs événements",
  "ดูเต็มจอ →": "Plein écran →",
  "ด้านล่าง ระบบจะพาไปหน้าตีเที่ยวพิเศษ 7 เมตรทันที": "ci-dessous ; le système ouvre aussitôt la page du tour supplémentaire à 7 m",
  "ตัวหนา = ผู้ชนะ": "Gras = vainqueur",
  "ตัวอย่าง": "Aperçu",
  "ตัวอย่างคิวสนาม": "Aperçu de la file du terrain",
  "ตัวอักษร (1=A … 8=H)": "Lettres (1=A … 8=H)",
  "ตัวอักษร (1=A, 2=B … 8=H)": "Lettres (1=A, 2=B … 8=H)",
  "ตัวเลข": "Chiffres",
  "ตัวเลข (1, 2, 3 …)": "Chiffres (1, 2, 3 …)",
  "ตั้งค่า Results": "Paramètres des résultats",
  "ตั้งค่า Results Approved": "Paramètres des résultats validés",
  "ตั้งค่าข้อความหัวเอกสารและรายชื่อกรรมการสำหรับ": "Textes d’en-tête et officiels pour",
  "ตั้งค่าร่วม (ใช้กับทุกอีเวนต์ในชุดนี้)": "Paramètres communs (pour tous les événements de cet ensemble)",
  "ตั้งค่าหลายอีเวนต์": "Paramètres de plusieurs événements",
  "ตั้งค่าหลายอีเวนต์พร้อมกัน": "Configurer plusieurs événements à la fois",
  "ตั้งรหัส": "Définir",
  "ตั้งรหัสผ่านใหม่": "Nouveau mot de passe",
  "ตามด้วยตาราง": "suivi d’un tableau",
  "ตำแหน่ง Technical Delegate": "Fonction du délégué technique",
  "ตำแหน่งประธาน / President Title": "Fonction du président",
  "ติ๊กเฉพาะหัวข้อที่ต้องการใช้เหมือนกัน ค่าอื่นของแต่ละอีเวนต์จะไม่ถูกเปลี่ยน": "Cochez seulement les paramètres à partager ; les autres valeurs restent inchangées",
  "ตี 7 เมตร →": "Tour à 7 m →",
  "ตีครบ 5 สถานีแล้วคะแนนเท่ากัน · กด": "Égalité après les 5 ateliers · cliquez",
  "ตีแล้ว": "Joué",
  "ต่อคิวสนามข้ามรุ่น": "Chaînes de terrains entre groupes",
  "ต่อจากผู้ผ่านตรงอีก N ลำดับ": "N places de plus après les qualifiés directs",
  "ต่อจากผู้ผ่านตรงอีก N ลำดับ (เท่ากันที่ลำดับสุดท้าย ตี Sudden death)": "N places de plus après les qualifiés directs (égalité à la dernière place → mort subite)",
  "ต้องการสร้างหลายรายการพร้อมกัน?": "Besoin de créer plusieurs événements à la fois ?",
  "ถอนตัว": "Forfait",
  "ถึงลำดับ (Class) ที่ N": "Jusqu’à la place (Class) N",
  "ที่ · ชื่ออปท. · อำเภอ · จังหวัด": "N° · Organisation · District · Province",
  "ทุกวันแข่งขัน": "Toutes les dates",
  "ท้ายภาพผลคะแนน · ชื่องาน": "Pied de l’image · nom de l’événement",
  "ท้ายภาพผลคะแนน · สโลแกน": "Pied de l’image · slogan",
  "ธีม": "Thème",
  "ธีมของระบบ": "Thèmes du site",
  "นักกีฬา/ทีม": "Athlètes/équipes",
  "นำเข้าจากไฟล์ใบสมัคร": "Importer depuis le fichier d’inscription",
  "นำเข้าจากไฟล์ใบสมัคร (.xls / .xlsx)": "Importer depuis le fichier d’inscription (.xls / .xlsx)",
  "นำเข้าชื่อ": "Importer les noms",
  "นำเข้าชื่อนักกีฬา": "Importer les noms des athlètes",
  "บรรทัดประเทศ/เจ้าภาพ": "Ligne pays / hôte",
  "บรรทัดรอง": "Sous-titre",
  "บรรทัดเล็กด้านบน": "Petite ligne du haut",
  "บังคับจับใหม่แม้เริ่มตีแล้ว (ล้างสาย Knockout)": "Forcer un nouveau tirage même après le début (efface le tableau KO)",
  "บัญชีกรรมการสนาม (court01, court02 …): ตั้งรหัสที่อีเวนต์ใดก็ได้ในสาย กรรมการคนเดิมจะเห็นคิวของสนามตัวเองต่อกันทุกรุ่นในสาย": "Comptes arbitres (court01, court02 …) : définissez le mot de passe dans n’importe quel événement de la chaîne ; le même arbitre voit la file de son terrain pour tous les groupes",
  "บันทึกตั้งค่า": "Enregistrer les paramètres",
  "บันทึกทีมที่เลือก": "Enregistrer les équipes choisies",
  "บันทึกพร้อมกัน": "Enregistrer ensemble",
  "บันทึกลายเซ็น": "Enregistrer la signature",
  "บันทึกสายคิว": "Enregistrer la chaîne",
  "บันทึกเป็น JPG": "Enregistrer en JPG",
  "บันทึกเป็น JPG ทั้งหมด": "Tout enregistrer en JPG",
  "บันทึกแล้วระบบจัดสนามให้ทุกอีเวนต์ในสายทันที": "À l’enregistrement, les terrains sont attribués à tous les événements de la chaîne",
  "บันทึกให้ 1 อีเวนต์": "Enregistrer pour 1 événement",
  "ประวัติการแก้คะแนน": "Historique des modifications",
  "ประวัติรอบคัดเลือก": "Historique des qualifications",
  "ปิด (คงลำดับการตี)": "Désactivé (ordre de tir)",
  "ปิด — คงลำดับการตี": "Désactivé — ordre de tir",
  "ปิดรอบ 2": "Tour 2 désactivé",
  "ปิดเมื่อรายการไม่ต้องการให้แถวสลับไปมา · อันดับ (Class) ยังแสดงตามปกติ": "Désactivez si les lignes ne doivent pas bouger · le classement reste affiché",
  "ปุ่มรอง": "Bouton secondaire",
  "ผลคะแนนรอบนี้ได้รับการรับรองแล้ว": "Les scores de ce tour sont validés",
  "ผลรวมหลายอีเวนต์": "Résultats combinés",
  "ผังรอบ 2": "Plan du tour 2",
  "ผังลงสนามรอบ 2": "Plan des terrains du tour 2",
  "ผู้ผ่านจากรอบ 2": "Qualifiés du tour 2",
  "ผู้ผ่านตรง": "Qualifiés directs",
  "ผู้แก้": "Modifié par",
  "ผู้แข่งขัน": "Compétiteurs",
  "พร้อมคีย์": "Prêt à saisir",
  "พร้อมบันทึก": "Prêt",
  "พิมพ์คำว่า": "Tapez",
  "พิมพ์ผัง": "Imprimer le plan",
  "ภาพพื้นหลังแบนเนอร์": "Image de fond de la bannière",
  "มากับระบบ": "Intégré",
  "มาตรฐาน (ทั่วไป)": "Standard (général)",
  "มาตรฐาน (ใบเดิมของระบบ)": "Standard (feuille d’origine)",
  "มีรอบ 2": "Tour 2 activé",
  "มีอีเวนต์ที่เริ่มกรอกคะแนนแล้ว การบันทึกจะย้ายสนามของนักกีฬา (คะแนนไม่หาย)": "Certains événements ont déjà des scores ; l’enregistrement déplace les athlètes (scores conservés)",
  "มือถือ/แท็บเล็ตแนวตั้งจะแสดงเฉพาะสถานีที่เลือก เพื่อให้ช่องใหญ่ขึ้นและกดง่ายขึ้น": "En portrait, téléphone/tablette n’affichent que l’atelier choisi pour des cases plus grandes",
  "ยังเสมอ": "Toujours à égalité",
  "ยังไม่ตั้งรหัส": "Pas encore de mot de passe",
  "ยังไม่มีการแก้คะแนนหลังลงลายเซ็น": "Aucune modification après signature",
  "ยังไม่มีสายคิว ·": "Aucune chaîne pour l’instant ·",
  "ยังไม่ได้ลงลายเซ็นผู้แก้ไข": "L’auteur de la modification n’a pas signé",
  "ยังไม่ได้เลือกอีเวนต์": "Aucun événement choisi",
  "ยิงลูกกลางขาว": "Tir boule blanche du milieu",
  "ยิงลูกสีขาวเหนือลูกดำ": "Tir boule blanche par-dessus la noire",
  "ยิงลูกเดี่ยว": "Tir boule seule",
  "ยิงลูกเป้า": "Tir au but",
  "ยิงลูกเหนือลูกเป้า": "Tir boule derrière le but",
  "ยืนยันจัดสนามใหม่": "Confirmer la réattribution des terrains",
  "รวมเที่ยวนี้": "Total de ce tour",
  "รหัสของ": "Mot de passe de",
  "รหัสผ่าน (อย่างน้อย 8 ตัว)": "Mot de passe (8 caractères min.)",
  "รหัสใหม่": "Nouveau mot de passe",
  "รอคู่แข่งขัน": "En attente de l’adversaire",
  "รอตี": "En attente",
  "เลือกคะแนนแบบ 0 / 1 / 3 / 5": "Choisissez un score 0 / 1 / 3 / 5",
  "รอบ 1 และรอบ 2": "Tours 1 et 2",
  "รอบ 16 คน": "Huitièmes de finale",
  "รอบการแข่งขัน": "Tour de compétition",
  "รอบที่จะล้าง": "Tour à effacer",
  "รอแข่งขัน": "En attente",
  "ระบบคัดอัตโนมัติ": "Sélection automatique",
  "ระบบจะแยกเป็น 1 อีเวนต์ต่อ 1 หัวข้อ และให้ตรวจสอบก่อนสร้างจริง": "Chaque titre devient un événement ; vous vérifiez avant la création",
  "ระบบจะใส่ชื่อให้นักกีฬาตาม id\n    ไม่เปลี่ยนสังกัด ลำดับตี สนาม หรือคะแนน · แถวที่ชื่อว่างจะข้าม · อัปโหลดซ้ำได้": "Les noms sont remplis selon l’id. Affiliation, ordre, terrain et scores inchangés · Lignes sans nom ignorées · Réimport possible",
  "ระบบบันทึกคะแนนและจัดอันดับเปตอง Shooting": "Système de saisie et classement du tir de précision",
  "ระบบเตรียม ID ตามจำนวนสนาม": "Identifiants préparés automatiquement pour",
  "สนามให้อัตโนมัติ · กำหนดรหัสเมื่อพร้อมใช้งาน": "terrains · définissez les mots de passe quand prêt",
  "ระบบเลือกสีหลักและสีเน้นให้แล้ว · กด “หลัก” หรือ “เน้น” ใต้สีเพื่อเปลี่ยน": "Couleurs principale et d’accent choisies · Cliquez « principale » ou « accent » sous une couleur pour changer",
  "ระบบใช้ตามรอบถัดไปอัตโนมัติ: รอบ 8 = 4, รอบ 16 = 8": "Défini selon le tour suivant : quarts = 4, huitièmes = 8",
  "ระยะ": "Distance",
  "รายการทั้งหมด": "Tous les événements",
  "รายการอีเวนต์": "Événements",
  "รายชื่อ": "Liste",
  "รายชื่อกรรมการ Umpire": "Arbitres",
  "รูปแบบชื่อในเอกสาร": "Format des noms",
  "รูปแบบใบ:": "Modèle de feuille :",
  "รูปแบบใบบันทึกคะแนน": "Modèle de feuille de score",
  "รูปแบบใบบันทึกคะแนน (พิมพ์)": "Modèle de feuille (impression)",
  "ลง": "Descendre",
  "ลงลายเซ็นผู้แก้ไข": "Signature de l’auteur",
  "ลงสนามตามผังนี้": "Appliquer ce plan",
  "ลบแถว": "Supprimer la ligne",
  "ลากภาพโปสเตอร์ / โลโก้งานมาวางที่นี่": "Glissez l’affiche / le logo ici",
  "ลายเซ็น": "Signature",
  "ลายเซ็นเฉพาะ": "Signatures du",
  "ล้างการคีย์ทั้งอีเวนต์": "Effacer toutes les saisies de l’événement",
  "ล้างการคีย์ทั้งอีเวนต์ (เช่น ทดลองคีย์ก่อนแข่งจริง)": "Effacer toutes les saisies (ex. tests avant la compétition)",
  "ล้างการคีย์รอบนี้ (กรณีคีย์ก่อนเวลาหรือคีย์ผิด)": "Effacer les saisies de ce tour (saisie anticipée ou erronée)",
  "ล้างคะแนน ใบแดง ลายเซ็น และ Sudden death ทุกคนในรอบที่เลือก กลับเป็นสถานะรอตี · ล้างสาย Knockout ของอีเวนต์นี้ · รายชื่อ ลำดับยิง และสนามไม่เปลี่ยน · ช่องที่เคยคีย์ถูกบันทึกในประวัติการแก้": "Efface scores, cartons rouges, signatures et mort subite de tous pour le tour choisi · Efface le tableau KO · Liste, ordre et terrains inchangés · Les saisies sont consignées dans l’historique",
  "ล้างคะแนนทุกช่อง ใบแดง ลายเซ็น/APPROVED และ Sudden death ของรอบนี้ · ช่องที่เคยคีย์ถูกบันทึกในประวัติการแก้พร้อมเหตุผล": "Efface tous les scores, cartons rouges, signatures/APPROVED et mort subite de ce tour · Consigné dans l’historique avec le motif",
  "ล้างทั้งหมด": "Tout effacer",
  "วันที่บนเอกสาร": "Date sur les documents",
  "วิธีคัดผู้มีสิทธิ์ตีรอบ 2": "Qualification pour le tour 2",
  "วิธีคัดรอบ 2": "Sélection du tour 2",
  "วิธีคัดเข้ารอบ 2": "Sélection du tour 2",
  "ศูนย์จับสลาก": "Centre de tirage",
  "สถานที่": "Lieu",
  "สถานที่ / Location": "Lieu",
  "สถานี": "Atelier",
  "สถิติประกอบการจัดอันดับ 5/3": "Statistiques de classement 5/3",
  "สนาม เริ่มที่สนาม": "terrains, à partir du terrain",
  "เริ่มที่สนาม": "à partir du terrain",
  "สนาม:": "Terrain :",
  "สนามแรกของอีเวนต์": "Premier terrain de l’événement",
  "สร้างธีม": "Créer le thème",
  "สร้างธีมใหม่": "Nouveau thème",
  "สร้างสายคิวใหม่": "Nouvelle chaîne",
  "สร้างสีรองจากสีหลักใหม่": "Régénérer les couleurs secondaires",
  "สร้างหลายอีเวนต์": "Créer plusieurs événements",
  "สร้างหลายอีเวนต์ / นำเข้าจากไฟล์": "Créer plusieurs événements / importer",
  "สร้างหลายอีเวนต์พร้อมกัน": "Créer plusieurs événements",
  "สร้างอีเวนต์ทั้งหมด": "Créer tous les événements",
  "สร้างแถวอัตโนมัติ": "Générer les lignes",
  "สะสมเดิม:": "Cumul précédent :",
  "สิทธิ์รอบ 2": "Éligible tour 2",
  "สีตัวอักษร": "Couleur du texte",
  "สีพื้นหน้า": "Fond de page",
  "สีพื้นอ่อน": "Fond clair",
  "สีหลัก (ปุ่ม/แบนเนอร์)": "Couleur principale (boutons/bannière)",
  "สีหลักเข้ม (หัวข้อ)": "Couleur principale foncée (titres)",
  "สีเน้น": "Couleur d’accent",
  "สีเน้น (เส้นตัด)": "Couleur d’accent (lignes)",
  "สีเส้นกรอบ": "Couleur de bordure",
  "หน้านี้อัปเดตเองทุก 15 วินาที": "Cette page s’actualise toutes les 15 secondes",
  "หน้ารวม: เรียงแถวตามอันดับสด": "Aperçu : trier par classement en direct",
  "หน้ารวมทุกสนาม": "Aperçu de tous les terrains",
  "หมายเหตุ (ไม่บังคับ)": "Remarque (facultatif)",
  "หรือคลิกเพื่อเลือกไฟล์ · PNG, JPG, WEBP · ภาพนี้จะถูกใช้เป็นพื้นหลังแบนเนอร์ด้วย": "ou cliquez pour choisir un fichier · PNG, JPG, WEBP · Cette image sert aussi de fond de bannière",
  "หลายรายการแล้วกดสร้างครั้งเดียว · เมื่อสร้างเสร็จจะไปหน้า": "plusieurs lignes, puis créez-les en une fois · ensuite vous allez à la page",
  "หัวข้อ": "Titre",
  "หัวคอลัมน์สังกัด": "En-tête de colonne affiliation",
  "หัวภาพผลคะแนน · ชื่องาน": "En-tête de l’image · nom de l’événement",
  "หัวภาพผลคะแนน · บรรทัดเล็ก": "En-tête de l’image · petite ligne",
  "หัวภาพผลคะแนน · สถานที่/หน่วยงาน": "En-tête de l’image · lieu/organisation",
  "หากต้องแก้คะแนน ระบบจะเก็บค่าเดิม/ค่าใหม่ ผู้แก้ เวลา และต้องมีลายเซ็นผู้แก้ไข": "Toute modification conserve ancienne/nouvelle valeur, auteur et heure, et exige la signature de l’auteur",
  "หากแก้จำนวนสนามในอีเวนต์ ระบบจะเพิ่มหรือลด Court ID ให้ตรงกับจำนวนสนามโดยอัตโนมัติ": "Si le nombre de terrains change, les identifiants sont ajoutés ou retirés automatiquement",
  "อนุญาตตามระบบ": "Selon le système",
  "อัปโหลดไฟล์ Excel ที่มีคอลัมน์": "Importez un fichier Excel avec les colonnes",
  "อัปโหลดไฟล์ png, jpg, jpeg, webp หรือ gif ได้เอง ถ้าไม่อัปโหลดระบบจะใช้โลโก้ค่าเริ่มต้นตามเอกสารตัวอย่าง": "Importez des fichiers png, jpg, jpeg, webp ou gif ; sinon les logos par défaut sont utilisés",
  "อีเวนต์ที่เลือก ·": "Événements choisis ·",
  "อีเวนต์ในสาย": "Événements de la chaîne",
  "อ่านไฟล์และแสดงตัวอย่าง": "Lire le fichier et prévisualiser",
  "เกิน 10 นาที ตัดออก (DQ)": "Plus de 10 min — disqualifié (DQ)",
  "เกิน 5 นาที −5": "Plus de 5 min −5",
  "เข้ารอบทันทีจากรอบแรก (คน)": "Qualifiés directs du tour 1",
  "เฉพาะหัวข้อที่มีคำว่า": "Seulement les titres contenant",
  "เช่น 3 สนาม เริ่มสนาม 5 = ใช้สนาม 5-7 · ลำดับตีไม่เปลี่ยน": "ex. 3 terrains à partir du 5 = terrains 5–7 · L’ordre de tir ne change pas",
  "เช่น": "ex.",
  "เดิม": "Ancien",
  "เปิด": "Activé",
  "เปิด — แถวสลับตามคะแนน": "Activé — tri par score",
  "เปิดรอบ 2": "Tour 2 activé",
  "เปิดใช้ทั้งระบบ": "Utiliser partout",
  "เพื่อกันคะแนนทับกัน หน้านี้จึงดูได้อย่างเดียว ถ้าเครื่องเดิมเลิกใช้แล้วให้กดรับช่วงคีย์ต่อ": "Pour éviter d’écraser les scores, cette page est en lecture seule. Si l’autre appareil n’est plus utilisé, reprenez la saisie.",
  "เพื่อยืนยัน": "pour confirmer",
  "เพื่อสั่งจับพร้อมกัน": "pour tirer ensemble",
  "เพื่อให้หลายรุ่นตีต่อกันบนสนามชุดเดียว": "pour enchaîner plusieurs groupes sur les mêmes terrains",
  "เริ่มจากสนาม 1 · คะแนนรอบ 1 น้อยสุดตีก่อน (เท่ากันดูจำนวน 5 แล้วจำนวน 3 น้อยกว่าตีก่อน)": "Commence au terrain 1 · Le plus petit score du tour 1 tire d’abord (égalité : moins de 5, puis moins de 3)",
  "เริ่มตีแล้ว": "Commencé",
  "เริ่มตีแล้ว · ล็อก": "Commencé · verrouillé",
  "เรียงตาม TOTAL → จำนวน 5 → จำนวน 3 ใช้ดูเหตุผลการจัดลำดับ ไม่รวมคะแนนรอบพิเศษในตารางหลัก": "Trié par TOTAL → nombre de 5 → nombre de 3 pour expliquer le classement ; tours supplémentaires exclus",
  "เรียงตามชื่อ (12→18, ชาย→หญิง)": "Trier par nom (12→18, hommes→femmes)",
  "เรียงตามอันดับ: เปิด": "Tri par classement : activé",
  "เรียงตามอันดับสด": "Tri par classement en direct",
  "เลือก": "Choisir",
  "เลือกทั้งหมดที่แสดง": "Tout sélectionner (affichés)",
  "เลือกทีมเข้ารอบ 2": "Choisir les équipes du tour 2",
  "เลือกที่พร้อมจับ (ที่แสดง)": "Choisir ceux prêts au tirage (affichés)",
  "เลือกทุกอีเวนต์": "Sélectionner tous les événements",
  "เลือกภาพของงาน": "Choisir l’image de l’événement",
  "เลือกรายการใหม่": "Choisir d’autres événements",
  "เลือกสถานีสำหรับหน้าจอแนวตั้ง": "Choisir l’atelier (écran portrait)",
  "เลือกสนาม": "Choisir le terrain",
  "เลือกหลายอีเวนต์แล้วกดจับสลากครั้งเดียว ระบบสุ่มลำดับยิงและกระจายสนามใหม่ให้ทุกอีเวนต์พร้อมกัน · อีเวนต์ที่เริ่มกรอกคะแนนแล้วจะถูกล็อกไว้": "Choisissez plusieurs événements et tirez en une fois : ordre et terrains sont tirés pour tous · Les événements déjà commencés sont verrouillés",
  "เลือกอีเวนต์ที่ใช้สนามชุดเดียวกัน แล้วเรียงลำดับรุ่นที่จะตีก่อน–หลัง · ลำดับยิงในแต่ละรุ่นมาจากการจับสลากของรุ่นนั้น": "Choisissez les événements partageant les terrains puis ordonnez les groupes · L’ordre de tir vient du tirage de chaque groupe",
  "เลือกแล้ว": "Sélectionnés",
  "เลือกได้ 2 วิธี:": "Deux façons :",
  "เลื่อนขึ้น": "Monter",
  "เลื่อนลง": "Descendre",
  "เวลา": "Heure",
  "เว้นว่าง = ทุกประเภท": "Vide = toutes les catégories",
  "เส้นตัด: เปิด": "Lignes : activées",
  "เส้นแบ่งเข้ารอบ": "Lignes de qualification",
  "เหตุผล": "Motif",
  "เหตุผล เช่น คีย์ก่อนนักกีฬาตีจริง": "Motif, ex. saisie avant le tir réel",
  "เหตุผลในการแก้ (ถ้ามี)": "Motif de la modification (facultatif)",
  "แตะคะแนนของแต่ละระยะ · บันทึกอัตโนมัติทุกครั้งที่แตะ": "Touchez le score de chaque distance · Enregistré à chaque touche",
  "แถบโลโก้ผู้สนับสนุน": "Bandeau des partenaires",
  "แยก FAMILY / GIVEN — คำสุดท้ายเป็นนามสกุล (แบบไทย)": "Séparer FAMILY / GIVEN — dernier mot = nom (style thaï)",
  "แยก FAMILY / GIVEN — คำแรกเป็นนามสกุล (แบบสากล)": "Séparer FAMILY / GIVEN — premier mot = nom (international)",
  "และ": "et",
  "แสดง": "Afficher",
  "แสดงผัง": "Afficher le plan",
  "แสดงหน้ากรรมการ": "Afficher la page des officiels",
  "แสดงเส้นแบ่ง": "Afficher les lignes",
  "แสดงแบนเนอร์ด้านบนของเว็บ": "Afficher la bannière en haut du site",
  "โลโก้งาน": "Logo de l’événement",
  "โลโก้มุมซ้ายในหน้าผล": "Logo du coin gauche (pages de résultats)",
  "โลโก้ล่างปก 1": "Logo bas de couverture 1",
  "โลโก้ล่างปก 2": "Logo bas de couverture 2",
  "โลโก้ล่างปก 3": "Logo bas de couverture 3",
  "โลโก้หลักบนปก": "Logo principal de couverture",
  "โลโก้หัวกระดาษ 1": "Logo d’en-tête 1",
  "โลโก้หัวกระดาษ 2": "Logo d’en-tête 2",
  "โลโก้หัวกระดาษ 3": "Logo d’en-tête 3",
  "โลโก้หัวกระดาษ 4": "Logo d’en-tête 4",
  "โลโก้เอกสาร": "Logos des documents",
  "ใช้คำนวณจำนวนสนาม (0 = ใช้ค่าคงที่)": "Sert au calcul du nombre de terrains (0 = valeur fixe)",
  "ใช้ค่าเริ่มต้น": "Valeur par défaut",
  "ใช้ได้ครั้งละ 1 ธีมทั้งเว็บ (แบนเนอร์ สีปุ่ม/เมนู และหัว–ท้ายภาพผลคะแนนที่บันทึกจากหน้ารวม) · สร้างธีมใหม่จากภาพของงานได้ ระบบดึงสีจากภาพให้อัตโนมัติ": "Un seul thème actif pour tout le site (bannière, couleurs, en-tête/pied des images) · Créez un thème depuis une image ; les couleurs sont extraites automatiquement",
  "ใบนี้กำลังถูกคีย์โดย": "Cette feuille est en cours de saisie par",
  "ใบแดง": "Carton rouge",
  "ใบแดงครบ 2 ครั้ง: ยุติการยิงรอบนี้และคงคะแนนที่ทำได้ไว้": "2 cartons rouges : fin du tir pour ce tour, le score acquis est conservé",
  "ใส่ชื่อจากไฟล์": "Remplir les noms depuis le fichier",
  "ใส่บรรทัดละ 1 คน รูปแบบ: ชื่อ | สังกัด เช่น MR. SOMCHAI JAIDEE | ABSC": "Une personne par ligne : nom | fédération, ex. MR. SOMCHAI JAIDEE | ABSC",
  "ใหม่": "Nouveau",
  "ให้หลายอีเวนต์ตีต่อกันบนสนามชุดเดียว เมื่อรุ่นหนึ่งเหลือคนไม่เต็มสนาม รุ่นถัดไปเข้าตีสนามที่ว่างได้ทันที · อันดับและสิทธิ์เข้ารอบยังแยกต่ออีเวนต์ · ใช้กับรอบ 1 · อีเวนต์ที่ไม่ได้เลือกยังจัดสนามแยกเหมือนเดิม": "Plusieurs événements s’enchaînent sur les mêmes terrains : quand un groupe ne remplit plus tous les terrains, le suivant prend les terrains libres · Classements et qualifications restent par événement · Tour 1 uniquement",
  "ไฟล์ค่าล้วน ไม่มีสูตร แก้ไขได้ทุกช่อง": "Valeurs uniquement, sans formules — toutes les cellules modifiables",
  "ไฟล์ที่มีหัวข้อแต่ละประเภท เช่น": "Un fichier avec un titre par catégorie, ex.",
  "ไฟล์ใบสมัคร": "Fichier d’inscription",
  "ไม่บังคับ": "Facultatif",
  "ไม่มาทำการแข่งขันตามเวลา": "Absent à l’heure prévue",
  "ไม่มีบทลงโทษ": "Aucune pénalité",
  "ไม่มีรอบ 2": "Pas de tour 2",
  "ไม่ล้างคนที่ APPROVED แล้ว": "Ne pas effacer les athlètes déjà APPROVED",
  "ไม่เลือก": "Non choisi",
  "ไม่แสดง": "Masquer",
  "← กลับ Overview": "← Retour à l’aperçu",
  "← กลับ scorecard": "← Retour à la feuille",
  "← กลับหน้าประกบคู่": "← Retour au tableau",
  "← ก่อนหน้า": "← Précédent",
  "↩ เลิกซ่อนที่เลือก": "↩ Réafficher la sélection",
  "⚙ ตั้งค่า Results Approved / กรรมการ": "⚙ Paramètres résultats validés / officiels",
  "⚙️ ตั้งค่าร่วมกัน": "⚙️ Paramètres communs",
  "⚠️ ผลรอบนี้ลงลายเซ็นยืนยันแล้ว": "⚠️ Ce tour a été signé et validé",
  "✓ ระบบลงสนามตามผังนี้อยู่แล้ว": "✓ Les terrains suivent déjà ce plan",
  "🎲 จับสลากพร้อมกัน": "🎲 Tirer la sélection",
  "🎲 ศูนย์จับสลาก": "🎲 Centre de tirage",
  "👀 เปิดดูรวมกัน": "👀 Voir ensemble",
  "👁 แสดงเฉพาะที่เลือก": "👁 Afficher seulement la sélection",
  "📄 Excel ไม่มีสูตร": "📄 Excel (sans formules)",
  "📄 QF / Final ไม่มีสูตร": "📄 QF / Finale (sans formules)",
  "📊 ดูสถิติ 5/3": "📊 Statistiques 5/3",
  "🔗 ต่อคิวสนาม": "🔗 Chaînes de terrains",
  "🔗 ต่อคิวสนามข้ามรุ่น": "🔗 Chaînes de terrains entre groupes",
  "🖨 พิมพ์ Score Sheet รอบนี้": "🖨 Imprimer les feuilles de ce tour",
  "🗂 ผังเข้ารอบ 2": "🗂 Plan du tour 2",
  "🙈 ซ่อนที่เลือก": "🙈 Masquer la sélection",
  "ซ่อนอยู่": "Masqué",
  "ซ่อนไว้": "Masqués :",
  "แสดงที่ซ่อนไว้": "Afficher les masqués",
  "ซ่อนจากรายการ": "Masquer de la liste"
 },
 "zh": {
  "ชื่อผู้ใช้ / สิทธิ์": "用户名 / 权限",
  "บัญชีที่กำลังใช้อยู่": "当前账号",
  "ลบบัญชีที่กำลังใช้อยู่ไม่ได้": "不能删除正在使用的账号",
  "ต้องมี superadmin อย่างน้อย 1 บัญชี": "至少需要 1 个 superadmin 账号",
  "กรุณาใส่ชื่อผู้ใช้": "请输入用户名",
  "เปลี่ยนสิทธิ์ของบัญชีที่กำลังใช้อยู่ไม่ได้": "不能更改正在使用账号的权限",
  "ไม่พบผู้ใช้ (บัญชีสนามให้จัดการจากหน้าอีเวนต์)": "未找到用户（场地账号请在赛事页面管理）",
  "รูปแบบการคีย์": "录入方式",
  "📋 ตาราง (แบบเดิม)": "📋 表格（原版）",
  "👆 ปุ่มแตะ (แบบรอบน็อค)": "👆 点按按钮（淘汰赛样式）",
  "รวมรอบนี้": "本轮合计",
  "สถานีสุดท้าย": "最后一站",
  "สถานีถัดไป →": "下一站 →",
  "กรุณาเลือกนักกีฬาที่ต้องตี Sudden death อย่างน้อย 2 คน": "请至少选择 2 名进行突然死亡的运动员",
  "กรุณาเลือกนักกีฬาอย่างน้อย 1 คน": "请至少选择 1 名运动员",
  "คุณไม่มีสิทธิ์ทำรายการนี้": "您无权执行此操作",
  "จบการตีเรียบร้อย": "击球已结束",
  "ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง": "用户名或密码错误",
  "ต้องลงชื่ออย่างใดอย่างหนึ่ง (พิมพ์ชื่อหรือเขียน) ให้ครบทั้ง 3 ฝ่าย หรือใช้สิทธิ์ข้าม": "三方都需签名（输入姓名或手写），或使用跳过权限",
  "นักกีฬาคนนี้ไม่มีสิทธิ์ตีรอบ 2": "该运动员无第2轮资格",
  "บันทึกผลเที่ยวพิเศษแล้ว ถ้ายังเท่ากันให้บันทึกเที่ยวพิเศษเพิ่มอีกครั้ง": "加赛成绩已保存；如仍平分请再录入一轮加赛",
  "บันทึกผู้ชนะแล้ว": "已保存胜者",
  "ผู้ชนะไม่ถูกต้อง": "胜者无效",
  "ยังไม่ได้ติดตั้ง python-docx ให้รัน: python -m pip install -r requirements.txt": "尚未安装 python-docx，请运行：python -m pip install -r requirements.txt",
  "รองรับเฉพาะไฟล์ .xlsx": "仅支持 .xlsx 文件",
  "รอบแข่งขันไม่ถูกต้อง": "比赛轮次无效",
  "ลบอีเวนต์แล้ว": "赛事已删除",
  "สร้างผู้ใช้สำเร็จ": "用户已创建",
  "สร้างผู้ใช้ไม่สำเร็จ กรุณาตรวจสอบข้อมูล": "创建用户失败，请检查信息",
  "สุ่มลำดับใหม่แล้ว": "已重新随机顺序",
  "ออกจากระบบแล้ว": "已退出登录",
  "เข้าสู่ระบบสำเร็จ": "登录成功",
  "เพิ่มนักกีฬาสำเร็จ": "已添加运动员",
  "แก้ไขอีเวนต์สำเร็จ": "赛事已修改",
  "ปิดการเรียงสดแล้ว · แถวอยู่ตามลำดับการตี": "已关闭实时排序 · 行按击球顺序",
  "กรุณาระบุเหตุผลในการล้างการคีย์": "请填写清除录入的原因",
  "กรุณาเพิ่มอย่างน้อย 1 อีเวนต์": "请至少添加 1 个赛事",
  "กรุณาเลือกอีเวนต์": "请选择赛事",
  "กรุณาเลือกอีเวนต์ที่จะจับสลาก": "请选择要抽签的赛事",
  "กรุณาเลือกอีเวนต์อย่างน้อย 1 รายการ": "请至少选择 1 个赛事",
  "กรุณาเลือกไฟล์": "请选择文件",
  "ข้ามอีเวนต์ที่เริ่มกรอกคะแนนแล้ว:": "已跳过已开始录分的赛事：",
  "คะแนนเท่ากัน · ตีเที่ยวพิเศษ 7 เมตรทุกสถานี": "平分 · 所有站点 7 米加赛",
  "คู่นี้ยังไม่ครบ 2 คน": "此对阵尚未满 2 人",
  "จบการตีคู่นี้แล้ว": "此对阵已结束",
  "ธีมที่มากับระบบลบไม่ได้ (แก้ไขหรือคัดลอกได้)": "系统内置主题不能删除（可编辑或复制）",
  "บัญชีนี้บันทึกบทลงโทษได้เฉพาะนักกีฬาในสนามของตัวเอง": "此账号只能对本场地运动员记录处罚",
  "บัญชีนี้ยังใช้รหัสผ่านตั้งต้นที่เปิดเผยในโค้ด จึงถูกระงับ ให้ superadmin ตั้งรหัสใหม่ที่หน้าผู้ใช้": "此账号仍使用代码中的默认密码，已被停用，请 superadmin 在用户页重设密码",
  "บัญชีสนามกรอกเที่ยวพิเศษได้เฉพาะคู่ Knockout ในสนามตัวเอง": "场地账号只能录入本场地淘汰赛的加赛",
  "บัญชีสนามนี้คีย์ได้เฉพาะคู่ในสนามตัวเอง": "此场地账号只能录入本场地的对阵",
  "บัญชีสนามให้สร้างจากหน้านักกีฬาของอีเวนต์": "场地账号请在赛事运动员页面创建",
  "บันทึก Sudden death แล้ว": "已保存突然死亡",
  "พิมพ์คำว่า ล้าง ในช่องยืนยันก่อน": "请先在确认框输入 ล้าง",
  "มี court ID นี้มากกว่า 1 อีเวนต์ที่ใช้รหัสเดียวกัน กรุณาเปลี่ยนรหัสของอีเวนต์ให้ไม่ซ้ำ": "此场地账号在多个赛事使用相同密码，请改为不同密码",
  "มีอีเวนต์ที่เริ่มกรอกคะแนนแล้ว การจัดสนามใหม่จะย้ายสนามของนักกีฬา · ติ๊กยืนยันก่อนบันทึก": "部分赛事已开始录分，重新分配会移动运动员场地 · 保存前请勾选确认",
  "ยังยิงไม่ครบทุกสถานี จึงยัง Approve ไม่ได้": "尚未完成所有站点，暂不能确认",
  "ยังไม่ได้ติ๊กหัวข้อการตั้งค่าที่ต้องการใช้ร่วมกัน": "尚未勾选要共用的设置",
  "ยังไม่ได้เลือกประเภทที่จะสร้าง": "尚未选择要创建的类别",
  "รหัสข้ามลายเซ็นไม่ถูกต้อง · หรือไปลงลายเซ็นที่ Scorecard รายคน": "跳过签名代码错误 · 或到各记分卡签名",
  "รองรับไฟล์ .xlsx และ .xls": "支持 .xlsx 和 .xls 文件",
  "ระดับบทลงโทษไม่ถูกต้อง": "处罚级别无效",
  "ลบธีมที่กำลังใช้งานไม่ได้ ให้เปิดใช้ธีมอื่นก่อน": "不能删除正在使用的主题，请先启用其他主题",
  "ลบอีเวนต์ไม่สำเร็จ กรุณาลองใหม่หรือแจ้งผู้ดูแลระบบ": "删除赛事失败，请重试或联系管理员",
  "ล้าง Sudden death ของกลุ่มนี้แล้ว": "已清除该组突然死亡",
  "สิทธิ์ไม่ถูกต้อง": "权限无效",
  "อีเวนต์นี้มีผล Knockout แล้ว · ตรวจสาย Knockout อีกครั้งหลังคีย์ใหม่": "本赛事已有淘汰赛成绩 · 重新录入后请再检查对阵",
  "อีเวนต์นี้อยู่ในสายอื่นแล้ว:": "本赛事已在其他链中：",
  "อีเวนต์นี้ไม่ได้เปิดใช้รอบ 2": "本赛事未开启第2轮",
  "อ่านไฟล์ไม่สำเร็จ:": "读取文件失败：",
  "เปิดการเรียงตามอันดับสดแล้ว": "已开启按实时排名排序",
  "ปิดการเรียงตามอันดับสดแล้ว": "已关闭按实时排名排序",
  "เลือกอย่างน้อย 2 อีเวนต์เพื่อต่อคิวกัน": "请至少选择 2 个赛事来串联",
  "แก้ไข/ล้าง Sudden death ได้เฉพาะ superadmin": "仅 superadmin 可编辑/清除突然死亡",
  "แสดงเส้นตัดแล้ว": "已显示分界线",
  "ซ่อนเส้นตัดแล้ว": "已隐藏分界线",
  "ไม่พบผู้ใช้ (บัญชีสนามให้เปลี่ยนรหัสจากหน้าอีเวนต์)": "未找到用户（场地账号请在赛事页面改密码）",
  "ไม่พบหัวข้อประเภทที่มีรายชื่อในไฟล์": "文件中未找到带名单的类别标题",
  "ไม่มีสิทธิ์ดูประวัติคะแนนของอีเวนต์หรือสนามอื่น": "无权查看其他赛事或场地的分数记录",
  "บันทึกตั้งค่า Results Approved แล้ว": "成绩确认设置已保存",
  "ล้างการคีย์": "清除录入：",
  "(จางๆ หลังแบนเนอร์)": "（淡化于横幅背后）",
  "(สร้างอีเวนต์พร้อมรายชื่อ) หรือ": "（创建赛事并导入名单）或",
  "(แนวนอน กว้าง)": "（横向，宽幅）",
  "+ สร้างธีมใหม่จากภาพ": "+ 从图片创建新主题",
  "+ สร้างสายคิวใหม่": "+ 新建场地队列链",
  "+ สร้างหลายอีเวนต์": "+ 批量创建赛事",
  "+ เพิ่ม 1 แถว": "+ 添加 1 行",
  "+ เพิ่มแถวตามที่เลือก": "+ 按所选添加行",
  "(ยังไม่ครบ)": "（未完成）",
  "1 · โยนภาพของงาน ให้ระบบสร้างสีธีม": "1 · 拖入赛事图片，系统生成主题配色",
  "อัปเดตหน้าอัตโนมัติทุก 15 วินาที": "每 15 秒自动刷新",
  "1) เลือกอีเวนต์ (เรียงตามลำดับที่จะตี) · 2) ดูผัง สนาม × ลำดับคิว · 3) กด \"ลงสนามตามผังนี้\" กรรมการสนามจะเห็นคิวรอบ 2 ตามผังทันที · ลำดับในรอบ 2: คะแนนรอบ 1 น้อยสุดตีก่อน เริ่มจากสนาม 1": "1) 选择赛事（按比赛顺序）· 2) 查看场地 × 队列图 · 3) 点击“按此图分配场地”，场地裁判立即看到第2轮队列 · 第2轮顺序：第1轮得分最低者先投，从1号场地开始",
  "2 · สี": "2 · 颜色",
  "3 · ภาพ": "3 · 图片",
  "4 · ข้อความ": "4 · 文字",
  "Court ID ของอีเวนต์นี้": "本赛事场地账号",
  "ID สำหรับล็อกอิน": "登录账号",
  "ผ่านตรงไม่ต้องตีรอบ 2:": "直接晋级（无需第2轮）：",
  "คีย์คู่": "对阵录入",
  "กดเพื่อซ่อนเส้นตัด QUARTERFINALS / QUALIFIED FOR ROUND 2": "点击隐藏 QUARTERFINALS / QUALIFIED FOR ROUND 2 分界线",
  "กดเพื่อปิด: แถวจะอยู่ตามลำดับการตี ไม่สลับตามคะแนน": "点击关闭：行保持击球顺序，不按得分重排",
  "กรรมการยกคะแนน": "计分裁判",
  "กรอกตาราง": "填写表格",
  "กรอกทีละสถานี": "逐站录入",
  "กรอกหลายอีเวนต์ในตาราง": "在表格中填写多个赛事",
  "กรองวันแข่งขัน": "按比赛日期筛选",
  "กลับ": "返回",
  "กลับ Overview": "返回总览",
  "กลับ Results Approved": "返回成绩确认",
  "กลับ Scorecard": "返回记分卡",
  "กลับหน้ารายการ": "返回列表",
  "กลับหน้าอีเวนต์": "返回赛事",
  "กลับไปหน้าธีม": "返回主题",
  "การจัดลำดับ": "排序",
  "กำลังแข่งขัน": "比赛中",
  "กำลังโหลด...": "加载中...",
  "กำลังใช้": "使用中",
  "กำหนด / เปลี่ยนรหัส": "设置 / 更改密码",
  "ขึ้น": "上移",
  "ข้อความ APPROVED": "APPROVED 文字",
  "ข้อความใหญ่บนปก (เว้นว่าง = ใช้ชื่องานหน้าเครื่องหมาย ·)": "封面大字（留空 = 使用 · 前的赛事名称）",
  "คงลำดับการตี": "保持击球顺序",
  "คนต่อสนามไม่เกิน": "每场地最多人数",
  "คลิกที่ทีมเพื่อเลือก/ไม่เลือก แล้วกดบันทึก · ผู้ผ่านตรงไม่อยู่ในรายการนี้": "点击队伍选择/取消，然后保存 · 直接晋级者不在此列表",
  "คอลัมน์ NAME เดียว (ชื่อเต็มตามที่ลงทะเบียน)": "单一 NAME 列（注册全名）",
  "คะแนนเท่ากัน: ตี 7 เมตรทุกสถานีอีก 1 เที่ยว บันทึกพร้อมกันทุกคน · ถ้ายังเท่ากันระบบจะให้ตี Sudden death ต่อ": "平分：所有站点再投一轮 7 米，所有人一起录入 · 仍平分则进入突然死亡",
  "คัดลอก": "复制",
  "คำนวณสีเข้ม/พื้นอ่อน/เส้นกรอบ/ตัวอักษร ให้เข้ากับสีหลัก และให้ตัวหนังสือสีขาวบนปุ่มยังอ่านง่าย": "根据主色计算深色/浅底/边框/文字颜色，并保证按钮白字清晰",
  "คำนำหน้าชื่องาน": "赛事名称前缀",
  "คีย์คู่นี้": "录入此对阵",
  "ค่า N": "N 值",
  "ค้นหาชื่องาน / รุ่น / ประเภท": "搜索赛事 / 组别 / 类别",
  "ค้นหาอีเวนต์": "搜索赛事",
  "จบการตีทั้งคู่": "两人均结束",
  "จากรอบ 2 คัดเข้ารอบต่อ (คน)": "第二轮晋级人数",
  "จำนวนคะแนน 3 รวมรอบ 1 และรอบ 2": "3分次数（第1、2轮合计）",
  "จำนวนคะแนน 5 รวมรอบ 1 และรอบ 2": "5分次数（第1、2轮合计）",
  "จำนวนผู้ผ่านเข้ารอบถัดไปทันทีจากรอบแรก": "第一轮直接晋级人数",
  "จำนวนสนาม (ค่าคงที่)": "场地数量（固定）",
  "จำนวนสนาม / สนามแรก": "场地数 / 起始场地",
  "จำนวนสนามเริ่มต้นต่อแถว": "每行默认场地数",
  "ชื่อ Technical Delegate": "技术代表姓名",
  "ชื่อ:": "姓名：",
  "ชื่อการแข่งขัน / Competition Title": "比赛名称",
  "ชื่อธีม (เห็นเฉพาะผู้ดูแล)": "主题名称（仅管理员可见）",
  "ชื่อนักกีฬา (ใส่ระบบ)": "运动员姓名（导入）",
  "ชื่อประธาน / President Name": "主席姓名",
  "ชื่อประเภท": "类别名称",
  "ชื่อย่องาน": "赛事简称",
  "ชื่อรอบถัดไป": "下一轮名称",
  "ชื่อสนาม": "场地名称",
  "ชื่อสาย": "链名称",
  "ชื่อใหญ่บนแบนเนอร์": "横幅大标题",
  "ซ่อน": "隐藏",
  "ซ่อนเส้นแบ่ง": "隐藏分界线",
  "ดาวน์โหลด DOCX": "下载 DOCX",
  "ดูตัวอย่างด้านขวาก่อน แล้วติ๊ก “เปิดใช้ทั้งระบบ” เมื่อพร้อม": "先查看右侧预览，准备好后勾选“全站启用”",
  "ดูผลหลายอีเวนต์": "查看多个赛事成绩",
  "ดูเต็มจอ →": "全屏 →",
  "ด้านล่าง ระบบจะพาไปหน้าตีเที่ยวพิเศษ 7 เมตรทันที": "下方，系统将立即进入 7 米加赛页面",
  "ตัวหนา = ผู้ชนะ": "粗体 = 胜者",
  "ตัวอย่าง": "预览",
  "ตัวอย่างคิวสนาม": "场地队列预览",
  "ตัวอักษร (1=A … 8=H)": "字母（1=A … 8=H）",
  "ตัวอักษร (1=A, 2=B … 8=H)": "字母（1=A, 2=B … 8=H）",
  "ตัวเลข": "数字",
  "ตัวเลข (1, 2, 3 …)": "数字（1, 2, 3 …）",
  "ตั้งค่า Results": "成绩设置",
  "ตั้งค่า Results Approved": "成绩确认设置",
  "ตั้งค่าข้อความหัวเอกสารและรายชื่อกรรมการสำหรับ": "设置文件页眉文字和裁判名单：",
  "ตั้งค่าร่วม (ใช้กับทุกอีเวนต์ในชุดนี้)": "共同设置（适用于本组所有赛事）",
  "ตั้งค่าหลายอีเวนต์": "多赛事设置",
  "ตั้งค่าหลายอีเวนต์พร้อมกัน": "同时设置多个赛事",
  "ตั้งรหัส": "设置",
  "ตั้งรหัสผ่านใหม่": "设置新密码",
  "ตามด้วยตาราง": "后接表格",
  "ตำแหน่ง Technical Delegate": "技术代表职务",
  "ตำแหน่งประธาน / President Title": "主席职务",
  "ติ๊กเฉพาะหัวข้อที่ต้องการใช้เหมือนกัน ค่าอื่นของแต่ละอีเวนต์จะไม่ถูกเปลี่ยน": "只勾选需要统一的项目，其他赛事设置不变",
  "ตี 7 เมตร →": "7 米加赛 →",
  "ตีครบ 5 สถานีแล้วคะแนนเท่ากัน · กด": "5 个站点后平分 · 点击",
  "ตีแล้ว": "已投",
  "ต่อคิวสนามข้ามรุ่น": "跨组场地队列",
  "ต่อจากผู้ผ่านตรงอีก N ลำดับ": "直接晋级者之后再取 N 名",
  "ต่อจากผู้ผ่านตรงอีก N ลำดับ (เท่ากันที่ลำดับสุดท้าย ตี Sudden death)": "直接晋级者之后再取 N 名（最后一名平分 → 突然死亡）",
  "ต้องการสร้างหลายรายการพร้อมกัน?": "需要同时创建多个赛事？",
  "ถอนตัว": "退赛",
  "ถึงลำดับ (Class) ที่ N": "至名次（Class）N",
  "ที่ · ชื่ออปท. · อำเภอ · จังหวัด": "序号 · 单位名称 · 区 · 府",
  "ทุกวันแข่งขัน": "所有日期",
  "ท้ายภาพผลคะแนน · ชื่องาน": "成绩图页脚 · 赛事名称",
  "ท้ายภาพผลคะแนน · สโลแกน": "成绩图页脚 · 口号",
  "ธีม": "主题",
  "ธีมของระบบ": "系统主题",
  "นักกีฬา/ทีม": "运动员/队伍",
  "นำเข้าจากไฟล์ใบสมัคร": "从报名文件导入",
  "นำเข้าจากไฟล์ใบสมัคร (.xls / .xlsx)": "从报名文件导入（.xls / .xlsx）",
  "นำเข้าชื่อ": "导入姓名",
  "นำเข้าชื่อนักกีฬา": "导入运动员姓名",
  "บรรทัดประเทศ/เจ้าภาพ": "国家/东道主行",
  "บรรทัดรอง": "副标题",
  "บรรทัดเล็กด้านบน": "顶部小字",
  "บังคับจับใหม่แม้เริ่มตีแล้ว (ล้างสาย Knockout)": "即使已开始也强制重新抽签（清除淘汰赛对阵）",
  "บัญชีกรรมการสนาม (court01, court02 …): ตั้งรหัสที่อีเวนต์ใดก็ได้ในสาย กรรมการคนเดิมจะเห็นคิวของสนามตัวเองต่อกันทุกรุ่นในสาย": "场地裁判账号（court01, court02 …）：在链中任一赛事设置密码，同一裁判可连续看到链内各组在本场地的队列",
  "บันทึกตั้งค่า": "保存设置",
  "บันทึกทีมที่เลือก": "保存所选队伍",
  "บันทึกพร้อมกัน": "一起保存",
  "บันทึกลายเซ็น": "保存签名",
  "บันทึกสายคิว": "保存队列链",
  "บันทึกเป็น JPG": "保存为 JPG",
  "บันทึกเป็น JPG ทั้งหมด": "全部保存为 JPG",
  "บันทึกแล้วระบบจัดสนามให้ทุกอีเวนต์ในสายทันที": "保存后系统立即为链内所有赛事分配场地",
  "บันทึกให้ 1 อีเวนต์": "保存到 1 个赛事",
  "ประวัติการแก้คะแนน": "改分记录",
  "ประวัติรอบคัดเลือก": "资格赛记录",
  "ปิด (คงลำดับการตี)": "关闭（保持击球顺序）",
  "ปิด — คงลำดับการตี": "关闭 — 保持击球顺序",
  "ปิดรอบ 2": "关闭第2轮",
  "ปิดเมื่อรายการไม่ต้องการให้แถวสลับไปมา · อันดับ (Class) ยังแสดงตามปกติ": "不希望行来回变动时关闭 · 名次（Class）仍正常显示",
  "ปุ่มรอง": "次要按钮",
  "ผลคะแนนรอบนี้ได้รับการรับรองแล้ว": "本轮成绩已确认",
  "ผลรวมหลายอีเวนต์": "多赛事汇总",
  "ผังรอบ 2": "第2轮场地图",
  "ผังลงสนามรอบ 2": "第2轮场地分配图",
  "ผู้ผ่านจากรอบ 2": "第2轮晋级者",
  "ผู้ผ่านตรง": "直接晋级者",
  "ผู้แก้": "修改人",
  "ผู้แข่งขัน": "参赛者",
  "พร้อมคีย์": "可录入",
  "พร้อมบันทึก": "就绪",
  "พิมพ์คำว่า": "输入",
  "พิมพ์ผัง": "打印场地图",
  "ภาพพื้นหลังแบนเนอร์": "横幅背景图",
  "มากับระบบ": "系统内置",
  "มาตรฐาน (ทั่วไป)": "标准（通用）",
  "มาตรฐาน (ใบเดิมของระบบ)": "标准（系统原版）",
  "มีรอบ 2": "有第2轮",
  "มีอีเวนต์ที่เริ่มกรอกคะแนนแล้ว การบันทึกจะย้ายสนามของนักกีฬา (คะแนนไม่หาย)": "部分赛事已开始录分；保存会调整运动员场地（分数保留）",
  "มือถือ/แท็บเล็ตแนวตั้งจะแสดงเฉพาะสถานีที่เลือก เพื่อให้ช่องใหญ่ขึ้นและกดง่ายขึ้น": "手机/平板竖屏只显示所选站点，格子更大更易点击",
  "ยังเสมอ": "仍平分",
  "ยังไม่ตั้งรหัส": "尚未设置密码",
  "ยังไม่มีการแก้คะแนนหลังลงลายเซ็น": "签名后暂无改分",
  "ยังไม่มีสายคิว ·": "暂无场地队列链 ·",
  "ยังไม่ได้ลงลายเซ็นผู้แก้ไข": "修改人尚未签名",
  "ยังไม่ได้เลือกอีเวนต์": "尚未选择赛事",
  "ยิงลูกกลางขาว": "中间白球射击",
  "ยิงลูกสีขาวเหนือลูกดำ": "白球越过黑球射击",
  "ยิงลูกเดี่ยว": "单球射击",
  "ยิงลูกเป้า": "目标球射击",
  "ยิงลูกเหนือลูกเป้า": "越过目标球射击",
  "ยืนยันจัดสนามใหม่": "确认重新分配场地",
  "รวมเที่ยวนี้": "本轮合计",
  "รหัสของ": "密码：",
  "รหัสผ่าน (อย่างน้อย 8 ตัว)": "密码（至少 8 位）",
  "รหัสใหม่": "新密码",
  "รอคู่แข่งขัน": "等待对手",
  "รอตี": "等待投掷",
  "เลือกคะแนนแบบ 0 / 1 / 3 / 5": "选择 0 / 1 / 3 / 5 分",
  "รอบ 1 และรอบ 2": "第1轮和第2轮",
  "รอบ 16 คน": "十六强",
  "รอบการแข่งขัน": "比赛轮次",
  "รอบที่จะล้าง": "要清除的轮次",
  "รอแข่งขัน": "等待比赛",
  "ระบบคัดอัตโนมัติ": "系统自动选拔",
  "ระบบจะแยกเป็น 1 อีเวนต์ต่อ 1 หัวข้อ และให้ตรวจสอบก่อนสร้างจริง": "每个标题生成一个赛事，创建前可先检查",
  "ระบบจะใส่ชื่อให้นักกีฬาตาม id\n    ไม่เปลี่ยนสังกัด ลำดับตี สนาม หรือคะแนน · แถวที่ชื่อว่างจะข้าม · อัปโหลดซ้ำได้": "按运动员 id 填入姓名，不更改单位、顺序、场地或分数 · 空姓名行跳过 · 可重复上传",
  "ระบบบันทึกคะแนนและจัดอันดับเปตอง Shooting": "法式滚球射击计分与排名系统",
  "ระบบเตรียม ID ตามจำนวนสนาม": "系统按场地数量自动准备账号：",
  "สนามให้อัตโนมัติ · กำหนดรหัสเมื่อพร้อมใช้งาน": "个场地 · 准备使用时设置密码",
  "ระบบเลือกสีหลักและสีเน้นให้แล้ว · กด “หลัก” หรือ “เน้น” ใต้สีเพื่อเปลี่ยน": "系统已选好主色和强调色 · 点击颜色下方“主色”或“强调”可更改",
  "ระบบใช้ตามรอบถัดไปอัตโนมัติ: รอบ 8 = 4, รอบ 16 = 8": "按下一轮自动设定：八强 = 4，十六强 = 8",
  "ระยะ": "距离",
  "รายการทั้งหมด": "全部赛事",
  "รายการอีเวนต์": "赛事列表",
  "รายชื่อ": "名单",
  "รายชื่อกรรมการ Umpire": "裁判名单",
  "รูปแบบชื่อในเอกสาร": "文件中姓名格式",
  "รูปแบบใบ:": "记分卡样式：",
  "รูปแบบใบบันทึกคะแนน": "记分卡样式",
  "รูปแบบใบบันทึกคะแนน (พิมพ์)": "记分卡样式（打印）",
  "ลง": "下移",
  "ลงลายเซ็นผู้แก้ไข": "修改人签名",
  "ลงสนามตามผังนี้": "按此图分配场地",
  "ลบแถว": "删除行",
  "ลากภาพโปสเตอร์ / โลโก้งานมาวางที่นี่": "将海报/赛事标志拖到这里",
  "ลายเซ็น": "签名",
  "ลายเซ็นเฉพาะ": "签名：",
  "ล้างการคีย์ทั้งอีเวนต์": "清除整个赛事的录入",
  "ล้างการคีย์ทั้งอีเวนต์ (เช่น ทดลองคีย์ก่อนแข่งจริง)": "清除整个赛事录入（如赛前试录）",
  "ล้างการคีย์รอบนี้ (กรณีคีย์ก่อนเวลาหรือคีย์ผิด)": "清除本轮录入（提前或误录时）",
  "ล้างคะแนน ใบแดง ลายเซ็น และ Sudden death ทุกคนในรอบที่เลือก กลับเป็นสถานะรอตี · ล้างสาย Knockout ของอีเวนต์นี้ · รายชื่อ ลำดับยิง และสนามไม่เปลี่ยน · ช่องที่เคยคีย์ถูกบันทึกในประวัติการแก้": "清除所选轮次所有人的分数、红牌、签名和突然死亡，恢复为等待 · 清除本赛事淘汰赛对阵 · 名单、顺序和场地不变 · 已录入的格子记入改分记录",
  "ล้างคะแนนทุกช่อง ใบแดง ลายเซ็น/APPROVED และ Sudden death ของรอบนี้ · ช่องที่เคยคีย์ถูกบันทึกในประวัติการแก้พร้อมเหตุผล": "清除本轮所有分数、红牌、签名/APPROVED 和突然死亡 · 已录入格子连同原因记入改分记录",
  "ล้างทั้งหมด": "全部清除",
  "วันที่บนเอกสาร": "文件日期",
  "วิธีคัดผู้มีสิทธิ์ตีรอบ 2": "第2轮资格方式",
  "วิธีคัดรอบ 2": "第2轮选拔方式",
  "วิธีคัดเข้ารอบ 2": "第2轮选拔方式",
  "ศูนย์จับสลาก": "抽签中心",
  "สถานที่": "地点",
  "สถานที่ / Location": "地点",
  "สถานี": "站点",
  "สถิติประกอบการจัดอันดับ 5/3": "排名参考统计 5/3",
  "สนาม เริ่มที่สนาม": "个场地，起始场地",
  "เริ่มที่สนาม": "起始场地",
  "สนาม:": "场地：",
  "สนามแรกของอีเวนต์": "赛事起始场地",
  "สร้างธีม": "创建主题",
  "สร้างธีมใหม่": "新建主题",
  "สร้างสายคิวใหม่": "新建场地队列链",
  "สร้างสีรองจากสีหลักใหม่": "根据主色重新生成辅助色",
  "สร้างหลายอีเวนต์": "批量创建赛事",
  "สร้างหลายอีเวนต์ / นำเข้าจากไฟล์": "批量创建赛事 / 从文件导入",
  "สร้างหลายอีเวนต์พร้อมกัน": "同时创建多个赛事",
  "สร้างอีเวนต์ทั้งหมด": "创建全部赛事",
  "สร้างแถวอัตโนมัติ": "自动生成行",
  "สะสมเดิม:": "此前累计：",
  "สิทธิ์รอบ 2": "第2轮资格",
  "สีตัวอักษร": "文字颜色",
  "สีพื้นหน้า": "页面背景色",
  "สีพื้นอ่อน": "浅色背景",
  "สีหลัก (ปุ่ม/แบนเนอร์)": "主色（按钮/横幅）",
  "สีหลักเข้ม (หัวข้อ)": "深主色（标题）",
  "สีเน้น": "强调色",
  "สีเน้น (เส้นตัด)": "强调色（分界线）",
  "สีเส้นกรอบ": "边框颜色",
  "หน้านี้อัปเดตเองทุก 15 วินาที": "本页每 15 秒自动刷新",
  "หน้ารวม: เรียงแถวตามอันดับสด": "总览：按实时排名排序",
  "หน้ารวมทุกสนาม": "所有场地总览",
  "หมายเหตุ (ไม่บังคับ)": "备注（可选）",
  "หรือคลิกเพื่อเลือกไฟล์ · PNG, JPG, WEBP · ภาพนี้จะถูกใช้เป็นพื้นหลังแบนเนอร์ด้วย": "或点击选择文件 · PNG, JPG, WEBP · 此图也将用作横幅背景",
  "หลายรายการแล้วกดสร้างครั้งเดียว · เมื่อสร้างเสร็จจะไปหน้า": "多行后一次创建 · 完成后前往",
  "หัวข้อ": "标题",
  "หัวคอลัมน์สังกัด": "单位列标题",
  "หัวภาพผลคะแนน · ชื่องาน": "成绩图页眉 · 赛事名称",
  "หัวภาพผลคะแนน · บรรทัดเล็ก": "成绩图页眉 · 小字",
  "หัวภาพผลคะแนน · สถานที่/หน่วยงาน": "成绩图页眉 · 地点/单位",
  "หากต้องแก้คะแนน ระบบจะเก็บค่าเดิม/ค่าใหม่ ผู้แก้ เวลา และต้องมีลายเซ็นผู้แก้ไข": "改分会保存原值/新值、修改人和时间，并需修改人签名",
  "หากแก้จำนวนสนามในอีเวนต์ ระบบจะเพิ่มหรือลด Court ID ให้ตรงกับจำนวนสนามโดยอัตโนมัติ": "修改场地数量时，系统会自动增减场地账号",
  "อนุญาตตามระบบ": "按系统允许",
  "อัปโหลดไฟล์ Excel ที่มีคอลัมน์": "上传包含以下列的 Excel 文件",
  "อัปโหลดไฟล์ png, jpg, jpeg, webp หรือ gif ได้เอง ถ้าไม่อัปโหลดระบบจะใช้โลโก้ค่าเริ่มต้นตามเอกสารตัวอย่าง": "可上传 png、jpg、jpeg、webp 或 gif；不上传则使用默认标志",
  "อีเวนต์ที่เลือก ·": "所选赛事 ·",
  "อีเวนต์ในสาย": "链内赛事",
  "อ่านไฟล์และแสดงตัวอย่าง": "读取文件并预览",
  "เกิน 10 นาที ตัดออก (DQ)": "超过 10 分钟 — 取消资格（DQ）",
  "เกิน 5 นาที −5": "超过 5 分钟 −5",
  "เข้ารอบทันทีจากรอบแรก (คน)": "第一轮直接晋级人数",
  "เฉพาะหัวข้อที่มีคำว่า": "仅包含以下文字的标题",
  "เช่น 3 สนาม เริ่มสนาม 5 = ใช้สนาม 5-7 · ลำดับตีไม่เปลี่ยน": "例：3 个场地从 5 号开始 = 使用 5–7 号 · 击球顺序不变",
  "เช่น": "例如",
  "เดิม": "原值",
  "เปิด": "开启",
  "เปิด — แถวสลับตามคะแนน": "开启 — 按得分重排",
  "เปิดรอบ 2": "开启第2轮",
  "เปิดใช้ทั้งระบบ": "全站启用",
  "เพื่อกันคะแนนทับกัน หน้านี้จึงดูได้อย่างเดียว ถ้าเครื่องเดิมเลิกใช้แล้วให้กดรับช่วงคีย์ต่อ": "为避免覆盖分数，本页只读。若原设备已不用，请接管录入。",
  "เพื่อยืนยัน": "以确认",
  "เพื่อสั่งจับพร้อมกัน": "以同时抽签",
  "เพื่อให้หลายรุ่นตีต่อกันบนสนามชุดเดียว": "让多个组别在同一组场地连续比赛",
  "เริ่มจากสนาม 1 · คะแนนรอบ 1 น้อยสุดตีก่อน (เท่ากันดูจำนวน 5 แล้วจำนวน 3 น้อยกว่าตีก่อน)": "从1号场地开始 · 第1轮得分最低者先投（平分时5分、3分次数少者先）",
  "เริ่มตีแล้ว": "已开始",
  "เริ่มตีแล้ว · ล็อก": "已开始 · 锁定",
  "เรียงตาม TOTAL → จำนวน 5 → จำนวน 3 ใช้ดูเหตุผลการจัดลำดับ ไม่รวมคะแนนรอบพิเศษในตารางหลัก": "按 TOTAL → 5分次数 → 3分次数排序，用于查看排名依据，主表不含加赛分",
  "เรียงตามชื่อ (12→18, ชาย→หญิง)": "按名称排序（12→18，男→女）",
  "เรียงตามอันดับ: เปิด": "按排名排序：开启",
  "เรียงตามอันดับสด": "按实时排名排序",
  "เลือก": "选择",
  "เลือกทั้งหมดที่แสดง": "全选显示项",
  "เลือกทีมเข้ารอบ 2": "选择第2轮队伍",
  "เลือกที่พร้อมจับ (ที่แสดง)": "选择可抽签项（显示）",
  "เลือกทุกอีเวนต์": "选择所有赛事",
  "เลือกภาพของงาน": "选择赛事图片",
  "เลือกรายการใหม่": "重新选择赛事",
  "เลือกสถานีสำหรับหน้าจอแนวตั้ง": "竖屏选择站点",
  "เลือกสนาม": "选择场地",
  "เลือกหลายอีเวนต์แล้วกดจับสลากครั้งเดียว ระบบสุ่มลำดับยิงและกระจายสนามใหม่ให้ทุกอีเวนต์พร้อมกัน · อีเวนต์ที่เริ่มกรอกคะแนนแล้วจะถูกล็อกไว้": "选择多个赛事一次抽签：系统同时为所有赛事随机击球顺序并分配场地 · 已开始录分的赛事将被锁定",
  "เลือกอีเวนต์ที่ใช้สนามชุดเดียวกัน แล้วเรียงลำดับรุ่นที่จะตีก่อน–หลัง · ลำดับยิงในแต่ละรุ่นมาจากการจับสลากของรุ่นนั้น": "选择共用场地的赛事并排列组别先后 · 各组击球顺序来自该组抽签",
  "เลือกแล้ว": "已选",
  "เลือกได้ 2 วิธี:": "两种方式：",
  "เลื่อนขึ้น": "上移",
  "เลื่อนลง": "下移",
  "เวลา": "时间",
  "เว้นว่าง = ทุกประเภท": "留空 = 所有类别",
  "เส้นตัด: เปิด": "分界线：开启",
  "เส้นแบ่งเข้ารอบ": "晋级分界线",
  "เหตุผล": "原因",
  "เหตุผล เช่น คีย์ก่อนนักกีฬาตีจริง": "原因，例如运动员实际投掷前已录入",
  "เหตุผลในการแก้ (ถ้ามี)": "修改原因（如有）",
  "แตะคะแนนของแต่ละระยะ · บันทึกอัตโนมัติทุกครั้งที่แตะ": "点击每个距离的分数 · 每次点击自动保存",
  "แถบโลโก้ผู้สนับสนุน": "赞助商标志条",
  "แยก FAMILY / GIVEN — คำสุดท้ายเป็นนามสกุล (แบบไทย)": "拆分 FAMILY / GIVEN — 末词为姓（泰式）",
  "แยก FAMILY / GIVEN — คำแรกเป็นนามสกุล (แบบสากล)": "拆分 FAMILY / GIVEN — 首词为姓（国际）",
  "และ": "和",
  "แสดง": "显示",
  "แสดงผัง": "显示场地图",
  "แสดงหน้ากรรมการ": "显示裁判页",
  "แสดงเส้นแบ่ง": "显示分界线",
  "แสดงแบนเนอร์ด้านบนของเว็บ": "在网站顶部显示横幅",
  "โลโก้งาน": "赛事标志",
  "โลโก้มุมซ้ายในหน้าผล": "成绩页左上角标志",
  "โลโก้ล่างปก 1": "封面底部标志 1",
  "โลโก้ล่างปก 2": "封面底部标志 2",
  "โลโก้ล่างปก 3": "封面底部标志 3",
  "โลโก้หลักบนปก": "封面主标志",
  "โลโก้หัวกระดาษ 1": "页眉标志 1",
  "โลโก้หัวกระดาษ 2": "页眉标志 2",
  "โลโก้หัวกระดาษ 3": "页眉标志 3",
  "โลโก้หัวกระดาษ 4": "页眉标志 4",
  "โลโก้เอกสาร": "文件标志",
  "ใช้คำนวณจำนวนสนาม (0 = ใช้ค่าคงที่)": "用于计算场地数量（0 = 固定值）",
  "ใช้ค่าเริ่มต้น": "使用默认值",
  "ใช้ได้ครั้งละ 1 ธีมทั้งเว็บ (แบนเนอร์ สีปุ่ม/เมนู และหัว–ท้ายภาพผลคะแนนที่บันทึกจากหน้ารวม) · สร้างธีมใหม่จากภาพของงานได้ ระบบดึงสีจากภาพให้อัตโนมัติ": "全站一次只启用一个主题（横幅、按钮/菜单颜色、成绩图页眉页脚）· 可从赛事图片创建新主题，系统自动取色",
  "ใบนี้กำลังถูกคีย์โดย": "该记分卡正在被录入：",
  "ใบแดง": "红牌",
  "ใบแดงครบ 2 ครั้ง: ยุติการยิงรอบนี้และคงคะแนนที่ทำได้ไว้": "红牌满 2 次：本轮停止射击并保留已得分",
  "ใส่ชื่อจากไฟล์": "从文件填入姓名",
  "ใส่บรรทัดละ 1 คน รูปแบบ: ชื่อ | สังกัด เช่น MR. SOMCHAI JAIDEE | ABSC": "每行一人：姓名 | 单位，例如 MR. SOMCHAI JAIDEE | ABSC",
  "ใหม่": "新值",
  "ให้หลายอีเวนต์ตีต่อกันบนสนามชุดเดียว เมื่อรุ่นหนึ่งเหลือคนไม่เต็มสนาม รุ่นถัดไปเข้าตีสนามที่ว่างได้ทันที · อันดับและสิทธิ์เข้ารอบยังแยกต่ออีเวนต์ · ใช้กับรอบ 1 · อีเวนต์ที่ไม่ได้เลือกยังจัดสนามแยกเหมือนเดิม": "多个赛事在同一组场地连续进行：某组人数不足以占满场地时，下一组立即使用空场地 · 排名和晋级仍按赛事分开 · 仅用于第1轮 · 未选赛事仍单独分配场地",
  "ไฟล์ค่าล้วน ไม่มีสูตร แก้ไขได้ทุกช่อง": "纯数值文件，无公式，所有单元格可编辑",
  "ไฟล์ที่มีหัวข้อแต่ละประเภท เช่น": "每个类别有标题的文件，例如",
  "ไฟล์ใบสมัคร": "报名文件",
  "ไม่บังคับ": "可选",
  "ไม่มาทำการแข่งขันตามเวลา": "未按时参赛",
  "ไม่มีบทลงโทษ": "无处罚",
  "ไม่มีรอบ 2": "无第2轮",
  "ไม่ล้างคนที่ APPROVED แล้ว": "不清除已 APPROVED 的人",
  "ไม่เลือก": "未选",
  "ไม่แสดง": "不显示",
  "← กลับ Overview": "← 返回总览",
  "← กลับ scorecard": "← 返回记分卡",
  "← กลับหน้าประกบคู่": "← 返回对阵页",
  "← ก่อนหน้า": "← 上一个",
  "↩ เลิกซ่อนที่เลือก": "↩ 取消隐藏所选",
  "⚙ ตั้งค่า Results Approved / กรรมการ": "⚙ 成绩确认 / 裁判设置",
  "⚙️ ตั้งค่าร่วมกัน": "⚙️ 共同设置",
  "⚠️ ผลรอบนี้ลงลายเซ็นยืนยันแล้ว": "⚠️ 本轮成绩已签名确认",
  "✓ ระบบลงสนามตามผังนี้อยู่แล้ว": "✓ 场地已按此图分配",
  "🎲 จับสลากพร้อมกัน": "🎲 同时抽签",
  "🎲 ศูนย์จับสลาก": "🎲 抽签中心",
  "👀 เปิดดูรวมกัน": "👀 合并查看",
  "👁 แสดงเฉพาะที่เลือก": "👁 仅显示所选",
  "📄 Excel ไม่มีสูตร": "📄 Excel（无公式）",
  "📄 QF / Final ไม่มีสูตร": "📄 QF / 决赛（无公式）",
  "📊 ดูสถิติ 5/3": "📊 查看 5/3 统计",
  "🔗 ต่อคิวสนาม": "🔗 场地队列链",
  "🔗 ต่อคิวสนามข้ามรุ่น": "🔗 跨组场地队列",
  "🖨 พิมพ์ Score Sheet รอบนี้": "🖨 打印本轮记分卡",
  "🗂 ผังเข้ารอบ 2": "🗂 第2轮场地图",
  "🙈 ซ่อนที่เลือก": "🙈 隐藏所选",
  "ซ่อนอยู่": "已隐藏",
  "ซ่อนไว้": "已隐藏",
  "แสดงที่ซ่อนไว้": "显示已隐藏",
  "ซ่อนจากรายการ": "从列表隐藏"
 }
};
  Object.keys(EXTRA_TRANSLATIONS).forEach(function (lang) {
    const dict = translations[lang] || (translations[lang] = {});
    Object.keys(EXTRA_TRANSLATIONS[lang]).forEach(function (k) { if (!Object.prototype.hasOwnProperty.call(dict, k)) dict[k] = EXTRA_TRANSLATIONS[lang][k]; });
  });
  /* ข้อความที่มีตัวเลข/คำสั้นปนข้อมูล เช่น "สนาม 3", "12 คน", "Open · ชาย" (ไม่แตะคำที่ติดตัวอักษรไทยอื่น) */
  const PATTERN_RULES_EARLY = {
 "en": [
  [
   "(^|[^\\u0E00-\\u0E7F])ลบผู้ใช้ (.+?) แล้ว",
   "$1User $2 deleted"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])แก้ไขผู้ใช้ (.+?) แล้ว",
   "$1User $2 updated"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])ชื่อผู้ใช้ (.+?) มีอยู่แล้ว",
   "$1Username $2 already exists"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])ลบผู้ใช้ (.+?) \\?",
   "$1Delete user $2?"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])ซ่อนแล้ว \\((\\d+) อีเวนต์\\)",
   "$1Hidden ($3 events)"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])แสดงเฉพาะที่เลือกแล้ว \\((\\d+) อีเวนต์\\)",
   "$1Showing only selected ($3 events)"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])แสดงแล้ว \\((\\d+) อีเวนต์\\)",
   "$1Shown ($3 events)"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])Approve ผลของ (.+?) แล้ว",
   "$1$2 approved"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])กรุณาระบุเลขสนาม (\\d+)-(\\d+) และรหัสผ่าน",
   "$1Enter a court number $2-$3 and a password"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])คัดลอกเป็น “(.+?)” แล้ว",
   "$1Copied as “$2”"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])จับสลากพร้อมกัน (\\d+) อีเวนต์ เวลา (.+)",
   "$1Drew $2 events together at $3"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])ตั้งรหัสผ่านใหม่ให้ (.+?) แล้ว",
   "$1New password set for $2"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])ตั้งรหัสให้ (court\\d+) แล้ว",
   "$1Password set for $2"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])บัญชี (.+?) ใช้งานได้เฉพาะอีเวนต์ที่ได้รับมอบหมายและสนาม (\\d+)",
   "$1Account $2 can only use its assigned event and court $3"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])บันทึกการตั้งค่าร่วมให้ (\\d+) อีเวนต์แล้ว",
   "$1Shared settings saved for $2 events"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])บันทึกสาย “(.+?)” แล้ว · (\\d+) อีเวนต์ · (\\d+) คน · (\\d+) รอบคิว",
   "$1Chain “$2” saved · $3 events · $4 athletes · $5 queue rounds"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])ยกเลิกสาย “(.+?)” แล้ว · แต่ละอีเวนต์กลับไปจัดสนามแยกเหมือนเดิม",
   "$1Chain “$2” removed · each event has its own courts again"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])ลงสนามรอบ 2 ตามผังแล้ว (\\d+) อีเวนต์ · กรรมการสนามเห็นคิวใหม่ทันที",
   "$1Round 2 plan applied to $2 events · court referees see the new queue now"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])ลบ (.+?) ของอีเวนต์นี้แล้ว",
   "$1Deleted $2 from this event"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])ลบธีม “(.+?)” แล้ว",
   "$1Theme “$2” deleted"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])ล้างการคีย์รอบ (\\d+) ของ (.+?) แล้ว \\((\\d+) ช่อง\\) · กลับเป็นสถานะรอตี",
   "$1Cleared round $2 entries of $3 ($4 cells) · back to waiting"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])สร้างอีเวนต์สำเร็จ · เตรียม Court ID (\\d+) สนามแล้ว",
   "$1Event created · $2 court IDs prepared"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])สร้างได้ครั้งละไม่เกิน (\\d+) อีเวนต์",
   "$1You can create at most $2 events at a time"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])เปลี่ยนธีมทั้งระบบเป็น “(.+?)” แล้ว",
   "$1Site theme changed to “$2”"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])ใส่ชื่อแล้ว (\\d+) คน · เหมือนเดิม (\\d+) · ไม่พบ id (.+)",
   "$1Names filled for $2 athletes · $3 unchanged · id not found: $4"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])\\(ใช้กับอีกอีเวนต์ในงานเดียวกัน (\\d+) รายการ\\)",
   "$1(also applied to $2 other events of the same games)"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])แสดงเฉพาะ (\\d+) อีเวนต์ที่เลือก และซ่อนอีเวนต์อื่นทั้งหมด\\?",
   "$1Show only the $2 selected events and hide all others?"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])ล้างการคีย์รอบ (\\d+)",
   "$1Clear round $2 entries"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])รุ่นอายุ (\\d+) ปี(?=$|[^\\u0E00-\\u0E7F])",
   "$1Age group $2"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])เลือกแล้ว (\\d+) ทีม(?=$|[^\\u0E00-\\u0E7F])",
   "$1$2 teams selected"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])เลือก (\\d+) อีเวนต์(?=$|[^\\u0E00-\\u0E7F])",
   "$1$2 events selected"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])ผ่านตรง (\\d+) คน → Class ([\\d–-]+) ได้ตีรอบ 2 \\((\\d+) คน\\)",
   "$1$2 direct → Class $3 play round 2 ($4 athletes)"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])ถ้ามี Class (\\d+) หลายคน ได้สิทธิ์ทุกคน",
   "$1if several share Class $2, all qualify"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])ลำดับที่ (\\d+)",
   "$1No. $2"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])ตัวอักษร ([\\d.]+:1)",
   "$1Text $2"
  ]
 ],
 "fr": [
  [
   "(^|[^\\u0E00-\\u0E7F])ลบผู้ใช้ (.+?) แล้ว",
   "$1Utilisateur $2 supprimé"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])แก้ไขผู้ใช้ (.+?) แล้ว",
   "$1Utilisateur $2 modifié"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])ชื่อผู้ใช้ (.+?) มีอยู่แล้ว",
   "$1Le nom $2 existe déjà"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])ลบผู้ใช้ (.+?) \\?",
   "$1Supprimer l’utilisateur $2 ?"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])นำเข้านักกีฬาสำเร็จ (\\d+) คน",
   "$1$2 athlètes importés"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])ลบรายการ (.+?) เรียบร้อย",
   "$1$2 supprimé"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])สร้างอีเวนต์สำเร็จ (\\d+) รายการ",
   "$1$2 événements créés"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])ไฟล์โลโก้ (.+?) ต้องเป็น png, jpg, jpeg, webp หรือ gif",
   "$1Le logo $2 doit être png, jpg, jpeg, webp ou gif"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])ซ่อนแล้ว \\((\\d+) อีเวนต์\\)",
   "$1Masqués ($3 événements)"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])แสดงเฉพาะที่เลือกแล้ว \\((\\d+) อีเวนต์\\)",
   "$1Seule la sélection est affichée ($3 événements)"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])แสดงแล้ว \\((\\d+) อีเวนต์\\)",
   "$1Affichés ($3 événements)"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])Approve ผลของ (.+?) แล้ว",
   "$1Résultat de $2 validé"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])กรุณาระบุเลขสนาม (\\d+)-(\\d+) และรหัสผ่าน",
   "$1Indiquez un terrain $2-$3 et un mot de passe"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])คัดลอกเป็น “(.+?)” แล้ว",
   "$1Copié sous « $2 »"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])จับสลากพร้อมกัน (\\d+) อีเวนต์ เวลา (.+)",
   "$1$2 événements tirés ensemble à $3"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])ตั้งรหัสผ่านใหม่ให้ (.+?) แล้ว",
   "$1Nouveau mot de passe défini pour $2"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])ตั้งรหัสให้ (court\\d+) แล้ว",
   "$1Mot de passe défini pour $2"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])บัญชี (.+?) ใช้งานได้เฉพาะอีเวนต์ที่ได้รับมอบหมายและสนาม (\\d+)",
   "$1Le compte $2 n’a accès qu’à son événement et au terrain $3"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])บันทึกการตั้งค่าร่วมให้ (\\d+) อีเวนต์แล้ว",
   "$1Paramètres communs enregistrés pour $2 événements"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])บันทึกสาย “(.+?)” แล้ว · (\\d+) อีเวนต์ · (\\d+) คน · (\\d+) รอบคิว",
   "$1Chaîne « $2 » enregistrée · $3 événements · $4 athlètes · $5 rotations"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])ยกเลิกสาย “(.+?)” แล้ว · แต่ละอีเวนต์กลับไปจัดสนามแยกเหมือนเดิม",
   "$1Chaîne « $2 » supprimée · chaque événement retrouve ses terrains"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])ลงสนามรอบ 2 ตามผังแล้ว (\\d+) อีเวนต์ · กรรมการสนามเห็นคิวใหม่ทันที",
   "$1Plan du tour 2 appliqué à $2 événements · les arbitres voient la nouvelle file"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])ลบ (.+?) ของอีเวนต์นี้แล้ว",
   "$1$2 supprimé de cet événement"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])ลบธีม “(.+?)” แล้ว",
   "$1Thème « $2 » supprimé"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])ล้างการคีย์รอบ (\\d+) ของ (.+?) แล้ว \\((\\d+) ช่อง\\) · กลับเป็นสถานะรอตี",
   "$1Saisies du tour $2 de $3 effacées ($4 cases) · retour en attente"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])สร้างอีเวนต์สำเร็จ · เตรียม Court ID (\\d+) สนามแล้ว",
   "$1Événement créé · $2 identifiants terrain préparés"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])สร้างได้ครั้งละไม่เกิน (\\d+) อีเวนต์",
   "$1$2 événements maximum à la fois"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])เปลี่ยนธีมทั้งระบบเป็น “(.+?)” แล้ว",
   "$1Thème du site changé en « $2 »"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])ใส่ชื่อแล้ว (\\d+) คน · เหมือนเดิม (\\d+) · ไม่พบ id (.+)",
   "$1Noms remplis pour $2 athlètes · $3 inchangés · id introuvable : $4"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])\\(ใช้กับอีกอีเวนต์ในงานเดียวกัน (\\d+) รายการ\\)",
   "$1(appliqué aussi à $2 autres événements de la même compétition)"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])แสดงเฉพาะ (\\d+) อีเวนต์ที่เลือก และซ่อนอีเวนต์อื่นทั้งหมด\\?",
   "$1Afficher seulement les $2 événements choisis et masquer les autres ?"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])ล้างการคีย์รอบ (\\d+)",
   "$1Effacer les saisies du tour $2"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])รุ่นอายุ (\\d+) ปี(?=$|[^\\u0E00-\\u0E7F])",
   "$1Catégorie $2 ans"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])เลือกแล้ว (\\d+) ทีม(?=$|[^\\u0E00-\\u0E7F])",
   "$1$2 équipes choisies"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])เลือก (\\d+) อีเวนต์(?=$|[^\\u0E00-\\u0E7F])",
   "$1$2 événements choisis"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])ผ่านตรง (\\d+) คน → Class ([\\d–-]+) ได้ตีรอบ 2 \\((\\d+) คน\\)",
   "$1$2 directs → Class $3 jouent le tour 2 ($4 athlètes)"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])ถ้ามี Class (\\d+) หลายคน ได้สิทธิ์ทุกคน",
   "$1en cas d’égalité à la Class $2, tous qualifiés"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])รอบที่กำลังดู:\\s*(\\d+)",
   "$1Tour affiché : $2"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])สถานีที่? (\\d+)",
   "$1Atelier $2"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])ลำดับที่ (\\d+)",
   "$1N° $2"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])ตัวอักษร ([\\d.]+:1)",
   "$1Texte $2"
  ]
 ],
 "zh": [
  [
   "(^|[^\\u0E00-\\u0E7F])ลบผู้ใช้ (.+?) แล้ว",
   "$1已删除用户 $2"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])แก้ไขผู้ใช้ (.+?) แล้ว",
   "$1已修改用户 $2"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])ชื่อผู้ใช้ (.+?) มีอยู่แล้ว",
   "$1用户名 $2 已存在"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])ลบผู้ใช้ (.+?) \\?",
   "$1删除用户 $2？"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])นำเข้านักกีฬาสำเร็จ (\\d+) คน",
   "$1已导入 $2 名运动员"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])ลบรายการ (.+?) เรียบร้อย",
   "$1已删除 $2"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])สร้างอีเวนต์สำเร็จ (\\d+) รายการ",
   "$1已创建 $2 个赛事"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])ไฟล์โลโก้ (.+?) ต้องเป็น png, jpg, jpeg, webp หรือ gif",
   "$1标志文件 $2 必须为 png、jpg、jpeg、webp 或 gif"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])ซ่อนแล้ว \\((\\d+) อีเวนต์\\)",
   "$1已隐藏 （$3 个赛事）"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])แสดงเฉพาะที่เลือกแล้ว \\((\\d+) อีเวนต์\\)",
   "$1仅显示所选 （$3 个赛事）"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])แสดงแล้ว \\((\\d+) อีเวนต์\\)",
   "$1已显示 （$3 个赛事）"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])Approve ผลของ (.+?) แล้ว",
   "$1已确认 $2 的成绩"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])กรุณาระบุเลขสนาม (\\d+)-(\\d+) และรหัสผ่าน",
   "$1请填写场地号 $2-$3 和密码"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])คัดลอกเป็น “(.+?)” แล้ว",
   "$1已复制为“$2”"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])จับสลากพร้อมกัน (\\d+) อีเวนต์ เวลา (.+)",
   "$1已同时抽签 $2 个赛事，时间 $3"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])ตั้งรหัสผ่านใหม่ให้ (.+?) แล้ว",
   "$1已为 $2 设置新密码"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])ตั้งรหัสให้ (court\\d+) แล้ว",
   "$1已为 $2 设置密码"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])บัญชี (.+?) ใช้งานได้เฉพาะอีเวนต์ที่ได้รับมอบหมายและสนาม (\\d+)",
   "$1账号 $2 只能使用指定赛事和场地 $3"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])บันทึกการตั้งค่าร่วมให้ (\\d+) อีเวนต์แล้ว",
   "$1已为 $2 个赛事保存共同设置"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])บันทึกสาย “(.+?)” แล้ว · (\\d+) อีเวนต์ · (\\d+) คน · (\\d+) รอบคิว",
   "$1已保存链“$2” · $3 个赛事 · $4 人 · $5 轮队列"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])ยกเลิกสาย “(.+?)” แล้ว · แต่ละอีเวนต์กลับไปจัดสนามแยกเหมือนเดิม",
   "$1已取消链“$2” · 各赛事恢复单独分配场地"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])ลงสนามรอบ 2 ตามผังแล้ว (\\d+) อีเวนต์ · กรรมการสนามเห็นคิวใหม่ทันที",
   "$1已为 $2 个赛事按图分配第2轮场地 · 场地裁判立即看到新队列"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])ลบ (.+?) ของอีเวนต์นี้แล้ว",
   "$1已从本赛事删除 $2"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])ลบธีม “(.+?)” แล้ว",
   "$1已删除主题“$2”"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])ล้างการคีย์รอบ (\\d+) ของ (.+?) แล้ว \\((\\d+) ช่อง\\) · กลับเป็นสถานะรอตี",
   "$1已清除 $3 第$2轮录入（$4 格）· 恢复为等待"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])สร้างอีเวนต์สำเร็จ · เตรียม Court ID (\\d+) สนามแล้ว",
   "$1赛事已创建 · 已准备 $2 个场地账号"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])สร้างได้ครั้งละไม่เกิน (\\d+) อีเวนต์",
   "$1每次最多创建 $2 个赛事"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])เปลี่ยนธีมทั้งระบบเป็น “(.+?)” แล้ว",
   "$1全站主题已改为“$2”"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])ใส่ชื่อแล้ว (\\d+) คน · เหมือนเดิม (\\d+) · ไม่พบ id (.+)",
   "$1已填入 $2 人姓名 · $3 人未变 · 未找到 id：$4"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])\\(ใช้กับอีกอีเวนต์ในงานเดียวกัน (\\d+) รายการ\\)",
   "$1（同时应用于同一赛会的另外 $2 个赛事）"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])แสดงเฉพาะ (\\d+) อีเวนต์ที่เลือก และซ่อนอีเวนต์อื่นทั้งหมด\\?",
   "$1仅显示所选 $2 个赛事并隐藏其他所有赛事？"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])ล้างการคีย์รอบ (\\d+)",
   "$1清除第$2轮录入"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])รุ่นอายุ (\\d+) ปี(?=$|[^\\u0E00-\\u0E7F])",
   "$1$2岁组"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])เลือกแล้ว (\\d+) ทีม(?=$|[^\\u0E00-\\u0E7F])",
   "$1已选 $2 队"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])เลือก (\\d+) อีเวนต์(?=$|[^\\u0E00-\\u0E7F])",
   "$1已选 $2 个赛事"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])ผ่านตรง (\\d+) คน → Class ([\\d–-]+) ได้ตีรอบ 2 \\((\\d+) คน\\)",
   "$1直接晋级 $2 人 → Class $3 进入第2轮（$4 人）"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])ถ้ามี Class (\\d+) หลายคน ได้สิทธิ์ทุกคน",
   "$1Class $2 并列多人时全部获得资格"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])รอบที่กำลังดู:\\s*(\\d+)",
   "$1当前轮次：$2"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])สถานีที่? (\\d+)",
   "$1站点 $2"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])ลำดับที่ (\\d+)",
   "$1第 $2 位"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])ตัวอักษร ([\\d.]+:1)",
   "$1文字 $2"
  ]
 ]
};
  const PATTERN_RULES = {
 "en": [
  [
   "^เช่น ",
   "e.g. "
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])(\\d+) ทีม(?=$|[^\\u0E00-\\u0E7F])",
   "$1$2 teams"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])(\\d+) รายการ(?=$|[^\\u0E00-\\u0E7F])",
   "$1$2 events"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])(\\d+) ปี(?=$|[^\\u0E00-\\u0E7F])",
   "$1$2 yrs"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])(\\d+) คน(?=$|[^\\u0E00-\\u0E7F])",
   "$1$2 athletes"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])รอบที่ (\\d+)",
   "$1Round $2"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])รอบ (\\d+)",
   "$1Round $2"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])สนาม (\\d+)",
   "$1Court $2"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])สถานี (\\d+)",
   "$1Station $2"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])(\\d+) ม\\.",
   "$1$2 m"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])ปุ่ม ([\\d.]+:1)",
   "$1Button $2"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])หัวข้อ ([\\d.]+:1)",
   "$1Heading $2"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])ชาย(?=$|[^\\u0E00-\\u0E7F])",
   "$1Men"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])หญิง(?=$|[^\\u0E00-\\u0E7F])",
   "$1Women"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])ผสม(?=$|[^\\u0E00-\\u0E7F])",
   "$1Mixed"
  ]
 ],
 "fr": [
  [
   "^เช่น ",
   "ex. "
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])(\\d+) ทีม(?=$|[^\\u0E00-\\u0E7F])",
   "$1$2 équipes"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])(\\d+) รายการ(?=$|[^\\u0E00-\\u0E7F])",
   "$1$2 événements"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])(\\d+) ปี(?=$|[^\\u0E00-\\u0E7F])",
   "$1$2 ans"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])จำนวน (\\d+) คน",
   "$1$2 athlètes"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])(\\d+) คน(?=$|[^\\u0E00-\\u0E7F])",
   "$1$2 athlètes"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])รอบที่ (\\d+)",
   "$1Tour $2"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])รอบ (\\d+)",
   "$1Tour $2"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])สนาม (\\d+)",
   "$1Terrain $2"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])(\\d+) ม\\.",
   "$1$2 m"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])ส(\\d+) (\\d+)/4",
   "$1Atelier $2 : $3/4"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])ปุ่ม ([\\d.]+:1)",
   "$1Bouton $2"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])หัวข้อ ([\\d.]+:1)",
   "$1Titre $2"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])ชาย(?=$|[^\\u0E00-\\u0E7F])",
   "$1Hommes"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])หญิง(?=$|[^\\u0E00-\\u0E7F])",
   "$1Femmes"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])ผสม(?=$|[^\\u0E00-\\u0E7F])",
   "$1Mixte"
  ]
 ],
 "zh": [
  [
   "^เช่น ",
   "例如 "
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])(\\d+) ทีม(?=$|[^\\u0E00-\\u0E7F])",
   "$1$2 队"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])(\\d+) รายการ(?=$|[^\\u0E00-\\u0E7F])",
   "$1$2 个赛事"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])(\\d+) ปี(?=$|[^\\u0E00-\\u0E7F])",
   "$1$2岁"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])จำนวน (\\d+) คน",
   "$1$2 人"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])(\\d+) คน(?=$|[^\\u0E00-\\u0E7F])",
   "$1$2 人"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])รอบที่ (\\d+)",
   "$1第$2轮"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])รอบ (\\d+)",
   "$1第$2轮"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])สนาม (\\d+)",
   "$1场地 $2"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])(\\d+) ม\\.",
   "$1$2 米"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])ส(\\d+) (\\d+)/4",
   "$1站点$2：$3/4"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])ปุ่ม ([\\d.]+:1)",
   "$1按钮 $2"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])หัวข้อ ([\\d.]+:1)",
   "$1标题 $2"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])ชาย(?=$|[^\\u0E00-\\u0E7F])",
   "$1男子"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])หญิง(?=$|[^\\u0E00-\\u0E7F])",
   "$1女子"
  ],
  [
   "(^|[^\\u0E00-\\u0E7F])ผสม(?=$|[^\\u0E00-\\u0E7F])",
   "$1混合"
  ]
 ]
};
  [PATTERN_RULES_EARLY, PATTERN_RULES].forEach(function (group) { Object.keys(group).forEach(function (lang) { group[lang] = group[lang].map(function (r) { return [new RegExp(r[0], 'g'), r[1]]; }); }); });

  function translateText(text, dict) {
    if (!text || !dict) return text;
    const leading = text.match(/^\s*/)[0];
    const trailing = text.match(/\s*$/)[0];
    const trimmed = text.trim();
    if (Object.prototype.hasOwnProperty.call(dict, trimmed)) {
      return leading + dict[trimmed] + trailing;
    }

    let result = text;
    if (dict === translations.en) {
      result = result
        /* Translate complete dynamic sentences before individual words. */
        .replace(/ยืนยันลบ\s+(.+?)\s*\?/g, 'Delete $1?')
        .replace(/ลบรายการ\s+(.+?)\s+เรียบร้อย/g, 'Deleted $1 successfully.')
        .replace(/นำเข้านักกีฬาสำเร็จ\s+(\d+)\s+คน/g, '$1 athletes were imported successfully.')
        .replace(/จำนวน\s+(\d+)\s+คน\s*\|/g, '$1 athletes |')
        .replace(/จำนวน\s+(\d+)\s+คน/g, '$1 athletes')
        .replace(/แถวที่\s+(\d+)\s+ต้องมีทั้งชื่อและสังกัด/g, 'Row $1 must include both a name and an affiliation.')
        .replace(/ไฟล์โลโก้\s+(.+?)\s+ต้องเป็น png, jpg, jpeg, webp หรือ gif/g, 'Logo file $1 must be a PNG, JPG, JPEG, WEBP, or GIF file.')
        .replace(/(\d+)\s*คะแนน\s*\|\s*5 คะแนน\s*(\d+)\s*ครั้ง\s*\|\s*3 คะแนน\s*(\d+)\s*ครั้ง/g, '$1 points | 5-point shots: $2 | 3-point shots: $3')
        .replace(/รอบ\s+(\d+)\s+·\s+เลือกคะแนนแบบ 0\s*\/\s*1\s*\/\s*3\s*\/\s*5/g, 'Round $1 · Select a score of 0, 1, 3, or 5')
        .replace(/กำลังตี\s+ส(\d+)/g, 'Shooting S$1')
        .replace(/(^|\|\s*)(ชาย|หญิง|ผสม)(?=\s*(?:\||$))/g, function (_, prefix, category) {
          return prefix + { 'ชาย': 'Men', 'หญิง': 'Women', 'ผสม': 'Mixed' }[category];
        })
        .replace(/ส(\d+)\s+(\d+)\/4/g, 'Station $1: $2/4 shots')
        .replace(/สถานีที่?\s+(\d+)/g, 'Station $1')
        .replace(/รอบที่กำลังดู:\s*(\d+)/g, 'Viewing Round $1');
    }

    const earlyLang = Object.keys(translations).find(function (l) { return translations[l] === dict; });
    if (earlyLang && PATTERN_RULES_EARLY[earlyLang]) {
      PATTERN_RULES_EARLY[earlyLang].forEach(function (rule) { result = result.replace(rule[0], rule[1]); });
    }

    const keys = Object.keys(dict)
      .filter(key => key.length >= 5 || /[\s()/?:|]/.test(key))
      .sort((a, b) => b.length - a.length);
    // ภาษาไทยไม่มีช่องว่างคั่นคำ ถ้าแทนคำย่อยตรง ๆ ชื่อคนจะพัง (เช่น "นายเดชสิทธิ์" -> "นายเดชRole")
    // จึงแทนเฉพาะวลีที่ขอบทั้งสองข้างไม่ติดตัวอักษรไทย
    for (const key of keys) {
      if (!result.includes(key)) continue;
      const re = new RegExp('(^|[^\u0E00-\u0E7F])' + key.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '(?=$|[^\u0E00-\u0E7F])', 'g');
      result = result.replace(re, (m, lead) => lead + dict[key]);
    }
    const langKey = Object.keys(translations).find(function (l) { return translations[l] === dict; });
    if (langKey && PATTERN_RULES[langKey] && /[\u0E00-\u0E7F]/.test(result)) {
      PATTERN_RULES[langKey].forEach(function (rule) { result = result.replace(rule[0], rule[1]); });
    }
    return result;
  }

  const TRANSLATABLE_ATTRIBUTES = ['placeholder', 'title', 'alt', 'data-title', 'aria-label'];

  function translateElementAttributes(el, dict) {
    TRANSLATABLE_ATTRIBUTES.forEach(attr => {
      if (el.hasAttribute(attr)) {
        el.setAttribute(attr, translateText(el.getAttribute(attr), dict));
      }
    });
  }

  function translateNode(root, dict) {
    if (!dict) return;
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, {
      acceptNode(node) {
        const parent = node.parentElement;
        if (!parent) return NodeFilter.FILTER_REJECT;
        const tag = parent.tagName;
        if (['SCRIPT', 'STYLE', 'TEXTAREA'].includes(tag)) return NodeFilter.FILTER_REJECT;
        // ข้อมูลจริง (ชื่อคน/สังกัด) ใส่ translate="no" เพื่อไม่ให้แปลเด็ดขาด
        if (parent.closest('[translate="no"], .notranslate')) return NodeFilter.FILTER_REJECT;
        if (!node.nodeValue.trim()) return NodeFilter.FILTER_SKIP;
        return NodeFilter.FILTER_ACCEPT;
      }
    });
    const nodes = [];
    while (walker.nextNode()) nodes.push(walker.currentNode);
    nodes.forEach(node => {
      node.nodeValue = translateText(node.nodeValue, dict);
    });

    if (root.nodeType === Node.ELEMENT_NODE) {
      translateElementAttributes(root, dict);
    }
    root.querySelectorAll('[placeholder], [title], [alt], [data-title], [aria-label]').forEach(el => {
      translateElementAttributes(el, dict);
    });

    root.querySelectorAll('[onsubmit], [onclick]').forEach(el => {
      ['onsubmit', 'onclick'].forEach(attr => {
        if (el.hasAttribute(attr)) {
          el.setAttribute(attr, translateText(el.getAttribute(attr), dict));
        }
      });
    });
  }

  function installDialogTranslation(dict) {
    if (!dict) return;
    const nativeAlert = window.alert;
    const nativeConfirm = window.confirm;
    window.alert = function (message) {
      return nativeAlert.call(window, translateText(String(message), dict));
    };
    window.confirm = function (message) {
      return nativeConfirm.call(window, translateText(String(message), dict));
    };
  }

  function installDynamicTranslation(dict) {
    if (!dict || !document.body) return;

    let observer;
    const startObserving = function () {
      observer.observe(document.body, {
        childList: true,
        subtree: true,
        characterData: true,
        attributes: true,
        attributeFilter: TRANSLATABLE_ATTRIBUTES
      });
    };

    observer = new MutationObserver(function (mutations) {
      observer.disconnect();

      mutations.forEach(mutation => {
        if (mutation.type === 'characterData') {
          mutation.target.nodeValue = translateText(mutation.target.nodeValue, dict);
          return;
        }

        if (mutation.type === 'attributes') {
          translateElementAttributes(mutation.target, dict);
          return;
        }

        mutation.addedNodes.forEach(node => {
          if (node.nodeType === Node.TEXT_NODE) {
            node.nodeValue = translateText(node.nodeValue, dict);
          } else if (node.nodeType === Node.ELEMENT_NODE) {
            translateNode(node, dict);
          }
        });
      });

      startObserving();
    });

    startObserving();
  }

  document.addEventListener('DOMContentLoaded', function () {
    const switcher = document.getElementById('languageSwitcher');
    if (switcher) {
      switcher.addEventListener('change', function () {
        localStorage.setItem('app_lang', this.value);
        const template = window.SET_LANGUAGE_URL || '/set-language/__LANG__';
        window.location.href = template.replace('__LANG__', this.value);
      });
    }

    localStorage.setItem('app_lang', LANG);
    if (LANG !== 'th') {
      const dict = translations[LANG];
      if (LANG === 'en') {
        document.documentElement.style.setProperty('--i18n-cut-line', '"QUALIFICATION CUT LINE"');
      }
      translateNode(document.body, dict);
      document.title = translateText(document.title, dict);
      installDialogTranslation(dict);
      installDynamicTranslation(dict);
    }
  });

  window.translateI18nText = function (text) {
    return LANG === 'th' ? text : translateText(text, translations[LANG]);
  };
})();
