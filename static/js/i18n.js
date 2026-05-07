(function () {
  const LANG = window.APP_LANG || localStorage.getItem('app_lang') || 'th';

  const translations = {
    en: {
      'หน้าหลัก': 'Home',
      'สร้างอีเวนต์': 'Create event',
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
      'รอบ 8 คน': 'Quaterfinal',
      'รอบ 4 คน': 'Semi final',
      'รอบรองชนะเลิศ': 'Semi Final',
      'รอบชิงชนะเลิศ': 'Final',
      'ตารางประกบ': 'Bracket',
      'หน้าประกบคู่': 'Bracket page',
      'พิมพ์ Scorecard หลายคน': 'Print multiple scorecards',
      'บันทึกเป็นรูปภาพ': 'Save as image',
      'พิมพ์ตาราง': 'Print table',
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
      'สังกัด': 'Contry',
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
      'ลายเซ็นนักกีฬา': 'signature',
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
      'ลบอีเวนต์นี้?': 'Delete this event?'
    },
    fr: {
      'หน้าหลัก': 'Accueil', 'สร้างอีเวนต์': 'Créer un événement', 'สร้างอีเวนต์ Shooting': 'Créer un événement de tir', 'แก้ไขอีเวนต์': 'Modifier l’événement', 'ผู้ใช้': 'Utilisateurs', 'ออกจากระบบ': 'Déconnexion', 'เข้าสู่ระบบ': 'Connexion', 'ชื่อผู้ใช้': 'Nom d’utilisateur', 'รหัสผ่าน': 'Mot de passe', 'สิทธิ์': 'Rôle', 'สร้างผู้ใช้ใหม่': 'Créer un utilisateur', 'รายการผู้ใช้': 'Liste des utilisateurs', 'บันทึก': 'Enregistrer', 'บันทึกการแก้ไข': 'Enregistrer les modifications',
      'แดชบอร์ด': 'Tableau de bord', 'ภาพรวมของรายการแข่งขันและนักกีฬาในระบบ': 'Vue d’ensemble des compétitions et des athlètes', 'อีเวนต์': 'Événements', 'รายการแข่งขัน': 'Compétitions', 'เลือกดูหน้ารวม สถิติ และจัดการนักกีฬาได้จากตารางนี้': 'Ouvrir l’aperçu, les statistiques et la gestion des athlètes depuis ce tableau', 'คะแนนสูงสุดในระบบ': 'Meilleur score du système', 'สถิติสังกัดที่ทำคะแนนดีที่สุด': 'Statistiques des clubs/affiliations les mieux notés', 'คะแนนสูงสุด': 'Meilleur score', 'คะแนนรวม': 'Score total', 'คะแนน': 'points',
      'ชื่องาน': 'Nom de l’événement', 'รุ่น': 'Groupe', 'รุ่นแข่งขัน': 'Groupe de compétition', 'ประเภท': 'Catégorie', 'วันแข่งขัน': 'Date de compétition', 'สนาม': 'Terrain', 'จำนวนสนาม': 'Nombre de terrains', 'จัดการ': 'Gérer', 'หน้ารวม': 'Aperçu', 'สถิติ': 'Statistiques', 'นักกีฬา': 'Athlètes', 'แก้ไข': 'Modifier', 'ลบ': 'Supprimer', 'ยังไม่มีอีเวนต์': 'Aucun événement', 'ยังไม่มีข้อมูล': 'Aucune donnée',
      'ทั่วไป': 'Open', 'อาวุโส': 'Senior', 'เยาวชน': 'Jeunes', 'ชาย': 'Hommes', 'หญิง': 'Femmes', 'ผสม': 'Mixte', 'มี': 'Oui', 'ไม่มี': 'Non', 'จากรอบแรกเข้ารอบถัดไปทันที กี่คน': 'Qualifiés directs depuis le 1er tour', 'มีรอบ 2 หรือไม่': 'Activer le 2e tour ?', 'ถ้ามีรอบ 2 ให้ตีถึงลำดับที่เท่าไหร่': 'Rang limite du 2e tour', 'จากรอบ 2 คัดเข้ารอบต่อกี่คน': 'Qualifiés depuis le 2e tour', 'รอบถัดไปคือรอบอะไร': 'Nom du tour suivant',
      'รอบที่กำลังดู': 'Tour affiché', 'รอบ 1': 'Tour 1', 'รอบ 2': 'Tour 2', 'รอบที่ 1': 'Tour 1', 'รอบที่ 2': 'Tour 2', 'รอบ 8 คน': 'Quart de finale', 'รอบ 4 คน': 'Demi-finale à 4', 'รอบรองชนะเลิศ': 'Demi-finale', 'รอบชิงชนะเลิศ': 'Finale', 'ตารางประกบ': 'Tableau', 'หน้าประกบคู่': 'Page du tableau', 'พิมพ์ Scorecard หลายคน': 'Imprimer plusieurs fiches', 'บันทึกเป็นรูปภาพ': 'Enregistrer comme image', 'พิมพ์ตาราง': 'Imprimer le tableau',
      'รอคิว': 'En attente', 'กำลังตี': 'En cours', 'ตีเสร็จแล้ว': 'Terminé', 'เข้ารอบตรง': 'Qualifié direct', 'ผ่านจากรอบ 2': 'Qualifié du 2e tour', 'มีสิทธิ์รอบ 2': 'Éligible au 2e tour', 'ตกรอบ': 'Éliminé', 'ลำดับการตี': 'Ordre de tir', 'ลำดับในสนาม': 'Ordre sur terrain', 'สถานะ': 'Statut', 'หมายเลข': 'Dossard', 'ชื่อ': 'Nom', 'สังกัด': 'Affiliation',
      'เพิ่มนักกีฬาและจัดลำดับยิง': 'Ajouter des athlètes et définir l’ordre de tir', 'ไปหน้ารวม': 'Aller à l’aperçu', 'หมายเลขและลำดับรวม ระบบสร้างให้อัตโนมัติ': 'Le dossard et l’ordre général sont générés automatiquement', 'เพิ่มนักกีฬา': 'Ajouter un athlète', 'การกระจายสนามอัตโนมัติ': 'Répartition automatique des terrains', 'สนาม = วนตามจำนวนสนาม เช่น 1, 2, 3, 4 แล้วกลับไป 1': 'Les terrains tournent selon le nombre disponible, ex. 1, 2, 3, 4 puis retour à 1', 'สุ่มลำดับทั้งหมด': 'Mélanger tous les ordres', 'นำเข้ารายชื่อนักกีฬาจาก Excel': 'Importer une liste depuis Excel', 'ไฟล์ต้องเป็น .xlsx และต้องมี 2 คอลัมน์คือ': 'Le fichier doit être .xlsx avec 2 colonnes :', 'เลือกไฟล์ Excel': 'Choisir un fichier Excel', 'อัปโหลด Excel': 'Importer Excel', 'ดาวน์โหลดไฟล์ตัวอย่าง': 'Télécharger le modèle', 'ยังไม่มีนักกีฬา': 'Aucun athlète',
      'ประเภทสุดยอดความแม่นยำ (SHOOTING)': 'Tir de précision', 'กลับหน้ารวม': 'Retour à l’aperçu', 'พิมพ์แบบฟอร์ม': 'Imprimer le formulaire', 'รอบที่คีย์': 'Tour saisi', 'สถานีที่': 'Atelier', 'ยิงลูกเดี่ยว': 'Tir boule seule', 'ยิงลูกเหนือเป้า': 'Tir par-dessus la cible', 'ยิงลูกกลางขวาง': 'Tir avec obstacle central', 'ยิงลูกขาวเหนือลูกดำ': 'Boule blanche par-dessus boule noire', 'ยิงลูกเป้า': 'Tir sur la cible', 'รอบแข่งขัน': 'Tour de compétition', 'รวม': 'Total', 'แดง': 'Rouge',
      'กรรมการตัดสิน': 'Arbitre', 'กรรมการบันทึกคะแนน': 'Marqueur', 'ลายเซ็นกรรมการตัดสิน': 'Signature de l’arbitre', 'ลายเซ็นกรรมการบันทึกคะแนน': 'Signature du marqueur', 'ลายเซ็นนักกีฬา': 'Signature de l’athlète', 'พิมพ์ชื่อ ถ้าไม่เซ็นระบบจะใช้ชื่อนี้': 'Saisir le nom ; utilisé si aucune signature n’est dessinée', 'เซ็นชื่อเต็มหน้าจอ': 'Signature plein écran', 'ล้างลายเซ็น': 'Effacer la signature', 'แตะปุ่มเพื่อเปิดพื้นที่เซ็นเต็มหน้าจอ': 'Touchez le bouton pour ouvrir la zone de signature plein écran', 'รหัสข้ามการลงชื่อ (admin)': 'Code de contournement signature (admin)', 'กรอกเมื่อจำเป็น': 'Saisir si nécessaire', 'จบการตี': 'Terminer le tir', 'ลงลายเซ็น': 'Signer', 'เซ็นเต็มหน้าจอ แล้วกดบันทึกเพื่อกลับไปยังฟอร์ม': 'Signer en plein écran puis enregistrer pour revenir au formulaire', 'ปิด': 'Fermer', 'ล้าง': 'Effacer', 'พิมพ์': 'Imprimer',
      'ลำดับ': 'Rang', 'ลายเซ็น / ผู้เกี่ยวข้องแต่ละรอบ': 'Signatures / officiels par tour', 'รอบ': 'Tour', 'กรรมการ': 'Arbitre', 'ผู้บันทึก': 'Marqueur', 'กลับหน้าเลือกนักกีฬา': 'Retour à la sélection', 'พิมพ์ทั้งหมดที่แสดง': 'Imprimer tout ce qui est affiché', 'จำนวน': 'Nombre', 'คน': 'personnes', 'ยังไม่มีนักกีฬาสำหรับพิมพ์ในรอบนี้': 'Aucun athlète à imprimer pour ce tour', 'เลือกพิมพ์ Scorecard': 'Choisir les fiches à imprimer', 'เลือกรอบ': 'Choisir le tour', 'เปลี่ยนรอบ': 'Changer de tour', 'เลือกทั้งหมด': 'Tout sélectionner', 'ยกเลิกทั้งหมด': 'Tout désélectionner', 'ยังไม่มีนักกีฬาในรอบนี้': 'Aucun athlète dans ce tour', 'พิมพ์คนที่เลือก': 'Imprimer la sélection', 'พิมพ์ทั้งหมดในรอบนี้': 'Imprimer tout ce tour',
      'คะแนนสูงสุดของประเภทนี้': 'Meilleur score de cette catégorie', 'อันดับ': 'Classement', 'ครั้ง': 'fois', 'ตีเที่ยวพิเศษ 7 เมตรทุกสถานี': 'Départage spécial : 7 m à chaque atelier', 'เที่ยวพิเศษ': 'Départage', 'บันทึกเที่ยวพิเศษ': 'Enregistrer le départage', 'กดที่ชื่อในแต่ละ match box เพื่อเปิด scorecard ของรอบนั้น': 'Cliquez sur un nom dans chaque case pour ouvrir la fiche du tour', 'ยังไม่มีคู่ Quarter Final': 'Aucun quart de finale', 'ยังไม่มีคู่ Semi Final': 'Aucune demi-finale', 'ยังไม่มีคู่ Final': 'Aucune finale', 'ยืนยันลบ': 'Confirmer la suppression', 'ลบอีเวนต์นี้?': 'Supprimer cet événement ?'
    },
    zh: {
      'หน้าหลัก': '首页', 'สร้างอีเวนต์': '创建赛事', 'สร้างอีเวนต์ Shooting': '创建射击赛事', 'แก้ไขอีเวนต์': '编辑赛事', 'ผู้ใช้': '用户', 'ออกจากระบบ': '退出登录', 'เข้าสู่ระบบ': '登录', 'ชื่อผู้ใช้': '用户名', 'รหัสผ่าน': '密码', 'สิทธิ์': '权限', 'สร้างผู้ใช้ใหม่': '创建新用户', 'รายการผู้ใช้': '用户列表', 'บันทึก': '保存', 'บันทึกการแก้ไข': '保存修改',
      'แดชบอร์ด': '仪表盘', 'ภาพรวมของรายการแข่งขันและนักกีฬาในระบบ': '系统中的赛事和运动员概览', 'อีเวนต์': '赛事', 'รายการแข่งขัน': '比赛列表', 'เลือกดูหน้ารวม สถิติ และจัดการนักกีฬาได้จากตารางนี้': '可从此表进入总览、统计和运动员管理', 'คะแนนสูงสุดในระบบ': '系统最高分', 'สถิติสังกัดที่ทำคะแนนดีที่สุด': '最佳单位/队伍统计', 'คะแนนสูงสุด': '最高分', 'คะแนนรวม': '总分', 'คะแนน': '分',
      'ชื่องาน': '赛事名称', 'รุ่น': '组别', 'รุ่นแข่งขัน': '比赛组别', 'ประเภท': '类别', 'วันแข่งขัน': '比赛日期', 'สนาม': '场地', 'จำนวนสนาม': '场地数量', 'จัดการ': '管理', 'หน้ารวม': '总览', 'สถิติ': '统计', 'นักกีฬา': '运动员', 'แก้ไข': '编辑', 'ลบ': '删除', 'ยังไม่มีอีเวนต์': '暂无赛事', 'ยังไม่มีข้อมูล': '暂无数据',
      'ทั่วไป': '公开组', 'อาวุโส': '长青组', 'เยาวชน': '青年组', 'ชาย': '男子', 'หญิง': '女子', 'ผสม': '混合', 'มี': '有', 'ไม่มี': '无', 'จากรอบแรกเข้ารอบถัดไปทันที กี่คน': '第一轮直接晋级人数', 'มีรอบ 2 หรือไม่': '是否有第二轮？', 'ถ้ามีรอบ 2 ให้ตีถึงลำดับที่เท่าไหร่': '第二轮截止名次', 'จากรอบ 2 คัดเข้ารอบต่อกี่คน': '第二轮晋级人数', 'รอบถัดไปคือรอบอะไร': '下一轮名称',
      'รอบที่กำลังดู': '当前查看轮次', 'รอบ 1': '第1轮', 'รอบ 2': '第2轮', 'รอบที่ 1': '第1轮', 'รอบที่ 2': '第2轮', 'รอบ 8 คน': '八强赛', 'รอบ 4 คน': '四强赛', 'รอบรองชนะเลิศ': '半决赛', 'รอบชิงชนะเลิศ': '决赛', 'ตารางประกบ': '对阵表', 'หน้าประกบคู่': '对阵页面', 'พิมพ์ Scorecard หลายคน': '打印多张记分卡', 'บันทึกเป็นรูปภาพ': '保存为图片', 'พิมพ์ตาราง': '打印表格',
      'รอคิว': '等待中', 'กำลังตี': '进行中', 'ตีเสร็จแล้ว': '已完成', 'เข้ารอบตรง': '直接晋级', 'ผ่านจากรอบ 2': '第二轮晋级', 'มีสิทธิ์รอบ 2': '可进入第二轮', 'ตกรอบ': '淘汰', 'ลำดับการตี': '击球顺序', 'ลำดับในสนาม': '场地内顺序', 'สถานะ': '状态', 'หมายเลข': '号码', 'ชื่อ': '姓名', 'สังกัด': '所属单位',
      'เพิ่มนักกีฬาและจัดลำดับยิง': '添加运动员并安排击球顺序', 'ไปหน้ารวม': '前往总览', 'หมายเลขและลำดับรวม ระบบสร้างให้อัตโนมัติ': '号码和总顺序由系统自动生成', 'เพิ่มนักกีฬา': '添加运动员', 'การกระจายสนามอัตโนมัติ': '自动分配场地', 'สนาม = วนตามจำนวนสนาม เช่น 1, 2, 3, 4 แล้วกลับไป 1': '场地按数量循环分配，例如 1、2、3、4 后回到 1', 'สุ่มลำดับทั้งหมด': '随机全部顺序', 'นำเข้ารายชื่อนักกีฬาจาก Excel': '从 Excel 导入运动员名单', 'ไฟล์ต้องเป็น .xlsx และต้องมี 2 คอลัมน์คือ': '文件必须为 .xlsx，且包含两列：', 'เลือกไฟล์ Excel': '选择 Excel 文件', 'อัปโหลด Excel': '上传 Excel', 'ดาวน์โหลดไฟล์ตัวอย่าง': '下载模板文件', 'ยังไม่มีนักกีฬา': '暂无运动员',
      'ประเภทสุดยอดความแม่นยำ (SHOOTING)': '精准射击', 'กลับหน้ารวม': '返回总览', 'พิมพ์แบบฟอร์ม': '打印表单', 'รอบที่คีย์': '录入轮次', 'สถานีที่': '站点', 'ยิงลูกเดี่ยว': '单球射击', 'ยิงลูกเหนือเป้า': '越过目标球射击', 'ยิงลูกกลางขวาง': '中间障碍射击', 'ยิงลูกขาวเหนือลูกดำ': '白球越过黑球射击', 'ยิงลูกเป้า': '目标球射击', 'รอบแข่งขัน': '比赛轮次', 'รวม': '合计', 'แดง': '红牌',
      'กรรมการตัดสิน': '裁判', 'กรรมการบันทึกคะแนน': '记录员', 'ลายเซ็นกรรมการตัดสิน': '裁判签名', 'ลายเซ็นกรรมการบันทึกคะแนน': '记录员签名', 'ลายเซ็นนักกีฬา': '运动员签名', 'พิมพ์ชื่อ ถ้าไม่เซ็นระบบจะใช้ชื่อนี้': '输入姓名；未签名时使用此姓名', 'เซ็นชื่อเต็มหน้าจอ': '全屏签名', 'ล้างลายเซ็น': '清除签名', 'แตะปุ่มเพื่อเปิดพื้นที่เซ็นเต็มหน้าจอ': '点击按钮打开全屏签名区域', 'รหัสข้ามการลงชื่อ (admin)': '跳过签名代码（管理员）', 'กรอกเมื่อจำเป็น': '必要时填写', 'จบการตี': '结束击球', 'ลงลายเซ็น': '签名', 'เซ็นเต็มหน้าจอ แล้วกดบันทึกเพื่อกลับไปยังฟอร์ม': '全屏签名后点击保存返回表单', 'ปิด': '关闭', 'ล้าง': '清除', 'พิมพ์': '打印',
      'ลำดับ': '名次', 'ลายเซ็น / ผู้เกี่ยวข้องแต่ละรอบ': '每轮签名 / 相关人员', 'รอบ': '轮次', 'กรรมการ': '裁判', 'ผู้บันทึก': '记录员', 'กลับหน้าเลือกนักกีฬา': '返回运动员选择页', 'พิมพ์ทั้งหมดที่แสดง': '打印当前显示全部', 'จำนวน': '数量', 'คน': '人', 'ยังไม่มีนักกีฬาสำหรับพิมพ์ในรอบนี้': '本轮暂无可打印运动员', 'เลือกพิมพ์ Scorecard': '选择打印记分卡', 'เลือกรอบ': '选择轮次', 'เปลี่ยนรอบ': '切换轮次', 'เลือกทั้งหมด': '全选', 'ยกเลิกทั้งหมด': '取消全选', 'ยังไม่มีนักกีฬาในรอบนี้': '本轮暂无运动员', 'พิมพ์คนที่เลือก': '打印所选', 'พิมพ์ทั้งหมดในรอบนี้': '打印本轮全部',
      'คะแนนสูงสุดของประเภทนี้': '该类别最高分', 'อันดับ': '排名', 'ครั้ง': '次', 'ตีเที่ยวพิเศษ 7 เมตรทุกสถานี': '加赛：每个站点 7 米', 'เที่ยวพิเศษ': '加赛', 'บันทึกเที่ยวพิเศษ': '保存加赛', 'กดที่ชื่อในแต่ละ match box เพื่อเปิด scorecard ของรอบนั้น': '点击对阵框中的姓名打开该轮记分卡', 'ยังไม่มีคู่ Quarter Final': '暂无四分之一决赛对阵', 'ยังไม่มีคู่ Semi Final': '暂无半决赛对阵', 'ยังไม่มีคู่ Final': '暂无决赛对阵', 'ยืนยันลบ': '确认删除', 'ลบอีเวนต์นี้?': '删除此赛事？'
    }
  };

  function translateText(text, dict) {
    if (!text || !dict) return text;
    const leading = text.match(/^\s*/)[0];
    const trailing = text.match(/\s*$/)[0];
    const trimmed = text.trim();
    if (Object.prototype.hasOwnProperty.call(dict, trimmed)) {
      return leading + dict[trimmed] + trailing;
    }

    let result = text;
    const keys = Object.keys(dict)
      .filter(key => key.length >= 5 || /[\s()/?:|]/.test(key))
      .sort((a, b) => b.length - a.length);
    for (const key of keys) {
      if (result.includes(key)) {
        result = result.split(key).join(dict[key]);
      }
    }
    return result;
  }

  function translateNode(root, dict) {
    if (!dict) return;
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, {
      acceptNode(node) {
        const parent = node.parentElement;
        if (!parent) return NodeFilter.FILTER_REJECT;
        const tag = parent.tagName;
        if (['SCRIPT', 'STYLE', 'TEXTAREA'].includes(tag)) return NodeFilter.FILTER_REJECT;
        if (!node.nodeValue.trim()) return NodeFilter.FILTER_SKIP;
        return NodeFilter.FILTER_ACCEPT;
      }
    });
    const nodes = [];
    while (walker.nextNode()) nodes.push(walker.currentNode);
    nodes.forEach(node => {
      node.nodeValue = translateText(node.nodeValue, dict);
    });

    root.querySelectorAll('[placeholder], [title], [alt], [data-title], [aria-label]').forEach(el => {
      ['placeholder', 'title', 'alt', 'data-title', 'aria-label'].forEach(attr => {
        if (el.hasAttribute(attr)) {
          el.setAttribute(attr, translateText(el.getAttribute(attr), dict));
        }
      });
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
      translateNode(document.body, dict);
      installDialogTranslation(dict);
    }
  });

  window.translateI18nText = function (text) {
    return LANG === 'th' ? text : translateText(text, translations[LANG]);
  };
})();
