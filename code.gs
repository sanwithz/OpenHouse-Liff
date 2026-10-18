/**
 * Google Apps Script สำหรับระบบ Open House LIFF สะสมคะแนนและรับเกียรติบัตร
 * บันทึกข้อมูลลง Google Sheets: ตาราง 'Users' และ ตาราง 'Scans'
 *
 * รองรับข้อมูลผู้สมัคร: UID, Name, Email, Phone, Role (นักเรียน/ครู/ผู้ปกครอง/บุคคลทั่วไป), Province, School, GradeLevel
 * ⚡ ปรับปรุงประสิทธิภาพความเร็วสูง:
 * - ใช้ createTextFinder + findNext ในการค้นหาผู้ใช้ทันที (ระดับ Native C++ แทนการวน for loop)
 * - ใช้ filter, map ในการดึงและประมวลผลข้อมูล (หลีกเลี่ยง for loop)
 *
 * วิธีการติดตั้ง:
 * 1. เปิด Google Sheet ที่ต้องการใช้งาน
 * 2. ไปที่ Extensions (ส่วนขยาย) > Apps Script
 * 3. ลบโค้ดเดิมออกทั้งหมด แล้ววางโค้ดจากไฟล์นี้ลงไป
 * 4. บันทึกและกดเลือกฟังก์ชัน "initSheets" แล้วกด Run (เรียกใช้) หนึ่งครั้ง เพื่อสร้างหรืออัปเดตหัวตารางอัตโนมัติ
 * 5. กด Deploy (การทำให้ใช้งานได้) > New deployment (การทำให้ใช้งานได้รายการใหม่)
 * 6. เลือกประเภท: Web app (เว็บแอป)
 *    - Description: OpenHouse API v2 (Fast Search)
 *    - Execute as (ดำเนินการในฐานะ): Me (ฉัน)
 *    - Who has access (ผู้มีสิทธิ์เข้าถึง): Anyone (ทุกคน)  <--- สำคัญมาก!
 * 7. กด Deploy แล้วคัดลอก "Web app URL" มาใส่ในตัวแปร API_URL ใน index.html
 */

// ชื่อแท็บใน Google Sheet
const SHEET_USERS = "Users";
const SHEET_SCANS = "Scans";
const SHEET_BOOTHS = "Booths";

// รายการบูธกิจกรรมมาตรฐานเริ่มต้น
const DEFAULT_BOOTHS = [
  {
    code: "AI01",
    name: "บูธปัญญาประดิษฐ์และหุ่นยนต์ AI",
    points: 25,
    status: "Active",
  },
  {
    code: "SCI02",
    name: "บูธวิทยาศาสตร์และเทคโนโลยี",
    points: 20,
    status: "Active",
  },
  {
    code: "ASTRO03",
    name: "บูธดาราศาสตร์และอวกาศ",
    points: 20,
    status: "Active",
  },
  {
    code: "CHEM04",
    name: "บูธเคมีและการทดลองสุดว้าว",
    points: 15,
    status: "Active",
  },
  {
    code: "BIO05",
    name: "บูธชีววิทยาและความหลากหลายทางชีวภาพ",
    points: 15,
    status: "Active",
  },
  {
    code: "MATH06",
    name: "บูธคณิตศาสตร์และการคิดคำนวณ",
    points: 15,
    status: "Active",
  },
  {
    code: "PSU26",
    name: "บูธแนะนำหลักสูตร PSU Phuket",
    points: 20,
    status: "Active",
  },
  {
    code: "WELCOME",
    name: "บูธต้อนรับ Open House 2026",
    points: 10,
    status: "Active",
  },
];

const USER_COLUMNS = [
  "UID",
  "Name",
  "Email",
  "Phone",
  "Role",
  "Province",
  "School",
  "GradeLevel",
  "PictureUrl",
  "RegisteredAt",
  "UpdatedAt",
];

/**
 * ฟังก์ชันสร้างหัวตารางอัตโนมัติ (รันครั้งแรกหรือเมื่อต้องการอัปเดตโครงสร้าง)
 */
function initSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ตาราง Users
  let usersSheet = ss.getSheetByName(SHEET_USERS);
  if (!usersSheet) {
    usersSheet = ss.insertSheet(SHEET_USERS);
    usersSheet.appendRow(USER_COLUMNS);
    usersSheet.setFrozenRows(1);
    usersSheet
      .getRange(1, 1, 1, USER_COLUMNS.length)
      .setFontWeight("bold")
      .setBackground("#E0E7FF");
  } else {
    // ตรวจสอบและอัปเดตหัวตารางหากยังเป็นเวอร์ชันเดิม
    const headerRow = usersSheet
      .getRange(1, 1, 1, Math.max(usersSheet.getLastColumn(), 1))
      .getValues()[0];
    if (headerRow.indexOf("Email") === -1 || headerRow.indexOf("Role") === -1) {
      usersSheet
        .getRange(1, 1, 1, USER_COLUMNS.length)
        .setValues([USER_COLUMNS]);
      usersSheet
        .getRange(1, 1, 1, USER_COLUMNS.length)
        .setFontWeight("bold")
        .setBackground("#E0E7FF");
    }
  }

  // ตาราง Scans
  let scansSheet = ss.getSheetByName(SHEET_SCANS);
  if (!scansSheet) {
    scansSheet = ss.insertSheet(SHEET_SCANS);
    scansSheet.appendRow([
      "UID",
      "Event",
      "Points",
      "Date",
      "Timestamp",
      "Name",
      "School",
      "Role",
    ]);
    scansSheet.setFrozenRows(1);
    scansSheet
      .getRange(1, 1, 1, 8)
      .setFontWeight("bold")
      .setBackground("#DCFCE7");
  }

  // ตาราง Booths (สำหรับเพิ่ม/แก้ไข/ลบบูธผ่าน Google Sheets ได้ทันทีโดยไม่ต้องแก้โค้ด)
  let boothsSheet = ss.getSheetByName(SHEET_BOOTHS);
  if (!boothsSheet) {
    boothsSheet = ss.insertSheet(SHEET_BOOTHS);
    boothsSheet.appendRow(["Code", "Name", "Points", "Status"]);
    DEFAULT_BOOTHS.forEach(function (b) {
      boothsSheet.appendRow([b.code, b.name, b.points, b.status || "Active"]);
    });
    boothsSheet.setFrozenRows(1);
    boothsSheet
      .getRange(1, 1, 1, 4)
      .setFontWeight("bold")
      .setBackground("#FEF3C7");
  }
}

/**
 * ดึงรายชื่อบูธกิจกรรมทั้งหมดจากแท็บ 'Booths' ใน Google Sheets
 */
function getBoothsFromSheet(ss) {
  try {
    const sheet = ss.getSheetByName(SHEET_BOOTHS);
    if (!sheet || sheet.getLastRow() < 2) {
      return DEFAULT_BOOTHS.map(function (b) {
        return { code: b.code, name: b.name, points: b.points };
      });
    }

    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
    const booths = data
      .filter(function (row) {
        const code = String(row[0] || "").trim();
        const name = String(row[1] || "").trim();
        const status = String(row[3] || "Active")
          .trim()
          .toLowerCase();
        return code && name && status !== "inactive" && status !== "hidden";
      })
      .map(function (row) {
        return {
          code: String(row[0]).trim().toUpperCase(),
          name: String(row[1]).trim(),
          points: parseInt(row[2], 10) || 15,
        };
      });

    return booths.length > 0
      ? booths
      : DEFAULT_BOOTHS.map(function (b) {
          return { code: b.code, name: b.name, points: b.points };
        });
  } catch (err) {
    return DEFAULT_BOOTHS.map(function (b) {
      return { code: b.code, name: b.name, points: b.points };
    });
  }
}

/**
 * รองรับ GET Request: ดึงข้อมูลโปรไฟล์ผู้ใช้ ประวัติการสแกน และรายชื่อบูธกิจกรรม
 */
function doGet(e) {
  try {
    const params = e ? e.parameter : {};
    const uid = params.uid ? String(params.uid).trim() : "";
    const action = params.action || (uid ? "getUserData" : "ping");
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (action === "ping") {
      return jsonResponse({ status: "success", message: "API is running" });
    }

    // ดึงรายชื่อบูธทั้งหมดจาก Google Sheets โดยไม่ต้องใช้ UID
    if (action === "getBooths") {
      return jsonResponse({
        status: "success",
        booths: getBoothsFromSheet(ss),
      });
    }

    // 📊 ดึงข้อมูลสถิติผู้เข้าร่วมงานทั้งหมดจากแท็บ Users สำหรับ Analytics Dashboard
    if (action === "getDashboardData" || action === "getUsers") {
      return jsonResponse(getDashboardAnalyticsData(ss));
    }

    if (!uid) {
      return jsonResponse({
        status: "error",
        message: "Missing UID parameter",
      });
    }

    const user = findUserByUid(ss, uid);

    if (action === "getUser") {
      if (user) {
        return jsonResponse({ status: "success", user: user });
      } else {
        return jsonResponse({
          status: "not_found",
          message: "User not registered yet",
        });
      }
    }

    if (action === "getUserData" || action === "getScans") {
      const scans = getUserScans(ss, uid);
      const totalPoints = scans.reduce(
        (sum, s) => sum + (parseInt(s.points) || 0),
        0,
      );
      const booths = getBoothsFromSheet(ss);

      return jsonResponse({
        status: "success",
        isRegistered: !!user,
        user: user || null,
        scans: scans,
        totalPoints: totalPoints,
        booths: booths,
      });
    }

    return jsonResponse({ status: "error", message: "Unknown action" });
  } catch (err) {
    return jsonResponse({ status: "error", message: err.toString() });
  }
}

  /**
   * รองรับ POST Request: สมัครสมาชิก (Register/Update) หรือ บันทึกการสแกน (Scan)
   */
  function doPost(e) {
    const lock = LockService.getScriptLock();
    lock.tryLock(15000);

    try {
      if (!e || !e.postData || !e.postData.contents) {
        return jsonResponse({
          status: "error",
          message: "No POST data received",
        });
      }

      const data = JSON.parse(e.postData.contents);
      const action = data.action || (data.event ? "scan" : "register");
      const ss = SpreadsheetApp.getActiveSpreadsheet();

      // ---------------------------------------------
      // กรณีที่ 1: ลงทะเบียนผู้ใช้งาน (Register / Update Profile)
      // ---------------------------------------------
      if (action === "register" || action === "updateProfile") {
        const uid = String(data.uid || "").trim();
        const name = String(data.name || "").trim();
        const email = String(data.email || "").trim();
        const phone = String(data.phone || "").trim();
        const role = String(data.role || "นักเรียน").trim();
        const province = String(data.province || "").trim();
        const school = String(
          data.school ||
            (role === "ผู้ปกครอง" || role === "บุคคลทั่วไป" ? "-" : ""),
        ).trim();
        const gradeLevel = String(
          data.gradeLevel || (role === "นักเรียน" ? "" : "-"),
        ).trim();
        const pictureUrl = String(data.pictureUrl || "").trim();
        const now = new Date().toISOString();

        if (!uid || !name || !email || !province) {
          return jsonResponse({
            status: "error",
            message: "กรุณากรอกข้อมูลที่จำเป็นให้ครบถ้วน",
          });
        }

        if ((role === "นักเรียน" || role === "ครู") && !school) {
          return jsonResponse({
            status: "error",
            message: "กรุณาระบุชื่อโรงเรียน/วิทยาลัย",
          });
        }

        if (role === "นักเรียน" && !gradeLevel) {
          return jsonResponse({
            status: "error",
            message: "กรุณาเลือกระดับชั้นปี",
          });
        }

        let usersSheet = ss.getSheetByName(SHEET_USERS);
        if (!usersSheet) {
          initSheets();
          usersSheet = ss.getSheetByName(SHEET_USERS);
        }

        const headerValues = usersSheet
          .getRange(1, 1, 1, Math.max(usersSheet.getLastColumn(), 1))
          .getValues()[0];
        const hasModernHeaders = headerValues.indexOf("Email") !== -1;

        // ⚡ ใช้ createTextFinder + findNext ในการค้นหาผู้ใช้เดิมทันที (หลีกเลี่ยง for loop)
        const userCell = usersSheet
          .getRange("A:A")
          .createTextFinder(uid)
          .matchEntireCell(true)
          .findNext();
        const userRowIndex =
          userCell && userCell.getRow() > 1 ? userCell.getRow() : -1;

        const userData = {
          uid: uid,
          name: name,
          email: email,
          phone: phone,
          role: role,
          province: province,
          school: school || "-",
          gradeLevel: gradeLevel || "-",
          pictureUrl: pictureUrl,
          updatedAt: now,
        };

        if (hasModernHeaders) {
          // โครงสร้าง 11 คอลัมน์
          let registeredAt = now;
          if (userRowIndex > 0) {
            const oldRegAt = usersSheet.getRange(userRowIndex, 10).getValue();
            if (oldRegAt) registeredAt = oldRegAt;
          }

          const rowData = [
            uid,
            name,
            email,
            phone ? "'" + phone : "-",
            role,
            province,
            school || "-",
            gradeLevel || "-",
            pictureUrl,
            registeredAt,
            now,
          ];

          if (userRowIndex > 0) {
            usersSheet
              .getRange(userRowIndex, 1, 1, rowData.length)
              .setValues([rowData]);
          } else {
            usersSheet.appendRow(rowData);
          }
        } else {
          // รองรับกรณี Sheet เดิม 8 คอลัมน์
          if (userRowIndex > 0) {
            usersSheet.getRange(userRowIndex, 2).setValue(name);
            usersSheet
              .getRange(userRowIndex, 3)
              .setValue(email || (phone ? "'" + phone : "-"));
            usersSheet.getRange(userRowIndex, 4).setValue(province);
            usersSheet.getRange(userRowIndex, 5).setValue(school || "-");
            if (pictureUrl)
              usersSheet.getRange(userRowIndex, 6).setValue(pictureUrl);
            usersSheet.getRange(userRowIndex, 8).setValue(now);
          } else {
            usersSheet.appendRow([
              uid,
              name,
              email || (phone ? "'" + phone : "-"),
              province,
              school || "-",
              pictureUrl,
              now,
              now,
            ]);
          }
        }

        return jsonResponse({
          status: "success",
          message: "บันทึกข้อมูลเรียบร้อยแล้ว",
          user: userData,
        });
      }

      // ---------------------------------------------
      // กรณีที่ 2: บันทึกการสแกนรับคะแนน (Scan)
      // ---------------------------------------------
      if (action === "scan") {
        const uid = String(data.uid || "").trim();
        const eventName = String(data.event || data.กิจกรรม || "").trim();
        const points = parseInt(data.points || data.Point || 0, 10);
        const date = String(
          data.date || new Date().toISOString().split("T")[0],
        ).trim();
        const timestamp = data.timestamp || new Date().toISOString();
        const name = String(data.name || "").trim();
        const school = String(data.school || "").trim();
        const role = String(data.role || "").trim();

        if (!uid || !eventName) {
          return jsonResponse({
            status: "error",
            message: "ข้อมูลการสแกนไม่สมบูรณ์",
          });
        }

        let scansSheet = ss.getSheetByName(SHEET_SCANS);
        if (!scansSheet) {
          initSheets();
          scansSheet = ss.getSheetByName(SHEET_SCANS);
        }

        // ⚡ ป้องกันการสแกนซ้ำเด็ดขาด: ผู้ใช้ 1 คนไม่สามารถรับคะแนนจากบูธหรือ QR Code เดียวกันซ้ำได้
        const existingScanCells = scansSheet
          .getRange("A:A")
          .createTextFinder(uid)
          .matchEntireCell(true)
          .findAll();

        if (existingScanCells && existingScanCells.length > 0) {
          const normalizeEvent = function (str) {
            return String(str || "")
              .toLowerCase()
              .replace(/\s+/g, "")
              .trim();
          };
          const targetEventKey = normalizeEvent(eventName);

          // ดึงชื่อ Event ทั้งหมดที่ UID นี้เคยสแกนไว้ (Column 2)
          const isDuplicate = existingScanCells
            .filter(function (cell) {
              return cell.getRow() > 1;
            })
            .map(function (cell) {
              return scansSheet.getRange(cell.getRow(), 2).getValue();
            })
            .some(function (existingEvent) {
              return normalizeEvent(existingEvent) === targetEventKey;
            });

          if (isDuplicate) {
            return jsonResponse({
              status: "duplicate",
              message:
                'คุณเคยได้รับคะแนนจากบูธ "' +
                eventName +
                '" ไปแล้ว ไม่สามารถรับคะแนนซ้ำได้',
            });
          }
        }

        // บันทึกรายการใหม่
        scansSheet.appendRow([
          uid,
          eventName,
          points,
          date,
          timestamp,
          name,
          school,
          role,
        ]);

        return jsonResponse({
          status: "success",
          message: "บันทึกคะแนนสำเร็จ",
          points: points,
          event: eventName,
        });
      }

      // ---------------------------------------------
      // กรณีที่ 3: เพิ่มบูธกิจกรรมใหม่ (Add Booth)
      // ---------------------------------------------
      if (action === "addBooth") {
        const code = String(data.code || "")
          .trim()
          .toUpperCase();
        const name = String(data.name || "").trim();
        const points = parseInt(data.points || 15, 10);
        const status = String(data.status || "Active").trim();

        if (!code || !name) {
          return jsonResponse({
            status: "error",
            message: "กรุณาระบุรหัสและชื่อบูธ",
          });
        }

        let boothsSheet = ss.getSheetByName(SHEET_BOOTHS);
        if (!boothsSheet) {
          initSheets();
          boothsSheet = ss.getSheetByName(SHEET_BOOTHS);
        }

        boothsSheet.appendRow([code, name, points, status]);

        return jsonResponse({
          status: "success",
          message: "เพิ่มบูธลงใน Google Sheets สำเร็จ",
          booth: { code: code, name: name, points: points, status: status },
        });
      }

      return jsonResponse({ status: "error", message: "Unknown action" });
    } catch (err) {
      return jsonResponse({ status: "error", message: err.toString() });
    } finally {
      lock.releaseLock();
    }
  }

  /**
   * ฟังก์ชันค้นหาผู้ใช้จาก UID ในแท็บ Users
   * ⚡ ใช้ createTextFinder + findNext: ค้นหาโดยตรงใน Column A แบบ Native C++
   * ไม่ต้องวน for loop และดึงข้อมูลเฉพาะแถวที่เจอเท่านั้น
   */
  function findUserByUid(ss, uid) {
    const sheet = ss.getSheetByName(SHEET_USERS);
    if (!sheet || sheet.getLastRow() < 2) return null;

    // ⚡ ค้นหา UID ผ่าน createTextFinder และ findNext โดยตรงในคอลัมน์ A (Native C++ search หลีกเลี่ยง for loop)
    const cell = sheet
      .getRange("A:A")
      .createTextFinder(uid)
      .matchEntireCell(true)
      .findNext();
    if (!cell || cell.getRow() === 1) return null;

    const rowNum = cell.getRow();
    const lastCol = Math.max(sheet.getLastColumn(), 1);
    const headers = sheet
      .getRange(1, 1, 1, lastCol)
      .getValues()[0]
      .map((h) => String(h).trim().toLowerCase());
    const rowValues = sheet.getRange(rowNum, 1, 1, lastCol).getValues()[0];

    // แมป header เข้ากับค่าในแถวนั้นอย่างรวดเร็วด้วย reduce (หลีกเลี่ยง for loop)
    const userMap = headers.reduce((acc, header, idx) => {
      acc[header] =
        rowValues[idx] !== undefined && rowValues[idx] !== ""
          ? String(rowValues[idx]).trim()
          : "";
      return acc;
    }, {});

    return {
      uid: userMap["uid"] || uid,
      name: userMap["name"] || "",
      email: userMap["email"] || "",
      phone: (userMap["phone"] || "").replace(/^'/, ""),
      role: userMap["role"] || "นักเรียน",
      province: userMap["province"] || "",
      school: userMap["school"] || "",
      gradeLevel: userMap["gradelevel"] || "",
      pictureUrl: userMap["pictureurl"] || "",
      registeredAt: userMap["registeredat"] || "",
    };
  }

  /**
   * ฟังก์ชันดึงประวัติการสแกนของผู้ใช้จากแท็บ Scans
   * ⚡ ใช้ createTextFinder + Filter + Map (Map Filter) หลีกเลี่ยง for loop ทั้งหมด
   */
  function getUserScans(ss, uid) {
    const sheet = ss.getSheetByName(SHEET_SCANS);
    if (!sheet || sheet.getLastRow() < 2) return [];

    // ⚡ ค้นหาเฉพาะ UID ในคอลัมน์ A ด้วย createTextFinder
    const matchedCells = sheet
      .getRange("A:A")
      .createTextFinder(uid)
      .matchEntireCell(true)
      .findAll();
    if (!matchedCells || matchedCells.length === 0) return [];

    // กรองเฉพาะแถวข้อมูล (ข้ามแถวหัวตาราง)
    const dataCells = matchedCells.filter((cell) => cell.getRow() > 1);
    if (dataCells.length === 0) return [];

    const formatDate = function (val) {
      if (!val) return "";
      if (val instanceof Date) {
        return Utilities.formatDate(
          val,
          Session.getScriptTimeZone() || "Asia/Bangkok",
          "yyyy-MM-dd",
        );
      }
      return String(val).trim();
    };

    // ⚡ สำหรับผู้ใช้ทั่วไป (สแกน 1-15 ฐาน) ดึงเฉพาะแถวที่เจอตรงๆ ช่วยให้เร็วกว่าโหลดทั้งชีตหลายสิบเท่า
    if (dataCells.length <= 15) {
      return dataCells
        .map((cell) => sheet.getRange(cell.getRow(), 1, 1, 5).getValues()[0])
        .filter((row) => row && row[0] && String(row[0]).trim() === uid)
        .map((row) => ({
          uid: String(row[0]).trim(),
          event: String(row[1]).trim(),
          points: parseInt(row[2], 10) || 0,
          date: formatDate(row[3]),
          timestamp: row[4],
        }))
        .sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
    }

    // กรณีมีประวัติการสแกนจำนวนมาก ดึงเฉพาะช่วง 5 คอลัมน์ที่ใช้งานจริง แล้วใช้ Filter & Map
    const rows = sheet.getRange(1, 1, sheet.getLastRow(), 5).getValues();
    return rows
      .slice(1)
      .filter((row) => String(row[0]).trim() === uid)
      .map((row) => ({
        uid: String(row[0]).trim(),
        event: String(row[1]).trim(),
        points: parseInt(row[2], 10) || 0,
        date: formatDate(row[3]),
        timestamp: row[4],
      }))
      .sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
  }

  /**
   * 📊 ฟังก์ชันดึงและสรุปข้อมูลผู้ใช้งานทั้งหมดสำหรับระบบ Dashboard
   * สรุปมิติหลัก:
   * 1. จำนวนผู้มาร่วมงานจริง (นับจาก Unique UID ในชีท Scans ที่มีการสแกนสะสมคะแนน)
   * 2. ผู้เข้าร่วมงานมาจากจังหวัดไหนมากที่สุด (Provinces)
   * 3. มาจากโรงเรียนอะไรมากที่สุด (Schools)
   * 4. ระดับชั้นไหนสนใจเข้าร่วมงานมากที่สุด (Grade Levels)
   * รวมทั้งรายชื่อผู้เข้าร่วมงานสำหรับตารางแสดงผลและการกรองข้อมูล
   */
  function getDashboardAnalyticsData(ss) {
    // 🎟️ 1. อ่านชีท Scans เพื่อนับ Unique UID ของผู้มาร่วมงานจริง (Actual Attendees)
    const scanSheet = ss.getSheetByName(SHEET_SCANS);
    const attendedUidSet = {};
    const scanCountsByUid = {};
    let totalScanRecords = 0;

    if (scanSheet && scanSheet.getLastRow() >= 2) {
      const scanLastRow = scanSheet.getLastRow();
      // คอลัมน์แรก (Column A) ในชีท Scans คือ UID
      const scanRows = scanSheet.getRange(2, 1, scanLastRow - 1, 1).getValues();
      scanRows.forEach(function(row) {
        const uidVal = row && row[0] ? String(row[0]).trim() : "";
        if (uidVal) {
          attendedUidSet[uidVal] = true;
          scanCountsByUid[uidVal] = (scanCountsByUid[uidVal] || 0) + 1;
          totalScanRecords++;
        }
      });
    }

    const uniqueAttendedUids = Object.keys(attendedUidSet);
    const totalAttendees = uniqueAttendedUids.length;

    // 👥 2. อ่านชีท Users
    const sheet = ss.getSheetByName(SHEET_USERS);
    if (!sheet || sheet.getLastRow() < 2) {
      return {
        status: "success",
        totalUsers: 0,
        totalAttendees: totalAttendees,
        totalScans: totalScanRecords,
        attendanceRate: 0,
        users: [],
        summary: {
          byProvince: {},
          bySchool: {},
          byGradeLevel: {},
          byRole: {}
        },
        updatedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "Asia/Bangkok", "yyyy-MM-dd HH:mm:ss")
      };
    }

    const lastRow = sheet.getLastRow();
    const lastCol = Math.max(sheet.getLastColumn(), 1);
    const headers = sheet
      .getRange(1, 1, 1, lastCol)
      .getValues()[0]
      .map(function(h) { return String(h).trim().toLowerCase(); });

    const uidIdx = headers.indexOf("uid");
    const nameIdx = headers.indexOf("name");
    const roleIdx = headers.indexOf("role");
    const provIdx = headers.indexOf("province");
    const schoolIdx = headers.indexOf("school");
    const gradeIdx = headers.indexOf("gradelevel");
    const regAtIdx = headers.indexOf("registeredat");

    const rows = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

    const byProvince = {};
    const bySchool = {};
    const byGradeLevel = {};
    const byRole = {};
    let registeredAttendeesCount = 0;

    const users = rows
      .filter(function(row) { return row && row[uidIdx !== -1 ? uidIdx : 0]; })
      .map(function(row) {
        let regAtStr = "";
        const rawDate = regAtIdx !== -1 ? row[regAtIdx] : "";
        if (rawDate instanceof Date) {
          regAtStr = Utilities.formatDate(rawDate, Session.getScriptTimeZone() || "Asia/Bangkok", "yyyy-MM-dd HH:mm");
        } else if (rawDate) {
          regAtStr = String(rawDate).trim();
        }

        const uidVal = uidIdx !== -1 ? String(row[uidIdx]).trim() : "";
        const role = roleIdx !== -1 && row[roleIdx] ? String(row[roleIdx]).trim() : "นักเรียน";
        const province = provIdx !== -1 && row[provIdx] ? String(row[provIdx]).trim() : "ไม่ระบุ";
        const school = schoolIdx !== -1 && row[schoolIdx] ? String(row[schoolIdx]).trim() : "-";
        const gradeLevel = gradeIdx !== -1 && row[gradeIdx] ? String(row[gradeIdx]).trim() : "-";

        // ตรวจสอบว่าผู้ใช้คนนี้ได้มาร่วมงาน (สแกนกิจกรรม) หรือไม่
        const hasAttended = !!attendedUidSet[uidVal];
        const scanCount = scanCountsByUid[uidVal] || 0;
        if (hasAttended) {
          registeredAttendeesCount++;
        }

        // รวบรวมสถิติ
        if (province && province !== "-") {
          byProvince[province] = (byProvince[province] || 0) + 1;
        }
        if (school && school !== "-") {
          bySchool[school] = (bySchool[school] || 0) + 1;
        }
        if (gradeLevel && gradeLevel !== "-") {
          byGradeLevel[gradeLevel] = (byGradeLevel[gradeLevel] || 0) + 1;
        }
        byRole[role] = (byRole[role] || 0) + 1;

        return {
          uid: uidVal,
          name: nameIdx !== -1 ? String(row[nameIdx]).trim() : "",
          role: role,
          province: province,
          school: school,
          gradeLevel: gradeLevel,
          registeredAt: regAtStr,
          hasAttended: hasAttended,
          scanCount: scanCount
        };
      })
      .reverse(); // รายการล่าสุดขึ้นก่อน

    const totalUsers = users.length;
    const attendanceRate = totalUsers > 0 ? Math.round((totalAttendees / totalUsers) * 100) : 0;

    return {
      status: "success",
      totalUsers: totalUsers,
      totalAttendees: totalAttendees, // จำนวนผู้มาร่วมงานจริง (Unique UID ในชีท Scans)
      totalScans: totalScanRecords,   // จำนวนการสแกนสะสมทั้งหมด
      attendanceRate: attendanceRate, // อัตราส่วนผู้มาร่วมงาน (%)
      users: users,
      summary: {
        byProvince: byProvince,
        bySchool: bySchool,
        byGradeLevel: byGradeLevel,
        byRole: byRole
      },
      updatedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "Asia/Bangkok", "yyyy-MM-dd HH:mm:ss")
    };
  }

  /**
   * ฟังก์ชันส่ง JSON Response กลับไปยังเบราว์เซอร์
   */
  function jsonResponse(obj) {
    return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
      ContentService.MimeType.JSON,
    );
  }
