import React, { useEffect, useState } from "react";
import ExcelJS from "exceljs";

const styles = {
  page: {
    minHeight: "100vh",
    background: "linear-gradient(135deg, #f5f7fa, #e4ecf5)",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    fontFamily: "system-ui, -apple-system, BlinkMacSystemFont",
  },
  card: {
    background: "#fff",
    width: 420,
    padding: "32px",
    borderRadius: 16,
    boxShadow: "0 20px 40px rgba(0,0,0,0.08)",
    textAlign: "right",
  },
  logo: {
    display: "block",
    margin: "0 auto 16px",
    width: 160,
    objectFit: "contain",
  },
  title: {
    margin: 0,
    marginBottom: 8,
    fontSize: 22,
    fontWeight: 700,
    color: "#1f2937",
    textAlign: "center",
  },
  subtitle: {
    margin: 0,
    marginBottom: 24,
    fontSize: 14,
    color: "#6b7280",
    lineHeight: 1.6,
    textAlign: "center",
  },
  inputGroup: { marginBottom: 20 },
  label: {
    display: "block",
    marginBottom: 8,
    fontSize: 14,
    fontWeight: 600,
    color: "#374151",
  },
  fileInput: {
    width: "100%",
    padding: "10px",
    borderRadius: 10,
    border: "1px solid #d1d5db",
    background: "#f9fafb",
    cursor: "pointer",
  },
  button: (disabled) => ({
    width: "100%",
    marginTop: 10,
    padding: "14px 0",
    borderRadius: 12,
    border: "none",
    background: "linear-gradient(135deg, #2563eb, #1d4ed8)",
    color: "#fff",
    fontSize: 15,
    fontWeight: 700,
    cursor: disabled ? "not-allowed" : "pointer",
    opacity: disabled ? 0.6 : 1,
  }),
};

function PriceUpdater() {
  const [baseWorkbook, setBaseWorkbook] = useState(null);
  const [userWorkbook1, setUserWorkbook1] = useState(null);
  const [userWorkbook2, setUserWorkbook2] = useState(null);
  const [userData1, setUserData1] = useState([]);
  const [userData2, setUserData2] = useState([]);
  const [date1, setDate1] = useState(null);
  const [date2, setDate2] = useState(null);
  const [error, setError] = useState(null);

  // تابع نرمال‌سازی پیشرفته برای مقایسه محصولات
  const normalizeProductForCompare = (s) => {
    if (!s) return "";
    let t = String(s);

    // جایگزینی حروف فارسی و عربی
    t = t.replace(/ي/g, "ی").replace(/ك/g, "ک");

    // حذف کاراکترهای نامرئی و کنترل‌ها
    t = t.replace(/[\u200B-\u200F\u202A-\u202E]/g, "");

    // کوچک کردن حروف انگلیسی
    t = t.toLowerCase();

    // جدا کردن حروف از اعداد (S70 -> S 70)
    t = t.replace(/([a-zA-Z])(\d)/g, "$1 $2").replace(/(\d)([a-zA-Z])/g, "$1 $2");

    // حذف فاصله اضافی و مرتب‌سازی کلمات
    t = t.replace(/\s+/g, " ").trim();
    t = t.split(" ").sort().join(" ");

    return t;
  };

  const normalize = (s) =>
    String(s || "")
      .replace(/ي/g, "ی")
      .replace(/ك/g, "ک")
      .replace(/‌/g, "")
      .replace(/\s+/g, " ")
      .trim();

  const getCellValue = (cell) => {
    if (!cell || cell.value == null) return null;
    if (typeof cell.value === "object") {
      if (cell.value.richText)
        return cell.value.richText.map((t) => t.text).join("");
      if (cell.value.formula) return cell.value.result;
    }
    return cell.value;
  };

  const extractDate = (ws) => {
    const regex =
      /((13|14)\d{2}\/\d{1,2}\/\d{1,2})|(\d{4}-\d{2}-\d{2})/;
    for (let r = 1; r <= 40; r++) {
      for (let c = 1; c <= 40; c++) {
        const v = getCellValue(ws.getCell(r, c));
        if (!v) continue;
        const m = String(v).match(regex);
        if (m) return m[0];
      }
    }
    return null;
  };

  const extractReversedDecimal = (worksheets) => {
    for (const ws of worksheets) {
      for (let r = 1; r <= ws.rowCount; r++) {
        for (let c = 1; c <= Math.min(10, ws.columnCount); c++) {
          const text = String(getCellValue(ws.getCell(r, c)) || "");
          if (text.includes("نرخ ارز") || text.includes("دلار")) {
            const m = text.match(/[\d.]+/);
            if (m) {
              const parts = m[0].split(".");
              const decimalPart = parts[1] || "0";
              return decimalPart.split("").reverse().join("");
            }
          }
        }
      }
    }
    return null;
  };

  useEffect(() => {
    const loadBase = async () => {
      const res = await fetch("/base.xlsx");
      const buf = await res.arrayBuffer();
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(buf);
      setBaseWorkbook(wb);
    };
    loadBase();
  }, []);

  const parseUserExcelAllSheets = async (file) => {
    const buf = await file.arrayBuffer();
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buf);

    let allData = [];
    let foundDate = null;

    for (const ws of wb.worksheets) {
      if (!foundDate) foundDate = extractDate(ws);

      let headerRow = -1,
        productCol = -1,
        priceCol = -1;

      for (let r = 1; r <= ws.rowCount; r++) {
        for (let c = 1; c <= ws.columnCount; c++) {
          const v = normalize(getCellValue(ws.getCell(r, c)));
          if (v === "محصول" || v === "نام محصول") productCol = c;
          if (v.includes("قیمت")) priceCol = c;
        }
        if (productCol !== -1 && priceCol !== -1) {
          headerRow = r;
          break;
        }
        productCol = -1;
        priceCol = -1;
      }

      if (headerRow === -1) continue;

      for (let r = headerRow + 1; r <= ws.rowCount; r++) {
        const product = normalize(getCellValue(ws.getCell(r, productCol)));
        const price = getCellValue(ws.getCell(r, priceCol));
        if (product) allData.push({ محصول: product, قیمت: price });
      }
    }

    if (!allData.length)
      throw new Error("هیچ داده معتبری در فایل پیدا نشد");

    return { wb, data: allData, date: foundDate };
  };

  const handleUserExcel = async (e, index) => {
    try {
      const file = e.target.files[0];
      if (!file) return;
      const parsed = await parseUserExcelAllSheets(file);

      if (index === 1) {
        setUserWorkbook1(parsed.wb);
        setUserData1(parsed.data);
        setDate1(parsed.date);
      } else {
        setUserWorkbook2(parsed.wb);
        setUserData2(parsed.data);
        setDate2(parsed.date);
      }
    } catch (err) {
      setError(err.message);
    }
  };

  const applyPrices = async () => {
    try {
      if (!baseWorkbook) return;
      const ws = baseWorkbook.worksheets[0];
  
      let headerRow = -1,
        productCol = -1,
        date1Col = -1,
        date2Col = -1;
  
      // پیدا کردن ستون‌ها
      for (let r = 1; r <= ws.rowCount; r++) {
        for (let c = 1; c <= ws.columnCount; c++) {
          const v = normalize(getCellValue(ws.getCell(r, c)));
          if (v === "نام محصول") productCol = c;
          if (v === "تاریخ 1") date1Col = c;
          if (v === "تاریخ 2") date2Col = c;
        }
        if (productCol !== -1 && date1Col !== -1 && date2Col !== -1) {
          headerRow = r;
          break;
        }
      }
  
      const updateSheet = (userData, userWb, targetCol) => {
        if (!userData || !userWb) return;
  
        const map = new Map();
        userData.forEach((r) => {
          const normalizedName = normalizeProductForCompare(r.محصول);
  
          // استثناء S57
          if (normalizedName.includes("s 57")) {
            if (r.مجتمع && normalize(r.مجتمع).includes("آبادان")) {
              map.set(normalizedName, r.قیمت);
            }
          } else {
            if (!map.has(normalizedName)) {
              map.set(normalizedName, r.قیمت);
            }
          }
        });
  
        const reversedDecimal = extractReversedDecimal(userWb.worksheets);
  
        for (let r = headerRow + 1; r <= ws.rowCount; r++) {
          const productCell = ws.getCell(r, productCol);
          const originalName = getCellValue(productCell) || "";
          const name = normalizeProductForCompare(originalName);
  
          const cell = ws.getCell(r, targetCol);
  
          if (originalName.includes("نرخ دلار") && reversedDecimal) {
            cell.value = Number(reversedDecimal);
            continue;
          }
  
          if (map.has(name)) {
            cell.value = Number(map.get(name));
          }
        }
      };
  
      // اعمال هر دو فایل روی sheet
      updateSheet(userData1, userWorkbook1, date1Col);
      updateSheet(userData2, userWorkbook2, date2Col);
  
      // 🔹 تغییر نام محصولات فقط بعد از اعمال هر دو فایل
      const renameMap = new Map([
        ["EPVC 7244 H", "پلی وینیل کلراید E 7244"],
        ["EPVC 7544 M", "پلی وینیل کلراید E 7544"],
        ["پلی پروپیلن نساجی Z30S", "پلی پروپیلن نساجی"],
        ["پلی اتیلن سنگین بادی 0035", "پلی اتیلن سنگین بادی"],
        ["اکریلونیتریل بوتادین استایرن 0150", "اکریلونیتریل بوتادین استایرن(0150و50 گرید طبیعی)"],
        ["پلی استایرن معمولی 1551", "پلی استایرن معمولی(1551و3160و1540)"],
        ["پلی استایرن انبساطی نسوز  200-F", "پلی استایرن انبساطی نسوزF(100,200,300)"],
        ["پلی اتیلن سنگین دورانی 3840UA", "پلی اتیلن سنگین دورانی (3840UA)"],
        ["پلی اتیلن سبک فیلم 0200", "پلی اتیلن سبک فیلم (0200,2119,0075)"],
        ["استایرن منومر*", "استایرن منومر (تلفیقی)"],
        ["پلی اتیلن سبک فیلم 2420E‏02", "پلی اتیلن سبک فیلم 2420E02‏"],
        ["آمونیاک (گاز)", "آمونیاک (گاز,مایع)"],
        ["پلی اتیلن سنگین فیلم EX5", "پلی اتیلن سنگین فیلم (EX5,F7000,5110)"],
        ["پلی اتیلن سنگین تزریقی I‏4", "پلی اتیلن سنگین تزریقی(HI0500, 62N07UV,I4)"],
        ["پلی پروپیلن فیلم HP525J", "پلی پروپیلن فیلم"],
                ["پلی وینیل کلراید E 60", "پلی وینیل کلراید (60,6644)E "],







        


      ]);
  
      for (let r = headerRow + 1; r <= ws.rowCount; r++) {
        const cell = ws.getCell(r, productCol);
        const name = getCellValue(cell)?.trim();
        if (name && renameMap.has(name)) {
          cell.value = renameMap.get(name);
        }
      }
  
      const applyDateToBase = (label, value) => {
        if (!value) return;
        for (let r = 1; r <= ws.rowCount; r++) {
          for (let c = 1; c <= ws.columnCount; c++) {
            const cell = ws.getCell(r, c);
            if (normalize(getCellValue(cell)) === label) {
              cell.value = value;
            }
          }
        }
      };
  
      applyDateToBase("تاریخ 1", date1);
      applyDateToBase("تاریخ 2", date2);
  
      for (let r = 1; r <= ws.rowCount; r++) {
        for (let c = 1; c <= ws.columnCount; c++) {
          const cell = ws.getCell(r, c);
          if (cell.value !== null) {
            cell.font = { name: "B Nazanin", bold: true, size: 14 };
          }
        }
      }
  
      ["F", "G"].forEach((col) => {
        ws.addConditionalFormatting({
          ref: `${col}${headerRow + 1}:${col}${ws.rowCount}`,
          rules: [
            {
              type: "cellIs",
              operator: "lessThan",
              formulae: ["0"],
              style: {
                fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FFFFC7CE" } },
                font: { color: { argb: "FF000000" } },
              },
            },
          ],
        });
      });
  
      baseWorkbook.calcProperties.fullCalcOnLoad = true;
  
      const buffer = await baseWorkbook.xlsx.writeBuffer();
      const blob = new Blob([buffer]);
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = "output.xlsx";
      a.click();
    } catch (err) {
      setError("خطا در تولید فایل خروجی: " + err.message);
    }
  };

  return (
    <div style={styles.page}>
      <div style={styles.card}>
        <img src="/logo.jpg" alt="لوگو" style={styles.logo} />
        <h2 style={styles.title}>📊 بروزرسانی قیمت‌ها</h2>
        <p style={styles.subtitle}>
          فایل‌های اکسل ورودی را بارگذاری کرده و خروجی نهایی را دانلود کنید
        </p>

        <div style={styles.inputGroup}>
          <label style={styles.label}>اکسل هفته گذشته</label>
          <input
            type="file"
            onChange={(e) => handleUserExcel(e, 1)}
            style={styles.fileInput}
          />
        </div>

        <div style={styles.inputGroup}>
          <label style={styles.label}>اکسل هفته جاری</label>
          <input
            type="file"
            onChange={(e) => handleUserExcel(e, 2)}
            style={styles.fileInput}
          />
        </div>

        <button
          onClick={applyPrices}
          style={styles.button(!userWorkbook1 && !userWorkbook2)}
          disabled={!userWorkbook1 && !userWorkbook2}
        >
          ⬇ دانلود خروجی اکسل
        </button>

        {error && <div style={{ color: "red", marginTop: 10 }}>{error}</div>}
      </div>
    </div>
  );
}

export default PriceUpdater;
