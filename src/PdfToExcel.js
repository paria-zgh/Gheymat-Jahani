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
  const fillNumberAfterRadefAuto = (ws) => {
    for (let r = 1; r <= ws.rowCount; r++) {
      for (let c = 1; c <= ws.columnCount; c++) {
        const cell = ws.getCell(r, c);
        if (getCellValue(cell) === "ردیف") {
          let tempRow = r + 1;
          let counter = 1;
          while (tempRow <= ws.rowCount) {
            const nextCell = ws.getCell(tempRow, c);
            const val = getCellValue(nextCell);

            // وقتی سلول شامل "*" شد، متوقف شود
            if (val && String(val).includes("*")) break;

            // شماره‌گذاری حتی اگر سلول خالی باشد
            nextCell.value = counter;
            counter++;
            tempRow++;
          }
        }
      }
    }
  };


  const applyPrices = async () => {
  try {
    if (!baseWorkbook) return;

    const colNumToLetter = (num) => {
      let letter = "";
      while (num > 0) {
        let mod = (num - 1) % 26;
        letter = String.fromCharCode(65 + mod) + letter;
        num = Math.floor((num - mod - 1) / 26);
      }
      return letter;
    };

    for (const ws of baseWorkbook.worksheets) {
      let r = 1;

      while (r <= ws.rowCount) {
        // 🔹 پیدا کردن هدر جدول
        let productCol = -1,
            date1Col = -1,
            date2Col = -1,
            diffCol = -1,
            percentCol = -1,
            diffRow = -1;

        for (let c = 1; c <= ws.columnCount; c++) {
          const v = normalize(getCellValue(ws.getCell(r, c)));
          if (v === "نام محصول") productCol = c;
          if (v === "تاریخ 1") date1Col = c;
          if (v === "تاریخ 2") date2Col = c;
          if (v === "اختلاف") {
            diffCol = c;
            diffRow = r;
          }
          if (v === "درصد تغییر") percentCol = c;
        }

        if (productCol === -1 || date1Col === -1 || date2Col === -1 || diffCol === -1) {
          r++;
          continue; // هنوز جدول پیدا نشده
        }

        // 🔹 اعمال تاریخ‌ها در هدر جدول
        ws.getCell(diffRow, date1Col).value = date1;
        ws.getCell(diffRow, date2Col).value = date2;

        // 🔹 پیدا کردن انتهای جدول
        let lastRow = diffRow + 1;
        while (lastRow <= ws.rowCount && getCellValue(ws.getCell(lastRow, productCol))) lastRow++;

        // 🔹 اعمال قیمت‌ها
        const updateSheetForTable = (userData, userWb, targetCol) => {
          if (!userData || !userWb) return;
          const map = new Map();
          userData.forEach((row) => {
            const normalizedName = normalizeProductForCompare(row.محصول);
            if (!map.has(normalizedName)) map.set(normalizedName, row.قیمت);
          });

          const reversedDecimal = extractReversedDecimal(userWb.worksheets);

          for (let rowNum = diffRow + 1; rowNum < lastRow; rowNum++) {
            const productCell = ws.getCell(rowNum, productCol);
            const originalName = getCellValue(productCell) || "";
            const name = normalizeProductForCompare(originalName);
            const cell = ws.getCell(rowNum, targetCol);

            if (originalName.includes("نرخ دلار") && reversedDecimal) {
              cell.value = Number(reversedDecimal);
              continue;
            }

            if (map.has(name)) cell.value = Number(map.get(name));
          }
        };

        updateSheetForTable(userData1, userWorkbook1, date1Col);
        updateSheetForTable(userData2, userWorkbook2, date2Col);

        // 🔹 تغییر نام محصولات
        const renameMap = new Map([
          ["EPVC 7244 H", "پلی وینیل کلراید E 7244"],
          ["EPVC 7544 M", "پلی وینیل کلراید E 7544"],
          ["پلی پروپیلن نساجی Z30S", "پلی پروپیلن نساجی"],
          ["پلی اتیلن سنگین بادی 0035", "پلی اتیلن سنگین بادی"],
          ["اکریلونیتریل بوتادین استایرن 0150", "اکریلونیتریل بوتادین استایرن(0150و50 گرید طبیعی)"],
          ["پلی استایرن معمولی 1551", "پلی استایرن معمولی(1551و3160و1540)"],
          ["پلی استایرن انبساطی نسوز  200-F", "پلی استایرن انبساطی نسوز(100,200,300)F"],
          ["پلی اتیلن سنگین دورانی 3840UA", "پلی اتیلن سنگین دورانی (3840UA)"],
          ["پلی اتیلن سبک فیلم 0200", "پلی اتیلن سبک فیلم (0200,2119,0075)"],
          ["استایرن منومر*", "استایرن منومر (تلفیقی)"],
          ["پلی اتیلن سبک فیلم 2420E‏02", "پلی اتیلن سبک فیلم 2420E02‏"],
          ["آمونیاک (گاز)", "آمونیاک (گاز,مایع)"],
          ["پلی اتیلن سنگین فیلم EX5", "پلی اتیلن سنگین فیلم (EX5,F7000,5110)"],
          ["پلی اتیلن سنگین تزریقی I‏4", "پلی اتیلن سنگین تزریقی(HI0500, 62N07UV,I4)"],
          ["پلی پروپیلن فیلم HP525J", "پلی پروپیلن فیلم"],
          ["پلی وینیل کلراید E 60", "پلی وینیل کلراید (60,6644)E "],
          ["پلی پروپیلن شیمیایی_ZR340R", "پلی پروپیلن شیمیاییZR340R "],
        ]);

        for (let rowNum = diffRow + 1; rowNum < lastRow; rowNum++) {
          const cell = ws.getCell(rowNum, productCol);
          const name = getCellValue(cell)?.trim();
          if (name && renameMap.has(name)) cell.value = renameMap.get(name);
        }

        // 🔹 مرتب‌سازی از ردیف دوم داده‌ها
        const dataStartRow = diffRow + 2;
        const rowsData = [];
        for (let rowNum = dataStartRow; rowNum < lastRow; rowNum++) {
          const val1 = Number(ws.getCell(rowNum, date1Col).value || 0);
          const val2 = Number(ws.getCell(rowNum, date2Col).value || 0);
          const diffCalc = val2 - val1;

          const rowValues = [];
          for (let c = 1; c <= ws.columnCount; c++)
            rowValues.push(ws.getCell(rowNum, c).value);

          rowsData.push({ values: rowValues, diffCalc });
        }
        rowsData.sort((a, b) => {
          const aSign = a.diffCalc >= 0 ? -1 : 1;
          const bSign = b.diffCalc >= 0 ? -1 : 1;
        
          if (aSign !== bSign) return aSign - bSign;
        
          const nameA = String(a.values[productCol - 1] || "").toLowerCase();
          const nameB = String(b.values[productCol - 1] || "").toLowerCase();
          return nameA.localeCompare(nameB);
        });
        

        rowsData.forEach((row, i) => {
          const targetRow = dataStartRow + i;
          for (let c = 1; c <= ws.columnCount; c++)
            ws.getCell(targetRow, c).value = row.values[c - 1];
        });

    

        for (let rowNum = diffRow + 1; rowNum < lastRow; rowNum++) {
          const val1 = Number(ws.getCell(rowNum, date1Col).value || 0);
          const val2 = Number(ws.getCell(rowNum, date2Col).value || 0);
          const diffValue = val2 - val1;
          const percentValue = val1 !== 0 ? diffValue / val1 : 0;
        
          ws.getCell(rowNum, diffCol).value = diffValue;

          if (percentCol !== -1) {
            ws.getCell(rowNum, percentCol).value =
              val1 !== 0 ? percentValue : 0;
          
            ws.getCell(rowNum, percentCol).numFmt = "0.00%";
          }
          
        }
        
        // 🔹 Conditional Formatting روی ستون اختلاف و درصد تغییر شامل ردیف اول و بقیه
        const columnsToFormat = [diffCol, percentCol];
        columnsToFormat.forEach((col) => {
          if (col === -1) return;
          const colLetter = colNumToLetter(col);
          ws.addConditionalFormatting({
            ref: `${colLetter}${diffRow + 1}:${colLetter}${lastRow - 1}`,
            rules: [
              {
                type: "cellIs",
                operator: "lessThan",
                formulae: ["0"],
                style: {
                  fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FFFFC7CE" } },
                  font: { color: { argb: "FF9C0006" } },
                },
              },
              {
                type: "cellIs",
                operator: "greaterThan",
                formulae: ["0"],
                style: {
                  fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FFC6EFCE" } },
                  font: { color: { argb: "FF006100" } },
                },
              },
            ],
          });
        });

        // 🔹 اعمال فونت
        for (let rowNum = diffRow + 1; rowNum < lastRow; rowNum++) {
          for (let c = 1; c <= ws.columnCount; c++) {
            const cell = ws.getCell(rowNum, c);
            if (cell.value !== null) cell.font = { name: "B Nazanin", bold: true, size: 14 };
          }
        }
        fillNumberAfterRadefAuto(ws);
// بعد از fillNumberAfterRadefAuto(ws)
for (let r = 1; r <= ws.rowCount; r++) {
  const cell = ws.getCell(r, 1); // ستون A یا ستونی که متن زیر جدول قرار دارد
  const val = getCellValue(cell);

  if (val && String(val).includes("*")) {
    const textRow = r + 1; // ردیف بعد از * متن زیر جدول
    if (textRow > ws.rowCount) continue;

    const textCell = ws.getCell(textRow, 1); // ستون متن
    const textVal = getCellValue(textCell);
    if (!textVal || typeof textVal !== "string") continue;

    let newText = textVal;
    newText = newText.replace(/تاریخ 1/g, date1 || "تاریخ 1");
    newText = newText.replace(/تاریخ 2/g, date2 || "تاریخ 2");

    if (newText !== textVal) textCell.value = newText;
  }
}



        // 🔹 رفتن به جدول بعدی
        r = lastRow + 1;
      }
    }

    // 🔹 دانلود فایل خروجی
    baseWorkbook.calcProperties.fullCalcOnLoad = true;
    const buffer = await baseWorkbook.xlsx.writeBuffer();
    const blob = new Blob([buffer]);
    const a = document.createElement("a");
    const url = URL.createObjectURL(blob);
    a.href = url;
    const safeDate1 = date1 ? date1.replace(/[\/\\:*?"<>|]/g, "-") : "بدون-تاریخ";
const safeDate2 = date2 ? date2.replace(/[\/\\:*?"<>|]/g, "-") : "بدون-تاریخ";

const fileName = `قیمت جهانی ${safeDate1} و ${safeDate2}.xlsx`;

a.download = fileName;
    a.click();
    URL.revokeObjectURL(url);

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
        <h2>
          قیمت جدید
        </h2>

        {error && <div style={{ color: "red", marginTop: 10 }}>{error}</div>}
      </div>
    </div>
  );
}

export default PriceUpdater;
