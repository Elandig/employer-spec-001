const excel = require("xlsx-js-style");
const fs = require("fs");

// You can define styles as json object
const styles = {
  headerGrey: {
    fill: {
      type: "pattern",
      patternType: "solid",
      fgColor: "FFDEE6EF",
    },
    border: {
      top: { style: "thin", color: "#404040" },
      bottom: { style: "thin", color: "#404040" },
      left: { style: "thin", color: "#404040" },
      right: { style: "thin", color: "#404040" },
    },
    font: {
      color: "FF000000",
      name: "Arial",
      size: 10,
      bold: true,
      underline: false,
    },
    alignment: {
      vertical: "top",
      horizontal: "center",
    },
    wrapText: true,
  },
  cellNum: {
    numberFormat: "#,##0",
    font: { name: "Arial", size: 10 },
    alignment: {
      vertical: "top",
      horizontal: "right",
    },
    border: {
      top: { style: "thin", color: "#404040" },
      bottom: { style: "thin", color: "#404040" },
      left: { style: "thin", color: "#404040" },
      right: { style: "thin", color: "#404040" },
    },
  },
  cellPercent: {
    numberFormat: "0.0%",
    font: { name: "Arial", size: 10 },
    alignment: {
      vertical: "top",
      horizontal: "right",
    },
    border: {
      top: { style: "thin", color: "#404040" },
      bottom: { style: "thin", color: "#404040" },
      left: { style: "thin", color: "#404040" },
      right: { style: "thin", color: "#404040" },
    },
  },
  cellCenter: {
    alignment: {
      vertical: "top",
      horizontal: "center",
    },
    numberFormat: "0",
    font: { name: "Arial", size: 10 },
    border: {
      top: { style: "thin", color: "#404040" },
      bottom: { style: "thin", color: "#404040" },
      left: { style: "thin", color: "#404040" },
      right: { style: "thin", color: "#404040" },
    },
  },
  cellQuantity: {
    alignment: {
      vertical: "top",
      horizontal: "center",
    },
    numberFormat: "0.0##",
    font: { name: "Arial", size: 10 },
    border: {
      top: { style: "thin", color: "#404040" },
      bottom: { style: "thin", color: "#404040" },
      left: { style: "thin", color: "#404040" },
      right: { style: "thin", color: "#404040" },
    },
  },
  cellDate: {
    alignment: {
      vertical: "top",
      horizontal: "center",
    },
    numberFormat: "dd.mm.yy",
    font: { name: "Arial", size: 10 },
    border: {
      top: { style: "thin", color: "#404040" },
      bottom: { style: "thin", color: "#404040" },
      left: { style: "thin", color: "#404040" },
      right: { style: "thin", color: "#404040" },
    },
  },
  cellDateTime: {
    alignment: {
      vertical: "top",
      horizontal: "center",
    },
    numberFormat: "yyyy-mm-dd hh:mm",
    font: { name: "Arial", size: 10 },
    border: {
      top: { style: "thin", color: "#404040" },
      bottom: { style: "thin", color: "#404040" },
      left: { style: "thin", color: "#404040" },
      right: { style: "thin", color: "#404040" },
    },
  },
  cellDefault: {
    alignment: {
      vertical: "top",
    },
    font: { name: "Arial", size: 10 },
    border: {
      top: { style: "thin", color: "#404040" },
      bottom: { style: "thin", color: "#404040" },
      left: { style: "thin", color: "#404040" },
      right: { style: "thin", color: "#404040" },
    },
  },
};

function getXLSX(data) {
  //Array of objects representing heading rows (very top)
  // const heading = [
  // 	[
  // 		{ value: "Заголовок 1", style: styles.headerDark },
  // 		{ value: "b1", style: styles.headerDark },
  // 		{ value: "c1", style: styles.headerDark },
  // 	],
  // 	["Заголовок 2", "пояснение", "еще"], // <-- It can be only values
  // ];
  // The data set should have the following shape (Array of Objects)
  // <- Notice that this is an array. Pass multiple sheets to create multi sheet report
  //     {
  //       name: "mshop", // <- Specify sheet name (optional)
  //       heading: heading, // <- Raw heading array (optional)
  //       specification: specification1, // <- Report specification
  //       data: data.sheet1, // <-- Report data
  //     },
  //   ]);
  const report = buildExport(data.sheets);
  if (report) return excel.write(report, { type: "buffer", compression: true });
}

// Определение типа для xlsx-js-style
function determineType(value) {
  if (value == null) value = undefined;
  switch (typeof value) {
    case "number":
      return "n";
    case "undefined":
      return "s";
    default:
      if (typeof value.getMonth === "function") {
        return "d";
      }
      return "s";
  }
}

function replaceLegacyStyles(style) {
  for (const key in style) {
    if (!style[key]) continue;

    if (typeof style[key] === "object") {
      style[key] = replaceLegacyStyles(style[key]);
      continue;
    }

    if (typeof style[key] === "string" && key.toLowerCase().includes("color")) {
      style[key] = {
        rgb: style[key].startsWith("#") ? style[key].slice(1) : style[key],
      };
    }

    if (key === "numberFormat") {
      style["numFmt"] = style[key].toUpperCase();
      delete style[key];
    }

    if (typeof style[key] === "number" && key === "size") {
      style.sz = String(style[key]);
      delete style[key];
    }
  }
  return style;
}

function setCell({ worksheet = [], value, row, column, style = {} }) {
  worksheet[row] ||= [];

  // Для совместимости legacy стилей с новой библиотекой
  style = replaceLegacyStyles(structuredClone(style));

  let type = determineType(value);

  worksheet[row][column] = {
    v: value || "",
    s: style,
    t: type,
    z: style.numFmt, // Указание numFmt в стилях не применяется к дате
  };
}

const defaultStyle = {
  font: {
    sz: "12",
  },
};

function buildExport(sheets) {
  const workbook = excel.utils.book_new();

  sheets.forEach((sheet) => {
    if (!sheet.specification) return;
    let worksheet = [];
    let merges = {};
    let colwidths = [];
    let heading = sheet.heading || [];
    let headrow = heading.length + 1;

    heading.forEach((r, rn) => {
      if (r instanceof Array) {
        r.forEach((val, cn) => {
          let m = { row: rn, column: cn };
          if (val && typeof val === "object" && val.style) {
            m.value = val.value;
            m.style = { ...defaultStyle, ...val.style };
          } else {
            m.value = val;
            m.style = defaultStyle;
          }
          setCell({ worksheet, ...m });
        });
      }
    });

    Object.keys(sheet.specification).forEach((colname, colno) => {
      let spec = sheet.specification[colname];
      let m = {
        value: spec.displayName,
        row: headrow - 1,
        column: colno,
      };
      if (styles[spec.headerStyle]) {
        m.style = { ...defaultStyle, ...styles[spec.headerStyle] };
      }
      setCell({ worksheet, ...m });
      colwidths[colno] = spec.width;
    });

    sheet.data.forEach((row, rowno) => {
      Object.keys(sheet.specification).forEach((colname, colno) => {
        let value = row[colname];
        let spec = sheet.specification[colname];
        let res = {};
        let sf;

        if (spec.styleFunc && typeof spec.styleFunc === "function") {
          sf = spec.styleFunc(value, row);
        }

        if (spec.beforeWrite && typeof spec.beforeWrite === "function") {
          res = spec.beforeWrite(value, {
            dataset: sheet.data,
            row,
            rowno,
            colname,
          });
          value = res.newvalue;
        }

        let m = { value, style: defaultStyle };

        if (styles[spec.cellStyle]) {
          m.style = { ...m.style, ...styles[spec.cellStyle] };
        }
        if (sf) {
          m.style = { ...m.style, ...sf };
        }
        if (res.style) {
          m.style = { ...m.style, ...res.style }; // стили beforeWrite применяем последними
        }

        // Индексируем через начальную клетку чтобы избежать дубликатов
        if (res.merges) {
          let index =
            "R" +
            (headrow + rowno - res.merges.up) +
            "C" +
            (colno - res.merges.left);
          if (
            merges[index] &&
            (merges[index].e.r < headrow + rowno || merges[index].e.c < colno)
          ) {
            merges[index].e = { r: headrow + rowno, c: colno };
          } else {
            merges[index] = {
              s: {
                r: headrow + rowno - res.merges.up,
                c: colno - res.merges.left,
              },
              e: { r: headrow + rowno, c: colno },
            };
          }
        }

        m.row = headrow + rowno;
        m.column = colno;
        setCell({ worksheet, ...m });
      });
    });

    const ws = excel.utils.aoa_to_sheet(worksheet);

    ws["!cols"] ||= [];
    ws["!rows"] ||= [];

    // Так как xlsx-js-style округляет до десятых высоту строк, добавляем небольшой margin
    // для корректного отображения. Большинство редакторов вовсе игнорируют эти значения.
    ws["!rows"].forEach((row, rowno) => (ws["!rows"][rowno] = { hpt: 12.85 })); // 12.8pt
    colwidths.forEach(
      (width, key) => (ws["!cols"][key] = { wch: width + 0.2 })
    );

    ws["!merges"] = Object.values(merges);

    excel.utils.book_append_sheet(workbook, ws, sheet.name);
  });
  return workbook;
}

module.exports = { getXLSX, styles };
