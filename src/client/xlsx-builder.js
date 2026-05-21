import { MIME_XLSX } from "./constants.js";
import {
  CALCULATED_STYLE_ID,
  CALCULATION_TYPE_OPERATORS,
  CENTER_DATE_STYLE_ID,
  CENTER_NUMBER_STYLE_ID,
  CENTER_TEXT_STYLE_ID,
  CUSTOM_CALCULATED_STYLE_ID,
  CUSTOM_DATE_STYLE_ID,
  CUSTOM_NUMBER_STYLE_ID,
  CUSTOM_ROW_AMOUNT_KEY,
  CUSTOM_ROW_FIELDS,
  CUSTOM_ROW_QUANTITY_KEY,
  CUSTOM_ROW_TYPE_KEY,
  CUSTOM_ROW_UNIT_PRICE_KEY,
  CUSTOM_TEXT_STYLE_ID,
  EMU_PER_PIXEL,
  EXPORT_NOTICE_TEXT,
  IMAGE_CELL_HEIGHT_PX,
  IMAGE_CELL_PADDING_PX,
  IMAGE_CELL_WIDTH_PX,
  SUPPORTED_IMAGE_MIMES,
  XML_NS_MAIN,
  XML_NS_REL,
  ZIP_EPOCH,
} from "./app-config.js";

const root = typeof window !== "undefined" ? window : globalThis;

function createXlsxBuilder(deps) {
  const {
    displayFieldLabel,
    findLast,
    getCalculatedFields,
    getRowValues,
    getSharedRemarkField,
    isAmountField,
    isCustomRow,
    isCustomTypeFallbackField,
    isQuantityField,
    isUnitPriceField,
    normalizeAppRow,
    normalizeCalculationRules,
    normalizeDisabledAutoCalculationTargets,
    normalizeFields,
    normalizeLabel,
    normalizeNonNegativeNumberValue,
    sameGroup,
  } = deps;

  function parseImageValue(value) {
    if (!value) {
      return null;
    }
  
    try {
      const image = JSON.parse(String(value));
      if (image && image.kind === "image" && image.dataUrl && image.type) {
        return image;
      }
    } catch (error) {
      return null;
    }
    return null;
  }
  
  function collectWorksheetImages(fields, rows) {
    const images = [];
    const dataStartRow = 5;
  
    rows.forEach((row, rowIndex) => {
      const values = getRowValues(row);
      fields.forEach((field, fieldIndex) => {
        if (field.type !== "image") {
          return;
        }
  
        const image = parseImageValue(values[field.key]);
        if (!image || !SUPPORTED_IMAGE_MIMES.includes(image.type)) {
          return;
        }
  
        const bytes = dataUrlToBytes(image.dataUrl);
        if (!bytes.length) {
          return;
        }
  
        images.push({
          index: images.length + 1,
          name: image.name || `图片 ${images.length + 1}`,
          type: image.type,
          extension: imageExtension(image.type),
          bytes,
          colIndex: fieldIndex,
          rowIndex: rowIndex + dataStartRow - 1,
          rowNumber: rowIndex + dataStartRow,
        });
      });
    });
  
    return images;
  }
  
  function rowHasImage(images, rowNumber) {
    return images.some((image) => image.rowNumber === rowNumber);
  }
  
  function imageExtension(type) {
    if (type === "image/jpeg") {
      return "jpg";
    }
    if (type === "image/gif") {
      return "gif";
    }
    return "png";
  }
  
  function imageContentType(extension) {
    if (extension === "jpg" || extension === "jpeg") {
      return "image/jpeg";
    }
    if (extension === "gif") {
      return "image/gif";
    }
    return "image/png";
  }
  
  function dataUrlToBytes(dataUrl) {
    const match = String(dataUrl || "").match(/^data:[^;]+;base64,(.+)$/);
    if (!match) {
      return new Uint8Array();
    }
  
    return base64ToBytes(match[1]);
  }
  
  function base64ToBytes(base64) {
    if (typeof root.atob === "function") {
      const binary = root.atob(base64);
      const bytes = new Uint8Array(binary.length);
      for (let index = 0; index < binary.length; index += 1) {
        bytes[index] = binary.charCodeAt(index);
      }
      return bytes;
    }
  
    if (typeof Buffer !== "undefined") {
      return new Uint8Array(Buffer.from(base64, "base64"));
    }
  
    return new Uint8Array();
  }

  async function createXlsxBlob(fields, rows, documentName, depositAmount, calculationRules, disabledAutoCalculationTargets) {
    const normalizedFields = normalizeFields(fields).filter((field) => field.label.trim());
    const normalizedRows = Array.isArray(rows) ? rows.map((row) => normalizeAppRow(row, normalizedFields)) : [];
    const normalizedRules = normalizeCalculationRules(calculationRules, normalizedFields);
    const normalizedDisabledTargets = normalizeDisabledAutoCalculationTargets(disabledAutoCalculationTargets, normalizedFields);
    const images = collectWorksheetImages(normalizedFields, normalizedRows);
    const xmlFiles = buildWorkbookFiles(
      normalizedFields,
      normalizedRows,
      images,
      documentName,
      depositAmount,
      normalizedRules,
      normalizedDisabledTargets
    );
    const zipBytes = createZip(xmlFiles);
    return new Blob([zipBytes], { type: MIME_XLSX });
  }
  
  function buildWorkbookFiles(fields, rows, images, documentName, depositAmount, calculationRules, disabledAutoCalculationTargets) {
    const worksheetImages = Array.isArray(images) ? images : collectWorksheetImages(fields, rows);
    const sheetXml = buildSheetXml(
      fields,
      rows,
      worksheetImages,
      documentName,
      depositAmount,
      calculationRules,
      disabledAutoCalculationTargets
    );
    const contentTypes = buildContentTypesXml(worksheetImages);
    const worksheetRelationships = worksheetImages.length ? buildWorksheetRelsXml() : null;
    const files = [
      {
        name: "[Content_Types].xml",
        content: contentTypes,
      },
      {
        name: "_rels/.rels",
        content: xmlDecl(`\
  <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>\
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>\
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>\
  </Relationships>`),
      },
      {
        name: "docProps/core.xml",
        content: xmlDecl(`\
  <cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">\
  <dc:title>自动生成表格</dc:title>\
  <dc:creator>Excel 自动生成工具</dc:creator>\
  <cp:lastModifiedBy>Excel 自动生成工具</cp:lastModifiedBy>\
  <dcterms:created xsi:type="dcterms:W3CDTF">${new Date().toISOString()}</dcterms:created>\
  <dcterms:modified xsi:type="dcterms:W3CDTF">${new Date().toISOString()}</dcterms:modified>\
  </cp:coreProperties>`),
      },
      {
        name: "docProps/app.xml",
        content: xmlDecl(`\
  <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">\
  <Application>Excel 自动生成工具</Application>\
  <DocSecurity>0</DocSecurity>\
  <ScaleCrop>false</ScaleCrop>\
  <HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant></vt:vector></HeadingPairs>\
  <TitlesOfParts><vt:vector size="1" baseType="lpstr"><vt:lpstr>数据</vt:lpstr></vt:vector></TitlesOfParts>\
  </Properties>`),
      },
      {
        name: "xl/workbook.xml",
        content: xmlDecl(`\
  <workbook xmlns="${XML_NS_MAIN}" xmlns:r="${XML_NS_REL}">\
  <fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="23426"/>\
  <workbookPr defaultThemeVersion="166925"/>\
  <bookViews><workbookView xWindow="0" yWindow="0" windowWidth="16000" windowHeight="9000"/></bookViews>\
  <sheets><sheet name="数据" sheetId="1" r:id="rId1"/></sheets>\
  <calcPr calcId="191029" calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/>\
  </workbook>`),
      },
      {
        name: "xl/_rels/workbook.xml.rels",
        content: xmlDecl(`\
  <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>\
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>\
  </Relationships>`),
      },
      { name: "xl/styles.xml", content: buildStylesXml() },
      { name: "xl/worksheets/sheet1.xml", content: sheetXml },
    ];
    if (worksheetRelationships) {
      files.push({
        name: "xl/worksheets/_rels/sheet1.xml.rels",
        content: worksheetRelationships,
      });
      files.push({
        name: "xl/drawings/drawing1.xml",
        content: buildDrawingXml(worksheetImages),
      });
      files.push({
        name: "xl/drawings/_rels/drawing1.xml.rels",
        content: buildDrawingRelsXml(worksheetImages),
      });
      worksheetImages.forEach((image) => {
        files.push({
          name: `xl/media/image${image.index}.${image.extension}`,
          content: image.bytes,
        });
      });
    }
    return files;
  }
  
  function buildContentTypesXml(images) {
    const defaults = new Set((images || []).map((image) => image.extension));
    const imageDefaults = Array.from(defaults)
      .map((extension) => `<Default Extension="${extension}" ContentType="${imageContentType(extension)}"/>`)
      .join("");
    const drawingOverride = images.length
      ? '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'
      : "";
  
    return xmlDecl(`\
  <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>\
  <Default Extension="xml" ContentType="application/xml"/>\
  ${imageDefaults}\
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>\
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>\
  ${drawingOverride}\
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>\
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>\
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>\
  </Types>`);
  }
  
  function buildWorksheetRelsXml() {
    return xmlDecl(`\
  <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>\
  </Relationships>`);
  }
  
  function buildDrawingRelsXml(images) {
    const relationships = images
      .map((image) => `<Relationship Id="rId${image.index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image${image.index}.${image.extension}"/>`)
      .join("");
  
    return xmlDecl(`\
  <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\
  ${relationships}\
  </Relationships>`);
  }
  
  function buildDrawingXml(images) {
    const anchors = images.map((image) => buildImageAnchorXml(image)).join("");
    return xmlDecl(`\
  <xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${XML_NS_REL}">\
  ${anchors}\
  </xdr:wsDr>`);
  }
  
  function buildImageAnchorXml(image) {
    const widthEmu = Math.round((IMAGE_CELL_WIDTH_PX - IMAGE_CELL_PADDING_PX * 2) * EMU_PER_PIXEL);
    const heightEmu = Math.round((IMAGE_CELL_HEIGHT_PX - IMAGE_CELL_PADDING_PX * 2) * EMU_PER_PIXEL);
    const offsetEmu = Math.round(IMAGE_CELL_PADDING_PX * EMU_PER_PIXEL);
  
    return `\
  <xdr:oneCellAnchor editAs="oneCell">\
  <xdr:from><xdr:col>${image.colIndex}</xdr:col><xdr:colOff>${offsetEmu}</xdr:colOff><xdr:row>${image.rowIndex}</xdr:row><xdr:rowOff>${offsetEmu}</xdr:rowOff></xdr:from>\
  <xdr:ext cx="${widthEmu}" cy="${heightEmu}"/>\
  <xdr:pic>\
  <xdr:nvPicPr><xdr:cNvPr id="${image.index}" name="${escapeXml(image.name || `图片 ${image.index}`)}"/><xdr:cNvPicPr/></xdr:nvPicPr>\
  <xdr:blipFill><a:blip r:embed="rId${image.index}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>\
  <xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${widthEmu}" cy="${heightEmu}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>\
  </xdr:pic>\
  <xdr:clientData/>\
  </xdr:oneCellAnchor>`;
  }
  
  function buildSheetXml(fields, rows, images, documentName, depositAmount, calculationRules, disabledAutoCalculationTargets) {
    const calculatedFields = getCalculatedFields(fields, calculationRules, disabledAutoCalculationTargets);
    const worksheetImages = Array.isArray(images) ? images : [];
    const sharedRemarkField = getSharedRemarkField(fields);
    const sharedRemarkIndex = sharedRemarkField ? fields.indexOf(sharedRemarkField) : -1;
    const customRowExportMap = getCustomRowExportMap(fields);
    const columnCount = Math.max(fields.length, 1);
    const titleRows = 2;
    const headerRows = 2;
    const dataStartRow = titleRows + headerRows + 1;
    const summaryRowNumber = rows.length + dataStartRow;
    const noticeRowNumber = summaryRowNumber + 1;
    const rowCount = Math.max(noticeRowNumber, dataStartRow - 1);
    const lastCell = `${columnName(columnCount)}${rowCount}`;
    const titleInfo = buildTitleRows(documentName, columnCount);
    const headerInfo = buildHeaderRows(fields, titleRows + 1);
    const summaryInfo = buildSummaryRow(fields, rows, dataStartRow, summaryRowNumber, depositAmount);
    const noticeInfo = buildNoticeRow(noticeRowNumber, columnCount);
    const customMergeRefs = [];
    const firstFieldMergeInfo = buildFirstFieldMergeInfo(fields, rows, dataStartRow);
    const dataRows = rows
      .map((row, rowIndex) => {
        const rowNumber = rowIndex + dataStartRow;
        const values = getRowValues(row);
        const custom = isCustomRow(row);
        const customRemarkMerge = custom && row.mergeRemark !== false && sharedRemarkIndex >= 0 && sharedRemarkIndex < columnCount - 1;
        const cells = fields
          .map((field, fieldIndex) => {
            const colNumber = fieldIndex + 1;
            if (custom) {
              const customCell = makeCustomRowExportCell(rowNumber, colNumber, field, fieldIndex, values, customRowExportMap);
              if (customCell) {
                return customCell;
              }
              return makeTextCell(rowNumber, colNumber, "", CUSTOM_TEXT_STYLE_ID);
            }
  
            const firstFieldMergeState = fieldIndex === 0 ? firstFieldMergeInfo.rowStates.get(rowNumber) : "";
            if (firstFieldMergeState === "tail") {
              return makeTextCell(rowNumber, colNumber, "", CENTER_TEXT_STYLE_ID);
            }
  
            const inMergedRemarkTail = customRemarkMerge && fieldIndex > sharedRemarkIndex;
            if (inMergedRemarkTail) {
              return makeTextCell(rowNumber, colNumber, "", 2);
            }
  
            const calculatedField = calculatedFields.get(field.key);
            if (calculatedField) {
              return makeFormulaCell(rowNumber, colNumber, calculatedField, CALCULATED_STYLE_ID);
            }
            const value = sharedRemarkField && field.key === sharedRemarkField.key && rowIndex > 0 && !custom
              ? ""
              : values[field.key];
            return firstFieldMergeState === "top"
              ? makeCenteredDataCell(rowNumber, colNumber, field, value)
              : makeDataCell(rowNumber, colNumber, field, value, custom);
          })
          .join("");
        if (customRemarkMerge) {
          customMergeRefs.push(`${cellRef(rowNumber, sharedRemarkIndex + 1)}:${cellRef(rowNumber, columnCount)}`);
        }
        const rowHeight = rowHasImage(worksheetImages, rowNumber) ? 78 : custom ? 42 : 0;
        const customHeight = rowHeight ? ` ht="${rowHeight}" customHeight="1"` : "";
        return `<row r="${rowNumber}"${customHeight}>${cells}</row>`;
      })
      .join("");
    const cols = fields
      .map((field, index) => {
        const width = field.type === "image"
          ? 19
          : Math.min(Math.max(String(field.label || "").length * 2 + 8, 14), 32);
        return `<col min="${index + 1}" max="${index + 1}" width="${width}" customWidth="1"/>`;
      })
      .join("");
    const drawing = worksheetImages.length ? '<drawing r:id="rId1"/>' : "";
    const sharedRemarkMergeRefs = buildSharedRemarkMergeRefs(rows, dataStartRow, sharedRemarkIndex);
  
    return xmlDecl(`\
  <worksheet xmlns="${XML_NS_MAIN}" xmlns:r="${XML_NS_REL}">\
  <dimension ref="A1:${lastCell}"/>\
  <sheetViews><sheetView workbookViewId="0"><pane ySplit="4" topLeftCell="A5" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft"/></sheetView></sheetViews>\
  <sheetFormatPr defaultRowHeight="18"/>\
  <cols>${cols}</cols>\
  <sheetData>${titleInfo.rows}${headerInfo.rows}${dataRows}${summaryInfo.rows}${noticeInfo.rows}</sheetData>\
  ${mergeBlocks(titleInfo.mergeRefs.concat(headerInfo.mergeRefs, firstFieldMergeInfo.mergeRefs, sharedRemarkMergeRefs, customMergeRefs, summaryInfo.mergeRefs, noticeInfo.mergeRefs))}\
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>\
  ${drawing}\
  </worksheet>`);
  }
  
  function buildFirstFieldMergeInfo(fields, rows, dataStartRow) {
    const mergeRefs = [];
    const rowStates = new Map();
    const firstField = fields[0];
    if (!firstField || rows.length < 2) {
      return { mergeRefs, rowStates };
    }
  
    let activeValue = "";
    let activeRows = [];
    const flushActiveRows = () => {
      if (activeRows.length > 1) {
        const firstRow = activeRows[0];
        const lastRow = activeRows[activeRows.length - 1];
        mergeRefs.push(`${cellRef(firstRow, 1)}:${cellRef(lastRow, 1)}`);
        activeRows.forEach((rowNumber, index) => {
          rowStates.set(rowNumber, index === 0 ? "top" : "tail");
        });
      }
      activeValue = "";
      activeRows = [];
    };
  
    rows.forEach((row, index) => {
      const rowNumber = dataStartRow + index;
      if (isCustomRow(row)) {
        flushActiveRows();
        return;
      }
  
      const value = getFirstFieldMergeValue(firstField, getRowValues(row));
      if (!value) {
        flushActiveRows();
        return;
      }
  
      if (activeRows.length && value !== activeValue) {
        flushActiveRows();
      }
  
      activeValue = value;
      activeRows.push(rowNumber);
    });
    flushActiveRows();
  
    return { mergeRefs, rowStates };
  }
  
  function getFirstFieldMergeValue(field, values) {
    const rawValue = values && values[field.key] != null ? String(values[field.key]) : "";
    if (field.type === "number") {
      return normalizeNonNegativeNumberValue(rawValue);
    }
    return rawValue.trim();
  }
  
  function buildSharedRemarkMergeRefs(rows, dataStartRow, sharedRemarkIndex) {
    if (sharedRemarkIndex < 0 || rows.length < 2) {
      return [];
    }
  
    const refs = [];
    let startRow = null;
    let endRow = null;
    rows.forEach((row, index) => {
      const rowNumber = dataStartRow + index;
      if (isCustomRow(row)) {
        if (startRow != null && endRow > startRow) {
          refs.push(`${cellRef(startRow, sharedRemarkIndex + 1)}:${cellRef(endRow, sharedRemarkIndex + 1)}`);
        }
        startRow = null;
        endRow = null;
        return;
      }
  
      if (startRow == null) {
        startRow = rowNumber;
      }
      endRow = rowNumber;
    });
  
    if (startRow != null && endRow > startRow) {
      refs.push(`${cellRef(startRow, sharedRemarkIndex + 1)}:${cellRef(endRow, sharedRemarkIndex + 1)}`);
    }
  
    return refs;
  }
  
  function buildNoticeRow(rowNumber, columnCount) {
    const cells = [makeTextCell(rowNumber, 1, EXPORT_NOTICE_TEXT, 2)];
    for (let colNumber = 2; colNumber <= columnCount; colNumber += 1) {
      cells.push(makeTextCell(rowNumber, colNumber, "", 2));
    }
  
    return {
      rows: `<row r="${rowNumber}" ht="42" customHeight="1">${cells.join("")}</row>`,
      mergeRefs: columnCount > 1 ? [`A${rowNumber}:${columnName(columnCount)}${rowNumber}`] : [],
    };
  }
  
  function buildSummaryRow(fields, rows, dataStartRow, rowNumber, depositAmount) {
    const columnCount = Math.max(fields.length, 1);
    const amountIndexes = fields
      .map((field, index) => isAmountField(field) ? index : -1)
      .filter((index) => index >= 0);
    const depositValue = normalizeNonNegativeNumberValue(depositAmount || "");
    const amountIndex = amountIndexes[0] != null ? amountIndexes[0] : -1;
    const depositLabelIndex = findDepositLabelIndex(columnCount, amountIndex);
    const depositValueIndex = depositLabelIndex >= 0 && depositLabelIndex + 1 < columnCount ? depositLabelIndex + 1 : -1;
    const totalFormula = makeTotalAmountFormula(amountIndexes, rows.length, dataStartRow);
    const cells = [];
  
    for (let fieldIndex = 0; fieldIndex < columnCount; fieldIndex += 1) {
      const colNumber = fieldIndex + 1;
      if (fieldIndex === 0) {
        cells.push(makeTextCell(rowNumber, colNumber, "总金额", CUSTOM_TEXT_STYLE_ID));
      } else if (fieldIndex === amountIndex) {
        cells.push(totalFormula
          ? `<c r="${cellRef(rowNumber, colNumber)}" s="${CUSTOM_CALCULATED_STYLE_ID}"><f>${escapeXml(totalFormula)}</f></c>`
          : makeDataCell(rowNumber, colNumber, { type: "number" }, "", true));
      } else if (fieldIndex === depositLabelIndex) {
        cells.push(makeTextCell(rowNumber, colNumber, "定金", CUSTOM_TEXT_STYLE_ID));
      } else if (fieldIndex === depositValueIndex) {
        cells.push(makeDataCell(rowNumber, colNumber, { type: "number" }, depositValue, true));
      } else {
        cells.push(makeTextCell(rowNumber, colNumber, "", CUSTOM_TEXT_STYLE_ID));
      }
    }
  
    return {
      rows: `<row r="${rowNumber}" ht="28" customHeight="1">${cells.join("")}</row>`,
      mergeRefs: [],
    };
  }
  
  function findDepositLabelIndex(columnCount, amountIndex) {
    if (columnCount < 2) {
      return -1;
    }
  
    const candidates = amountIndex + 2 < columnCount
      ? [amountIndex + 1, columnCount - 2, 1]
      : [columnCount - 2, 1];
    return candidates.find((index) => index >= 0 && index + 1 < columnCount && index !== amountIndex && index + 1 !== amountIndex) ?? -1;
  }
  
  function makeTotalAmountFormula(amountIndexes, dataRowCount, dataStartRow) {
    if (!amountIndexes.length || !dataRowCount) {
      return "";
    }
  
    const dataEndRow = dataStartRow + dataRowCount - 1;
    return amountIndexes
      .map((fieldIndex) => {
        const colName = columnName(fieldIndex + 1);
        return `SUM(${colName}${dataStartRow}:${colName}${dataEndRow})`;
      })
      .join("+");
  }
  
  function buildTitleRows(documentName, columnCount) {
    const lastColumn = columnName(columnCount);
    const labelCells = [makeTextCell(1, 1, "名称", 6)];
    const valueCells = [makeTextCell(2, 1, String(documentName || ""), 7)];
    for (let colNumber = 2; colNumber <= columnCount; colNumber += 1) {
      labelCells.push(makeTextCell(1, colNumber, "", 6));
      valueCells.push(makeTextCell(2, colNumber, "", 7));
    }
    return {
      rows: `<row r="1" ht="24" customHeight="1">${labelCells.join("")}</row><row r="2" ht="26" customHeight="1">${valueCells.join("")}</row>`,
      mergeRefs: [`A1:${lastColumn}1`, `A2:${lastColumn}2`],
    };
  }
  
  function buildHeaderRows(fields, startRow) {
    const topCells = [];
    const secondCells = [];
    const mergeRefs = [];
    const topRow = startRow || 1;
    const secondRow = topRow + 1;
    let index = 0;
  
    while (index < fields.length) {
      const field = fields[index];
      const group = String(field.group || "").trim();
      const colNumber = index + 1;
  
      if (!group) {
        topCells.push(makeTextCell(topRow, colNumber, field.label, 1));
        secondCells.push(makeTextCell(secondRow, colNumber, "", 1));
        mergeRefs.push(`${cellRef(topRow, colNumber)}:${cellRef(secondRow, colNumber)}`);
        index += 1;
        continue;
      }
  
      let endIndex = index;
      while (endIndex + 1 < fields.length && String(fields[endIndex + 1].group || "").trim() === group) {
        endIndex += 1;
      }
  
      topCells.push(makeTextCell(topRow, colNumber, group, 1));
      for (let blankIndex = index + 1; blankIndex <= endIndex; blankIndex += 1) {
        topCells.push(makeTextCell(topRow, blankIndex + 1, "", 1));
      }
      for (let childIndex = index; childIndex <= endIndex; childIndex += 1) {
        secondCells.push(makeTextCell(secondRow, childIndex + 1, fields[childIndex].label, 1));
      }
      if (endIndex > index) {
        mergeRefs.push(`${cellRef(topRow, colNumber)}:${cellRef(topRow, endIndex + 1)}`);
      }
      index = endIndex + 1;
    }
  
    const headerRows = `<row r="${topRow}" ht="24" customHeight="1">${topCells.join("")}</row><row r="${secondRow}" ht="24" customHeight="1">${secondCells.join("")}</row>`;
    return { rows: headerRows, mergeRefs };
  }
  
  function mergeBlocks(mergeRefs) {
    return mergeRefs.length
      ? `<mergeCells count="${mergeRefs.length}">${mergeRefs.map((ref) => `<mergeCell ref="${ref}"/>`).join("")}</mergeCells>`
      : "";
  }
  
  function getCustomRowExportMap(fields) {
    const sourceFields = fields || [];
    const amountField = findCustomRowAmountField(sourceFields);
    const amountIndex = amountField ? sourceFields.indexOf(amountField) : -1;
    const beforeAmountFields = amountIndex > -1 ? sourceFields.slice(0, amountIndex) : sourceFields;
    const amountGroup = amountField && normalizeLabel(amountField.group);
    const groupFields = amountGroup ? beforeAmountFields.filter((field) => sameGroup(field, amountField)) : beforeAmountFields;
    const scopedFields = groupFields.length ? groupFields : beforeAmountFields;
  
    const quantityField = findLast(scopedFields, isQuantityField) || findLast(beforeAmountFields, isQuantityField);
    const unitPriceField = findLast(scopedFields, (field) => isUnitPriceField(field) && field !== amountField) ||
      findLast(beforeAmountFields, (field) => isUnitPriceField(field) && field !== amountField);
    const typeField = findLast(scopedFields, isCustomTypeFallbackField) || findLast(beforeAmountFields, isCustomTypeFallbackField) || null;
  
    return {
      fields: sourceFields,
      typeField,
      typeIndex: typeField ? sourceFields.indexOf(typeField) : -1,
      quantityField: quantityField || null,
      quantityIndex: quantityField ? sourceFields.indexOf(quantityField) : -1,
      unitPriceField: unitPriceField || null,
      unitPriceIndex: unitPriceField ? sourceFields.indexOf(unitPriceField) : -1,
      amountField,
      amountIndex,
    };
  }
  
  function findCustomRowAmountField(fields) {
    const sourceFields = fields || [];
    const amountFields = sourceFields.filter(isAmountField);
    return amountFields.find((field) => {
      const index = sourceFields.indexOf(field);
      const previousFields = sourceFields.slice(0, index);
      const sameGroupFields = normalizeLabel(field.group)
        ? previousFields.filter((candidate) => sameGroup(candidate, field))
        : previousFields;
      const scopedFields = sameGroupFields.length ? sameGroupFields : previousFields;
      return findLast(scopedFields, isQuantityField) &&
        findLast(scopedFields, (candidate) => isUnitPriceField(candidate) && candidate !== field);
    }) || amountFields[0] || null;
  }
  
  function makeCustomRowExportCell(rowNumber, colNumber, field, fieldIndex, values, exportMap) {
    if (exportMap.typeField && field.key === exportMap.typeField.key) {
      return makeDataCell(rowNumber, colNumber, CUSTOM_ROW_FIELDS[0], values[CUSTOM_ROW_TYPE_KEY], true);
    }
  
    if (exportMap.quantityField && field.key === exportMap.quantityField.key) {
      return makeDataCell(rowNumber, colNumber, CUSTOM_ROW_FIELDS[1], values[CUSTOM_ROW_QUANTITY_KEY], true);
    }
  
    if (exportMap.unitPriceField && field.key === exportMap.unitPriceField.key) {
      return makeDataCell(rowNumber, colNumber, CUSTOM_ROW_FIELDS[2], values[CUSTOM_ROW_UNIT_PRICE_KEY], true);
    }
  
    if (exportMap.amountField && field.key === exportMap.amountField.key) {
      if (exportMap.quantityIndex > -1 && exportMap.unitPriceIndex > -1) {
        return makeFormulaCell(
          rowNumber,
          colNumber,
          {
            type: "quantityAmount",
            operator: "multiply",
            sourceIndexes: [exportMap.quantityIndex, exportMap.unitPriceIndex],
          },
          CUSTOM_CALCULATED_STYLE_ID
        );
      }
      return makeDataCell(rowNumber, colNumber, CUSTOM_ROW_FIELDS[3], values[CUSTOM_ROW_AMOUNT_KEY], true);
    }
  
    return "";
  }
  
  function buildStylesXml() {
    return xmlDecl(`\
  <styleSheet xmlns="${XML_NS_MAIN}">\
  <numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/></numFmts>\
  <fonts count="4">\
  <font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/></font>\
  <font><b/><sz val="11"/><color rgb="FF000000"/><name val="Calibri"/><family val="2"/></font>\
  <font><b/><sz val="14"/><color rgb="FF000000"/><name val="Calibri"/><family val="2"/></font>\
  <font><b/><sz val="12"/><color rgb="FF000000"/><name val="Calibri"/><family val="2"/></font>\
  </fonts>\
  <fills count="4">\
  <fill><patternFill patternType="none"/></fill>\
  <fill><patternFill patternType="gray125"/></fill>\
  <fill><patternFill patternType="solid"><fgColor rgb="FFC0C0C0"/><bgColor indexed="64"/></patternFill></fill>\
  <fill><patternFill patternType="solid"><fgColor rgb="FFE7EAEE"/><bgColor indexed="64"/></patternFill></fill>\
  </fills>\
  <borders count="2">\
  <border><left/><right/><top/><bottom/><diagonal/></border>\
  <border><left style="thin"><color rgb="FFD8E0DE"/></left><right style="thin"><color rgb="FFD8E0DE"/></right><top style="thin"><color rgb="FFD8E0DE"/></top><bottom style="thin"><color rgb="FFD8E0DE"/></bottom><diagonal/></border>\
  </borders>\
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>\
  <cellXfs count="14">\
  <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>\
  <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>\
  <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>\
  <xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>\
  <xf numFmtId="2" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>\
  <xf numFmtId="2" fontId="0" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>\
  <xf numFmtId="0" fontId="2" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>\
  <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>\
  <xf numFmtId="0" fontId="3" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>\
  <xf numFmtId="164" fontId="3" fillId="0" borderId="1" xfId="0" applyFont="1" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>\
  <xf numFmtId="2" fontId="3" fillId="0" borderId="1" xfId="0" applyFont="1" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>\
  <xf numFmtId="2" fontId="3" fillId="3" borderId="1" xfId="0" applyFont="1" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>\
  <xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>\
  <xf numFmtId="2" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>\
  </cellXfs>\
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>\
  <dxfs count="0"/>\
  <tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>\
  </styleSheet>`);
  }
  
  function makeDataCell(rowNumber, colNumber, field, rawValue, bold) {
    const rawText = rawValue == null ? "" : String(rawValue);
    const value = field.type === "number" ? normalizeNonNegativeNumberValue(rawText) : rawText;
    const textStyleId = bold ? CUSTOM_TEXT_STYLE_ID : 2;
    const dateStyleId = bold ? CUSTOM_DATE_STYLE_ID : 3;
    const numberStyleId = bold ? CUSTOM_NUMBER_STYLE_ID : 4;
    if (field.type === "image") {
      return makeTextCell(rowNumber, colNumber, "", textStyleId);
    }
  
    if (value === "") {
      return makeTextCell(rowNumber, colNumber, "", textStyleId);
    }
  
    if (field.type === "number") {
      const numericValue = Number(value);
      if (Number.isFinite(numericValue)) {
        return `<c r="${cellRef(rowNumber, colNumber)}" s="${numberStyleId}"><v>${numericValue}</v></c>`;
      }
    }
  
    if (field.type === "date" && /^\d{4}-\d{2}-\d{2}$/.test(value)) {
      return `<c r="${cellRef(rowNumber, colNumber)}" s="${dateStyleId}"><v>${dateToExcelSerial(value)}</v></c>`;
    }
  
    return makeTextCell(rowNumber, colNumber, value, textStyleId);
  }
  
  function makeCenteredDataCell(rowNumber, colNumber, field, rawValue) {
    const rawText = rawValue == null ? "" : String(rawValue);
    const value = field.type === "number" ? normalizeNonNegativeNumberValue(rawText) : rawText;
    if (field.type === "image" || value === "") {
      return makeTextCell(rowNumber, colNumber, "", CENTER_TEXT_STYLE_ID);
    }
  
    if (field.type === "number") {
      const numericValue = Number(value);
      if (Number.isFinite(numericValue)) {
        return `<c r="${cellRef(rowNumber, colNumber)}" s="${CENTER_NUMBER_STYLE_ID}"><v>${numericValue}</v></c>`;
      }
    }
  
    if (field.type === "date" && /^\d{4}-\d{2}-\d{2}$/.test(value)) {
      return `<c r="${cellRef(rowNumber, colNumber)}" s="${CENTER_DATE_STYLE_ID}"><v>${dateToExcelSerial(value)}</v></c>`;
    }
  
    return makeTextCell(rowNumber, colNumber, value, CENTER_TEXT_STYLE_ID);
  }
  
  function makeFormulaCell(rowNumber, colNumber, calculatedField, styleId) {
    const [firstSourceIndex, secondSourceIndex] = calculatedField.sourceIndexes;
    const firstRef = cellRef(rowNumber, firstSourceIndex + 1);
    const secondRef = cellRef(rowNumber, secondSourceIndex + 1);
    const operator = calculatedField.operator || CALCULATION_TYPE_OPERATORS[calculatedField.type] || "multiply";
    const symbol = getFormulaOperator(operator);
    const emptyCheck = operator === "divide"
      ? `OR(${firstRef}="",${secondRef}="",${secondRef}=0)`
      : `OR(${firstRef}="",${secondRef}="")`;
    const formula = `IF(${emptyCheck},"",${firstRef}${symbol}${secondRef})`;
    return `<c r="${cellRef(rowNumber, colNumber)}" s="${styleId || CALCULATED_STYLE_ID}"><f>${escapeXml(formula)}</f></c>`;
  }
  
  function getFormulaOperator(operator) {
    if (operator === "add") {
      return "+";
    }
    if (operator === "subtract") {
      return "-";
    }
    if (operator === "divide") {
      return "/";
    }
    return "*";
  }
  
  function makeTextCell(rowNumber, colNumber, value, styleId) {
    return `<c r="${cellRef(rowNumber, colNumber)}" t="inlineStr" s="${styleId}"><is><t${needsPreserveSpace(value) ? ' xml:space="preserve"' : ""}>${escapeXml(value)}</t></is></c>`;
  }
  
  function dateToExcelSerial(dateValue) {
    const [year, month, day] = dateValue.split("-").map(Number);
    const utcDate = Date.UTC(year, month - 1, day);
    const epoch = Date.UTC(1899, 11, 30);
    return Math.round((utcDate - epoch) / 86400000);
  }
  
  function formatTimestamp(date) {
    const pad = (value) => String(value).padStart(2, "0");
    return `${date.getFullYear()}${pad(date.getMonth() + 1)}${pad(date.getDate())}_${pad(date.getHours())}${pad(date.getMinutes())}`;
  }
  
  function xmlDecl(xml) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${xml}`;
  }
  
  function escapeXml(value) {
    return String(value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&apos;");
  }
  
  function escapeHtml(value) {
    return escapeXml(value);
  }
  
  function escapeAttr(value) {
    return escapeXml(value);
  }
  
  function needsPreserveSpace(value) {
    return /^\s|\s$/.test(String(value));
  }
  
  function cellRef(rowNumber, colNumber) {
    return `${columnName(colNumber)}${rowNumber}`;
  }
  
  function columnName(colNumber) {
    let name = "";
    let number = colNumber;
    while (number > 0) {
      const remainder = (number - 1) % 26;
      name = String.fromCharCode(65 + remainder) + name;
      number = Math.floor((number - 1) / 26);
    }
    return name;
  }
  
  function createZip(files) {
    const encoder = typeof TextEncoder === "function" ? new TextEncoder() : null;
    const localParts = [];
    const centralParts = [];
    let offset = 0;
  
    files.forEach((file) => {
      const nameBytes = encodeUtf8(file.name, encoder);
      const dataBytes = typeof file.content === "string" ? encodeUtf8(file.content, encoder) : file.content;
      const crc = crc32(dataBytes);
      const timestamp = dosTimestamp(new Date());
      const localHeader = new Uint8Array(30 + nameBytes.length);
      const localView = new DataView(localHeader.buffer);
      localView.setUint32(0, 0x04034b50, true);
      localView.setUint16(4, 20, true);
      localView.setUint16(6, 0x0800, true);
      localView.setUint16(8, 0, true);
      localView.setUint16(10, timestamp.time, true);
      localView.setUint16(12, timestamp.date, true);
      localView.setUint32(14, crc, true);
      localView.setUint32(18, dataBytes.length, true);
      localView.setUint32(22, dataBytes.length, true);
      localView.setUint16(26, nameBytes.length, true);
      localView.setUint16(28, 0, true);
      localHeader.set(nameBytes, 30);
      localParts.push(localHeader, dataBytes);
  
      const centralHeader = new Uint8Array(46 + nameBytes.length);
      const centralView = new DataView(centralHeader.buffer);
      centralView.setUint32(0, 0x02014b50, true);
      centralView.setUint16(4, 0x0314, true);
      centralView.setUint16(6, 20, true);
      centralView.setUint16(8, 0x0800, true);
      centralView.setUint16(10, 0, true);
      centralView.setUint16(12, timestamp.time, true);
      centralView.setUint16(14, timestamp.date, true);
      centralView.setUint32(16, crc, true);
      centralView.setUint32(20, dataBytes.length, true);
      centralView.setUint32(24, dataBytes.length, true);
      centralView.setUint16(28, nameBytes.length, true);
      centralView.setUint16(30, 0, true);
      centralView.setUint16(32, 0, true);
      centralView.setUint16(34, 0, true);
      centralView.setUint16(36, 0, true);
      centralView.setUint32(38, 0, true);
      centralView.setUint32(42, offset, true);
      centralHeader.set(nameBytes, 46);
      centralParts.push(centralHeader);
  
      offset += localHeader.length + dataBytes.length;
    });
  
    const centralOffset = offset;
    const centralSize = centralParts.reduce((sum, part) => sum + part.length, 0);
    const endRecord = new Uint8Array(22);
    const endView = new DataView(endRecord.buffer);
    endView.setUint32(0, 0x06054b50, true);
    endView.setUint16(4, 0, true);
    endView.setUint16(6, 0, true);
    endView.setUint16(8, files.length, true);
    endView.setUint16(10, files.length, true);
    endView.setUint32(12, centralSize, true);
    endView.setUint32(16, centralOffset, true);
    endView.setUint16(20, 0, true);
  
    const allParts = localParts.concat(centralParts, [endRecord]);
    const totalLength = allParts.reduce((sum, part) => sum + part.length, 0);
    const output = new Uint8Array(totalLength);
    let cursor = 0;
    allParts.forEach((part) => {
      output.set(part, cursor);
      cursor += part.length;
    });
    return output;
  }
  
  function encodeUtf8(value, encoder) {
    if (encoder) {
      return encoder.encode(value);
    }
  
    const text = unescape(encodeURIComponent(String(value)));
    const bytes = new Uint8Array(text.length);
    for (let index = 0; index < text.length; index += 1) {
      bytes[index] = text.charCodeAt(index);
    }
    return bytes;
  }
  
  function dosTimestamp(date) {
    const safeDate = date < ZIP_EPOCH ? ZIP_EPOCH : date;
    return {
      time: (safeDate.getHours() << 11) | (safeDate.getMinutes() << 5) | Math.floor(safeDate.getSeconds() / 2),
      date: ((safeDate.getFullYear() - 1980) << 9) | ((safeDate.getMonth() + 1) << 5) | safeDate.getDate(),
    };
  }
  
  const CRC_TABLE = (() => {
    const table = new Uint32Array(256);
    for (let i = 0; i < 256; i += 1) {
      let c = i;
      for (let k = 0; k < 8; k += 1) {
        c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      }
      table[i] = c >>> 0;
    }
    return table;
  })();
  
  function crc32(bytes) {
    let crc = 0xffffffff;
    for (let i = 0; i < bytes.length; i += 1) {
      crc = CRC_TABLE[(crc ^ bytes[i]) & 0xff] ^ (crc >>> 8);
    }
    return (crc ^ 0xffffffff) >>> 0;
  }

  return {
    buildWorkbookFiles,
    collectWorksheetImages,
    createXlsxBlob,
    dateToExcelSerial,
    columnName,
  };
}

export { createXlsxBuilder };
