const STORAGE_KEY = "excel-auto-tool-fields-v2";
const DRAFT_DB_NAME = "excel-auto-tool-drafts";
const DRAFT_DB_VERSION = 1;
const DRAFT_STORE_NAME = "drafts";
const DRAFT_VERSION = 1;
const MAX_DRAFTS = 2;
const DRAFT_SAVE_DELAY = 600;
const DEFAULT_FIELDS = [
  { key: "install_location", label: "安装位置", group: "", type: "text", required: false, options: [] },
  { key: "model", label: "型号", group: "", type: "text", required: false, options: [] },
  { key: "image", label: "图片", group: "", type: "image", required: false, options: [] },
  { key: "finished_width", label: "宽", group: "成品规格(M)", type: "number", required: false, options: [] },
  { key: "finished_height", label: "高", group: "成品规格(M)", type: "number", required: false, options: [] },
  { key: "finished_ratio", label: "比例", group: "成品规格(M)", type: "text", required: false, options: [] },
  { key: "finished_style", label: "款式", group: "成品规格(M)", type: "text", required: false, options: [] },
  { key: "material_meters", label: "米数", group: "用料", type: "number", required: false, options: [] },
  { key: "material_unit_price", label: "元/米", group: "单价", type: "number", required: false, options: [] },
  { key: "amount", label: "金额", group: "", type: "number", required: false, options: [] },
  { key: "track_name", label: "品名", group: "轨道", type: "text", required: false, options: [] },
  { key: "track_quantity", label: "数量", group: "轨道", type: "number", required: false, options: [] },
  { key: "track_unit_price", label: "单价", group: "轨道", type: "number", required: false, options: [] },
  { key: "track_amount", label: "金额", group: "轨道", type: "number", required: false, options: [] },
  { key: "remark", label: "备注", group: "", type: "text", required: false, options: [] },
];
const CUSTOM_ROW_TYPE_KEY = "custom_type";
const CUSTOM_ROW_QUANTITY_KEY = "custom_quantity";
const CUSTOM_ROW_UNIT_PRICE_KEY = "custom_unit_price";
const CUSTOM_ROW_AMOUNT_KEY = "custom_amount";
const CUSTOM_ROW_FIELDS = [
  { key: CUSTOM_ROW_TYPE_KEY, label: "类型", group: "", type: "text", required: false, options: [] },
  { key: CUSTOM_ROW_QUANTITY_KEY, label: "数量", group: "", type: "number", required: false, options: [] },
  { key: CUSTOM_ROW_UNIT_PRICE_KEY, label: "单价", group: "", type: "number", required: false, options: [] },
  { key: CUSTOM_ROW_AMOUNT_KEY, label: "金额", group: "", type: "number", required: false, options: [] },
];
const FIELD_TYPES = ["text", "number", "date", "select", "image"];
const TYPE_TEXT = {
  text: "文本",
  number: "数字",
  date: "日期",
  select: "下拉",
  image: "图片",
};
const XML_NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const XML_NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const ZIP_EPOCH = new Date("1980-01-01T00:00:00Z");
const SUPPORTED_IMAGE_MIMES = ["image/png", "image/jpeg", "image/gif"];
const IMAGE_CELL_WIDTH_PX = 132;
const IMAGE_CELL_HEIGHT_PX = 96;
const IMAGE_CELL_PADDING_PX = 6;
const EMU_PER_PIXEL = 9525;
const CALCULATED_STYLE_ID = 5;
const CENTER_TEXT_STYLE_ID = 7;
const CUSTOM_TEXT_STYLE_ID = 8;
const CUSTOM_DATE_STYLE_ID = 9;
const CUSTOM_NUMBER_STYLE_ID = 10;
const CUSTOM_CALCULATED_STYLE_ID = 11;
const CENTER_DATE_STYLE_ID = 12;
const CENTER_NUMBER_STYLE_ID = 13;
const CALCULATION_OPERATORS = ["add", "subtract", "multiply", "divide"];
const CALCULATION_TYPE_OPERATORS = {
  sum: "add",
  meterAmount: "multiply",
  quantityAmount: "multiply",
};
const EXPORT_NOTICE_TEXT = "\u6e29\u99a8\u63d0\u793a\uff1a1\uff1a\u7a97\u5e18\u4e3a\u5b9a\u5236\u4ea7\u54c1\uff0c\u82e5\u65e0\u8d28\u91cf\u95ee\u9898\u4e0d\u4e88\u9000\u6362\u30022\uff1a\u9762\u6599\u5899\u7eb8\u8272\u5dee\u5141\u8bb8\u5728\u56fd\u5bb6\u89c4\u5b9a\u8272\u5dee\u5141\u8bb85%\u4e4b\u5185\u3002\uff08\u5ba3\u745e\u8f6f\u88c5 \u9648\u59d0\uff1a13688428383\uff09";

export {
  CALCULATED_STYLE_ID,
  CALCULATION_OPERATORS,
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
  DEFAULT_FIELDS,
  DRAFT_DB_NAME,
  DRAFT_DB_VERSION,
  DRAFT_SAVE_DELAY,
  DRAFT_STORE_NAME,
  DRAFT_VERSION,
  EMU_PER_PIXEL,
  EXPORT_NOTICE_TEXT,
  FIELD_TYPES,
  IMAGE_CELL_HEIGHT_PX,
  IMAGE_CELL_PADDING_PX,
  IMAGE_CELL_WIDTH_PX,
  MAX_DRAFTS,
  STORAGE_KEY,
  SUPPORTED_IMAGE_MIMES,
  TYPE_TEXT,
  XML_NS_MAIN,
  XML_NS_REL,
  ZIP_EPOCH,
};
