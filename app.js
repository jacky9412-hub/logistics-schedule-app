const STORAGE_KEY = "logistics-schedule-app-state-v7";
const SESSION_AUTH_KEY = "logistics-schedule-auth-cache";

// ─── Firebase Configuration ───
const firebaseConfig = {
  apiKey: "AIzaSyBl8BuPyoaHAIoa4yQkkmAHSgvl1l3kZsw",
  authDomain: "logistics-schedule-2c6c5.firebaseapp.com",
  databaseURL: "https://logistics-schedule-2c6c5-default-rtdb.asia-southeast1.firebasedatabase.app",
  projectId: "logistics-schedule-2c6c5",
  storageBucket: "logistics-schedule-2c6c5.firebasestorage.app",
  messagingSenderId: "515233843932",
  appId: "1:515233843932:web:958248c6095362ac338219",
  measurementId: "G-KEWZ2SMJYW",
};

firebase.initializeApp(firebaseConfig);
const firebaseDb = firebase.database();
const stateRef = firebaseDb.ref("appState");
let firebaseReady = false;
let lastFirebaseSaveTime = 0;
const FIREBASE_ECHO_DELAY = 3000; // Ignore listener echoes within 3 seconds of our own save

// Firebase converts arrays to objects; this ensures they stay arrays
function ensureArrays(data) {
  if (!data) return data;
  const arrayFields = ["employees", "routes", "assignments", "auditLogs"];
  arrayFields.forEach((field) => {
    if (data[field] && !Array.isArray(data[field])) {
      data[field] = Object.values(data[field]);
    }
  });
  if (data.companySettings?.holidays && !Array.isArray(data.companySettings.holidays)) {
    data.companySettings.holidays = Object.values(data.companySettings.holidays);
  }
  if (data.employees && Array.isArray(data.employees)) {
    data.employees.forEach((emp) => {
      if (emp && emp.supportLineIds && !Array.isArray(emp.supportLineIds)) {
        emp.supportLineIds = Object.values(emp.supportLineIds);
      }
    });
  }
  return data;
}

let syncStatusTimer = null;
function showSyncStatus(type, message) {
  const el = document.getElementById("syncStatus");
  if (!el) return;
  if (syncStatusTimer) clearTimeout(syncStatusTimer);
  el.className = "sync-status " + type;
  el.textContent = message;
  el.style.display = "block";
  if (type === "synced") {
    syncStatusTimer = setTimeout(() => { el.style.display = "none"; }, 2000);
  } else if (type === "error") {
    syncStatusTimer = setTimeout(() => { el.style.display = "none"; }, 5000);
  }
}

function hideLoading() {
  const overlay = document.getElementById("loadingOverlay");
  if (overlay) {
    overlay.classList.add("fade-out");
    setTimeout(() => { overlay.remove(); }, 500);
  }
}

const taiwanHolidaySeeds2026 = [
  "2026-01-01",
  "2026-02-16",
  "2026-02-17",
  "2026-02-18",
  "2026-02-19",
  "2026-02-20",
  "2026-02-27",
  "2026-04-03",
  "2026-04-06",
  "2026-05-01",
  "2026-06-19",
  "2026-09-25",
  "2026-10-09",
];

function renderHolidaySeedList() {
  return taiwanHolidaySeeds2026
    .map((dateString) => `<span>${dateString}</span>`)
    .join("");
}

function renderCompanyHolidayList() {
  const holidays = [...state.companySettings.holidays].sort();
  if (!holidays.length) {
    return "<span>目前沒有公司休假日</span>";
  }
  return holidays
    .map((dateString) => {
      const isBuiltIn = taiwanHolidaySeeds2026.includes(dateString);
      return `<span class="${isBuiltIn ? "" : "brand"}">${dateString}${isBuiltIn ? "" : "｜手動"}</span>`;
    })
    .join("");
}

const roleLabels = {
  operator: "運務員",
  teamLeader: "組長",
  reliefStaff: "抵休",
  adminStaff: "行政",
  supervisor: "主管",
};

const fixedEmployeeOrderByName = new Map([
  ["洪主任", 1],
  ["宏擷", 2],
  ["凱強", 3],
  ["明怡", 4],
  ["志佳", 5],
  ["邦漢", 6],
  ["翔清", 7],
  ["倉熙", 8],
  ["長霖", 9],
  ["繼光", 10],
  ["智順", 11],
  ["琮緯", 12],
  ["佑謙", 13],
  ["宥冰", 14],
  ["家豪", 15],
  ["炎億", 16],
  ["鴻民", 17],
  ["裕庭", 18],
  ["懷慶", 19],
  ["庭愷", 20],
  ["建仁", 21],
  ["哲維", 22],
  ["玉女", 23],
  ["志剛", 24],
  ["良賢", 25],
  ["建凱", 26],
  ["家瑞", 27],
  ["逸文", 28],
  ["德豐", 29],
  ["育宗", 30],
  ["坤宗", 31],
  ["孟峰", 32],
  ["維遠", 33],
  ["珽安", 34],
  ["祥瑜", 35],
  ["書豪", 36],
  ["昱瑋", 37],
  ["新洋", 38],
  ["鴻昇", 39],
  ["彥鈞", 40],
  ["信佑", 41],
  ["明軒", 42],
  ["家弘", 43],
  ["獻斌", 44],
  ["崇傑", 45],
  ["凱瑋", 46],
  ["虹享", 47],
  ["政偉", 48],
  ["曉謙", 49],
  ["秋貴", 50],
  ["文聖", 51],
  ["世文", 52],
]);

const roleOrderForList = {
  supervisor: 1,
  teamLeader: 2,
  adminStaff: 3,
  reliefStaff: 4,
  operator: 5,
};

const employmentStatusLabels = {
  active: "在職",
  resigned: "已離職",
  unpaidLeave: "留職停薪",
};

const routeTypeLabels = {
  car: "汽車線",
  scooter: "機車段",
  special: "特殊勤務",
};

const routeSeeds = [
  { name: "大夜班", type: "special", approvedMileage: 6, owner: "長霖" },
  { name: "嘉義線", type: "car", approvedMileage: 279, owner: "繼光" },
  { name: "雲林線", type: "car", approvedMileage: 189, owner: "智順" },
  { name: "苗栗線", type: "car", approvedMileage: 187, owner: "琮緯" },
  { name: "員林線", type: "car", approvedMileage: 197, owner: "佑謙" },
  { name: "大甲線", type: "car", approvedMileage: 167, owner: "宥冰" },
  { name: "埔里線", type: "car", approvedMileage: 210, owner: "家豪" },
  { name: "鹿港線", type: "car", approvedMileage: 118, owner: "炎億" },
  { name: "作業中心線", type: "car", approvedMileage: 62, owner: "鴻民" },
  { name: "中區段(晚班)", type: "scooter", approvedMileage: 36, owner: "裕庭" },
  { name: "軍功段", type: "scooter", approvedMileage: 49, owner: "倉熙" },
  { name: "南屯段", type: "scooter", approvedMileage: 47, owner: "懷慶" },
  { name: "市政段", type: "scooter", approvedMileage: 53, owner: "庭愷" },
  { name: "黎明段", type: "scooter", approvedMileage: 57, owner: "建仁" },
  { name: "松竹段", type: "scooter", approvedMileage: 66, owner: "哲維" },
  { name: "中工段", type: "scooter", approvedMileage: 81, owner: "玉女" },
  { name: "中港段", type: "scooter", approvedMileage: 52, owner: "志剛" },
  { name: "文心段", type: "scooter", approvedMileage: 41, owner: "良賢" },
  { name: "大雅段", type: "scooter", approvedMileage: 66, owner: "建凱" },
  { name: "太平段", type: "scooter", approvedMileage: 45, owner: "家瑞" },
  { name: "大里段", type: "scooter", approvedMileage: 64, owner: "逸文" },
  { name: "東勢段", type: "scooter", approvedMileage: 94, owner: "德豐" },
  { name: "后里段", type: "scooter", approvedMileage: 88, owner: "育宗" },
  { name: "清水段", type: "scooter", approvedMileage: 77, owner: "坤宗" },
  { name: "沙鹿段", type: "scooter", approvedMileage: 70, owner: "孟峰" },
  { name: "梧棲段(半日)", type: "scooter", approvedMileage: 67, owner: "維遠" },
  { name: "玉山大雅(半日)", type: "scooter", approvedMileage: 43, owner: "珽安" },
  { name: "玉山南屯(晚班)", type: "scooter", approvedMileage: 30, owner: "祥瑜" },
  { name: "市政二段(晚班)", type: "scooter", approvedMileage: 38, owner: "書豪" },
  { name: "南區段(晚班)", type: "scooter", approvedMileage: 28, owner: "昱瑋" },
  { name: "彰化段", type: "scooter", approvedMileage: 103, owner: "新洋" },
  { name: "和美段", type: "scooter", approvedMileage: 97, owner: "鴻昇" },
  { name: "員林段", type: "scooter", approvedMileage: 50, owner: "彥鈞" },
  { name: "溪湖段", type: "scooter", approvedMileage: 44, owner: "信佑" },
  { name: "頭份段", type: "scooter", approvedMileage: 61, owner: "明軒" },
  { name: "竹南段", type: "scooter", approvedMileage: 61, owner: "家弘" },
  { name: "草屯段", type: "scooter", approvedMileage: 88, owner: "獻斌" },
  { name: "竹山段", type: "scooter", approvedMileage: 116, owner: "崇傑" },
  { name: "南投段(半日)", type: "scooter", approvedMileage: 70, owner: "凱瑋" },
  { name: "西螺段", type: "scooter", approvedMileage: 78, owner: "虹享" },
  { name: "北港段", type: "scooter", approvedMileage: 104, owner: "政偉" },
  { name: "虎尾段", type: "scooter", approvedMileage: 65, owner: "曉謙" },
  { name: "中埔段", type: "scooter", approvedMileage: 66, owner: "秋貴" },
  { name: "朴子段", type: "scooter", approvedMileage: 76, owner: "文聖" },
  { name: "新港段", type: "scooter", approvedMileage: 62, owner: "世文" },
];

const specialStaffSeeds = [
  { name: "宏擷", role: "teamLeader", isRelief: false },
  { name: "凱強", role: "teamLeader", isRelief: false },
  { name: "志佳", role: "reliefStaff", isRelief: true },
  { name: "邦漢", role: "reliefStaff", isRelief: true },
  { name: "翔清", role: "reliefStaff", isRelief: true },
  { name: "明怡", role: "adminStaff", isRelief: false },
  { name: "洪主任", role: "supervisor", isRelief: false },
];

const shiftCoverageOperators = new Set([
  "倉熙",
  "懷慶",
  "庭愷",
  "建仁",
  "哲維",
  "玉女",
  "志剛",
  "良賢",
  "家瑞",
  "逸文",
  "虹享",
  "政偉",
  "曉謙",
  "秋貴",
  "文聖",
  "世文",
]);

// Canonical display order for routes — matches mileage reference table numbering
const ROUTE_DISPLAY_ORDER = routeSeeds.map((r) => r.name);

function sortedRoutes() {
  return [...state.routes].sort((a, b) => {
    const ia = ROUTE_DISPLAY_ORDER.indexOf(a.name);
    const ib = ROUTE_DISPLAY_ORDER.indexOf(b.name);
    // Routes not in seed list go to the end, sorted by name
    if (ia === -1 && ib === -1) return a.name.localeCompare(b.name, "zh-Hant");
    if (ia === -1) return 1;
    if (ib === -1) return -1;
    return ia - ib;
  });
}

function makeId(prefix) {
  return `${prefix}-${Math.random().toString(36).slice(2, 10)}`;
}

function inferShift(routeName) {
  if (routeName.includes("大夜班")) {
    return "night";
  }
  if (routeName.includes("晚班")) {
    return "evening";
  }
  return "day";
}

function isWeekend(dateString) {
  const day = new Date(`${dateString}T00:00:00`).getDay();
  return day === 0 || day === 6;
}

function enumerateDates(startDate, endDate) {
  const dates = [];
  const current = new Date(`${startDate}T00:00:00`);
  const end = new Date(`${endDate}T00:00:00`);
  while (current <= end) {
    const year = current.getFullYear();
    const month = String(current.getMonth() + 1).padStart(2, "0");
    const day = String(current.getDate()).padStart(2, "0");
    dates.push(`${year}-${month}-${day}`);
    current.setDate(current.getDate() + 1);
  }
  return dates;
}

function getWorkingDates(startDate, endDate, companySettings) {
  return enumerateDates(startDate, endDate).filter((dateString) => {
    if (companySettings.weekendDaysOff && isWeekend(dateString)) {
      return false;
    }
    return !companySettings.holidays.includes(dateString);
  });
}

function buildInitialState() {
  const routes = routeSeeds.map((route, index) => ({
    id: `route-${String(index + 1).padStart(3, "0")}`,
    type: route.type,
    name: route.name,
    approvedMileage: route.approvedMileage,
  }));

  const employees = [
    ...specialStaffSeeds.map((employee, index) => ({
      id: `emp-special-${String(index + 1).padStart(2, "0")}`,
      name: employee.name,
      role: employee.role,
      defaultRouteId: "",
      supportLineIds: [],
      isRelief: employee.isRelief,
      canCoverShift: employee.isRelief || ["teamLeader", "supervisor"].includes(employee.role),
      shift: "day",
      employmentStatus: "active",
      active: true,
      fixedDuty: "",
      isNightOwner: false,
    })),
    ...routeSeeds.map((route, index) => ({
      id: `emp-${String(index + 1).padStart(3, "0")}`,
      name: route.owner,
      role: "operator",
      defaultRouteId: routes[index].id,
      supportLineIds: [routes[index].id],
      isRelief: false,
      canCoverShift: shiftCoverageOperators.has(route.owner),
      shift: inferShift(route.name),
      employmentStatus: "active",
      active: true,
      fixedDuty: route.name,
      isNightOwner: route.name === "大夜班",
    })),
  ];

  // 倉熙 is 軍功段 owner but also serves as 抵休 (relief staff)
  const linXixi = employees.find((e) => e.name === "倉熙");
  if (linXixi) { linXixi.role = "reliefStaff"; linXixi.isRelief = true; linXixi.canCoverShift = true; }

  const companySettings = {
    weekendDaysOff: true,
    holidays: [...taiwanHolidaySeeds2026],
  };

  const assignments = [];
  const dates = getWorkingDates("2026-03-18", "2026-03-31", companySettings);
  employees.filter((employee) => employee.defaultRouteId).forEach((employee) => {
    dates.forEach((dateString) => {
      assignments.push({
        id: makeId("asg"),
        date: dateString,
        employeeId: employee.id,
        routeId: employee.defaultRouteId,
        shift: employee.shift,
        status: "working",
        leaveType: "",
        note: "",
        source: "default",
      });
    });
  });

  const nightOwner = employees.find((employee) => employee.name === "長霖");
  const relief = employees.find((employee) => employee.role === "reliefStaff");
  if (nightOwner && relief) {
    const leaveAssignment = assignments.find((assignment) => assignment.employeeId === nightOwner.id && assignment.date === "2026-03-20");
    if (leaveAssignment) {
      leaveAssignment.status = "leave";
      leaveAssignment.leaveType = "annual";
      leaveAssignment.note = "固定大夜班休假";
      leaveAssignment.source = "override";
    }
    assignments.push({
      id: makeId("asg"),
      date: "2026-03-20",
      employeeId: relief.id,
      routeId: nightOwner.defaultRouteId,
      shift: "night",
      status: "reassigned",
      leaveType: "",
      note: "代班大夜班",
      source: "override",
    });
  }

  return {
    session: {
      role: "operator",
      userId: nightOwner?.id || employees[0].id,
      lastGeneratedRange: null,
    },
    labelSettings: {
      shifts: { day: "白班", evening: "晚班", night: "大夜班" },
      leaveTypes: { annual: "特休", personal: "事假", sick: "病假", official: "公假", injury: "公傷", other: "其他" },
      statuses: { working: "上班", leave: "休假", reassigned: "代班", standby: "待命" },
    },
    companySettings,
    pinSettings: {
      teamLeader: "1234",
      adminStaff: "1234",
      supervisor: "0000",
      individual: {},
    },
    employees,
    routes,
    assignments,
    mileageTable: [
      { id: 1, name: "中區段(晚班)", am: "X", pm: 36, total: 36, note: "" },
      { id: 2, name: "軍功段", am: 21, pm: 28, total: 49, note: "" },
      { id: 3, name: "南屯段", am: 13, pm: 34, total: 47, note: "" },
      { id: 4, name: "市政段", am: 14, pm: 39, total: 53, note: "" },
      { id: 5, name: "黎明段", am: 25, pm: 32, total: 57, note: "" },
      { id: 6, name: "松竹段", am: 30, pm: 36, total: 66, note: "" },
      { id: 7, name: "中工段", am: 29, pm: 52, total: 81, note: "" },
      { id: 8, name: "中港段", am: 25, pm: 27, total: 52, note: "" },
      { id: 9, name: "文心段", am: 17, pm: 24, total: 41, note: "" },
      { id: 10, name: "大雅段", am: 27, pm: 39, total: 66, note: "" },
      { id: 11, name: "太平段", am: 20, pm: 25, total: 45, note: "" },
      { id: 12, name: "大里段", am: 32, pm: 32, total: 64, note: "" },
      { id: 13, name: "東勢段", am: 44, pm: 50, total: 94, note: "" },
      { id: 14, name: "后里段", am: 42, pm: 46, total: 88, note: "" },
      { id: 15, name: "清水段", am: 39, pm: 38, total: 77, note: "" },
      { id: 16, name: "沙鹿段", am: 35, pm: 35, total: 70, note: "" },
      { id: 17, name: "梧棲段(半日)", am: "X", pm: 67, total: 67, note: "" },
      { id: 18, name: "玉山大雅(半日)", am: "X", pm: 43, total: 43, note: "" },
      { id: 19, name: "玉山南屯(晚班)", am: "X", pm: 30, total: 30, note: "" },
      { id: 20, name: "市政二段(晚班)", am: "X", pm: 38, total: 38, note: "" },
      { id: 21, name: "南區段(晚班)", am: "X", pm: 28, total: 28, note: "" },
      { id: 22, name: "彰化段", am: 48, pm: 55, total: 103, note: "" },
      { id: 23, name: "和美段", am: 53, pm: 44, total: 97, note: "" },
      { id: 24, name: "員林段", am: 37, pm: 13, total: 50, note: "" },
      { id: 25, name: "溪湖段", am: 18, pm: 26, total: 44, note: "" },
      { id: 26, name: "頭份段", am: 31, pm: 30, total: 61, note: "" },
      { id: 27, name: "竹南段", am: 30, pm: 31, total: 61, note: "" },
      { id: 28, name: "草屯段", am: 53, pm: 35, total: 88, note: "" },
      { id: 29, name: "竹山段", am: "X", pm: "X", total: 116, note: "公務機車" },
      { id: 30, name: "南投段(半日)", am: "X", pm: 70, total: 70, note: "" },
      { id: 31, name: "西螺段", am: 39, pm: 39, total: 78, note: "" },
      { id: 32, name: "北港段", am: 52, pm: 52, total: 104, note: "" },
      { id: 33, name: "虎尾段", am: 32, pm: 33, total: 65, note: "" },
      { id: 34, name: "中埔段", am: 31, pm: 35, total: 66, note: "" },
      { id: 35, name: "朴子段", am: 39, pm: 37, total: 76, note: "" },
      { id: 36, name: "新港段", am: 31, pm: 31, total: 62, note: "" },
    ],
    auditLogs: [
      {
        id: makeId("log"),
        timestamp: "2026-03-18T08:20:00+08:00",
        actorId: employees.find((employee) => employee.role === "teamLeader")?.id || employees[0].id,
        action: "default-sync",
        targetType: "assignment",
        targetId: "seed-defaults",
        summary: "初始化 Excel 固定配置班表",
        detail: "已載入固定路線、人員與核定里程，並建立預設班表。",
      },
    ],
  };
}

const nameRenameMap = new Map([
  ["齊x擷","宏擷"],["林x強","凱強"],["許x佳","志佳"],["劉x漢","邦漢"],["陳x清","翔清"],["張xx怡","明怡"],
  ["何x霖","長霖"],["黃x光","繼光"],["郭x順","智順"],["蔡x緯","琮緯"],["張x謙","佑謙"],["姚x冰","宥冰"],
  ["張x豪","家豪"],["陳x億","炎億"],["藍x民","鴻民"],["邱x庭","裕庭"],["林x熙","倉熙"],["林x裕","國裕"],
  ["黃x慶","懷慶"],["張x愷","庭愷"],["陳x仁","建仁"],["吳x維","哲維"],["黃x女","玉女"],["趙x剛","志剛"],
  ["洪x賢","良賢"],["王x俊","溪俊"],["王x凱","建凱"],["許x瑞","家瑞"],["李x文","逸文"],["楊x豐","德豐"],
  ["劉x宗","育宗"],["江x宗","坤宗"],["吳x峰","孟峰"],["羅x遠","維遠"],["魏x安","珽安"],["傅x瑜","祥瑜"],
  ["歐x豪","書豪"],["賴x瑋","昱瑋"],["魏x洋","新洋"],["葉x昇","鴻昇"],["黃x鈞","彥鈞"],["陳x佑","信佑"],
  ["謝x軒","明軒"],["俞x弘","家弘"],["洪x斌","獻斌"],["柳x傑","崇傑"],["黃x瑋","凱瑋"],["陳x享","虹享"],
  ["蔡x偉","政偉"],["沈x謙","曉謙"],["徐x文","秋貴"],["蔡x貴","文聖"],["林x聖","世文"],
]);

function applyEmployeeMigrations(stateObj) {
  if (!stateObj.employees) return false;
  let changed = false;
  stateObj.employees.forEach((employee) => {
    const newName = nameRenameMap.get(employee.name);
    if (newName) { employee.name = newName; changed = true; }
    // 倉熙 is 軍功段 owner but also serves as 抵休 (relief staff)
    if (employee.name === "倉熙" && employee.role !== "reliefStaff") {
      employee.role = "reliefStaff";
      employee.isRelief = true;
      employee.canCoverShift = true;
      changed = true;
    }
  });
  return changed;
}

function loadState() {
  const stored = localStorage.getItem(STORAGE_KEY);
  if (!stored) {
    return buildInitialState();
  }
  try {
    const parsed = JSON.parse(stored);
    if (!parsed.routes?.[0]?.approvedMileage || !parsed.employees?.some((employee) => "defaultRouteId" in employee)) {
      return buildInitialState();
    }
    applyEmployeeMigrations(parsed);
    parsed.employees.forEach((employee) => {
      const isLeaderOrSupervisor = ["teamLeader", "supervisor"].includes(employee.role);
      const isCoverageOperator = shiftCoverageOperators.has(employee.name);
      if (!employee.employmentStatus) {
        employee.employmentStatus = employee.active ? "active" : "resigned";
      }
      employee.active = employee.employmentStatus === "active";
      if (isLeaderOrSupervisor || isCoverageOperator) {
        employee.canCoverShift = true;
      } else if (typeof employee.canCoverShift !== "boolean") {
        employee.canCoverShift = !!employee.isRelief;
      }
    });
    parsed.session = parsed.session || {};
    parsed.session.lastGeneratedRange = parsed.session.lastGeneratedRange || null;
    parsed.labelSettings = parsed.labelSettings || {};
    parsed.labelSettings.leaveTypes = {
      annual: "特休",
      personal: "事假",
      sick: "病假",
      official: "公假",
      injury: "公傷",
      other: "其他",
      ...(parsed.labelSettings.leaveTypes || {}),
    };
    if (!parsed.pinSettings) {
      parsed.pinSettings = { teamLeader: "1234", adminStaff: "1234", supervisor: "0000", individual: {} };
    }
    if (!parsed.pinSettings.individual) {
      parsed.pinSettings.individual = {};
    }
    if (!parsed.mileageTable) {
      parsed.mileageTable = buildInitialState().mileageTable;
    }
    return parsed;
  } catch (error) {
    return buildInitialState();
  }
}

let state = loadState();
const authenticatedRoles = new Set();

const protectedRoles = ["teamLeader", "adminStaff", "supervisor"];

// authenticatedUsers tracks which individual users have been verified (by employee ID)
const authenticatedUsers = new Set();

// ─── Session Auth Cache (sessionStorage) ───
function loadAuthCache() {
  try {
    const raw = sessionStorage.getItem(SESSION_AUTH_KEY);
    return raw ? JSON.parse(raw) : null;
  } catch { return null; }
}

function saveAuthCache() {
  const toggle = document.querySelector("#keepLoggedInToggle");
  const cached = loadAuthCache();
  const keepLoggedIn = toggle ? toggle.checked : (cached && cached.keepLoggedIn);
  if (!keepLoggedIn) return;
  sessionStorage.setItem(SESSION_AUTH_KEY, JSON.stringify({
    keepLoggedIn: true,
    role: state.session.role,
    userId: state.session.userId,
    authenticatedRoles: [...authenticatedRoles],
    authenticatedUsers: [...authenticatedUsers],
  }));
}

function clearAuthCache() {
  sessionStorage.removeItem(SESSION_AUTH_KEY);
}

function getEmployeePin(employeeId) {
  // Individual PIN takes priority; fallback to role-level PIN
  if (state.pinSettings.individual && state.pinSettings.individual[employeeId]) {
    return state.pinSettings.individual[employeeId];
  }
  const emp = getEmployeeById(employeeId);
  return emp ? (state.pinSettings[emp.role] || "") : "";
}

function showPinDialog(title) {
  return new Promise((resolve) => {
    const overlay = document.createElement("div");
    overlay.className = "pin-overlay";
    const dialog = document.createElement("div");
    dialog.className = "pin-dialog";
    dialog.innerHTML = `
      <p class="pin-title">${title}</p>
      <input type="password" class="pin-input" maxlength="8" placeholder="請輸入 PIN 碼" autocomplete="off">
      <div class="pin-error" style="display:none">PIN 碼錯誤</div>
      <div class="pin-buttons">
        <button type="button" class="pin-cancel">取消</button>
        <button type="button" class="pin-confirm">確認</button>
      </div>
    `;
    overlay.appendChild(dialog);
    document.body.appendChild(overlay);

    const input = dialog.querySelector(".pin-input");
    const errorMsg = dialog.querySelector(".pin-error");
    const confirmBtn = dialog.querySelector(".pin-confirm");
    const cancelBtn = dialog.querySelector(".pin-cancel");

    setTimeout(() => input.focus(), 50);

    function cleanup(result) {
      overlay.remove();
      resolve(result);
    }

    confirmBtn.addEventListener("click", () => cleanup(input.value));
    cancelBtn.addEventListener("click", () => cleanup(null));
    overlay.addEventListener("click", (e) => { if (e.target === overlay) cleanup(null); });
    input.addEventListener("keydown", (e) => {
      if (e.key === "Enter") cleanup(input.value);
      if (e.key === "Escape") cleanup(null);
    });
  });
}

async function requirePin(role) {
  if (!protectedRoles.includes(role)) return true;
  if (authenticatedRoles.has(role)) return true;
  const pin = await showPinDialog(`請輸入「${roleLabels[role]}」的 PIN 碼`);
  if (pin === null) return false;
  if (pin === state.pinSettings[role]) {
    authenticatedRoles.add(role);
    saveAuthCache();
    return true;
  }
  window.alert("PIN 碼錯誤，無法切換至該角色。");
  return false;
}

async function requireUserPin(employeeId) {
  if (!employeeId) return false;
  if (authenticatedUsers.has(employeeId)) return true;
  const emp = getEmployeeById(employeeId);
  if (!emp || !protectedRoles.includes(emp.role)) return true;
  const correctPin = getEmployeePin(employeeId);
  if (!correctPin) return true;
  const pin = await showPinDialog(`請輸入「${emp.name}」的個人 PIN 碼`);
  if (pin === null) return false;
  if (pin === correctPin) {
    authenticatedUsers.add(employeeId);
    authenticatedRoles.add(emp.role);
    saveAuthCache();
    return true;
  }
  window.alert("PIN 碼錯誤，無法切換至此帳號。");
  return false;
}

const appEl = document.querySelector("#app");
const roleSelect = document.querySelector("#roleSelect");
const userSelect = document.querySelector("#userSelect");
const roleHint = document.querySelector("#roleHint");
const keepLoggedInToggle = document.querySelector("#keepLoggedInToggle");
const keepLoggedInLabel = document.querySelector("#keepLoggedInLabel");
const statCardTemplate = document.querySelector("#statCardTemplate");

let firebaseInitialized = false; // true after first Firebase load completes

function saveState() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  if (firebaseReady && firebaseInitialized) {
    lastFirebaseSaveTime = Date.now();
    // Only sync non-session data to Firebase (session is per-device)
    const syncData = { ...state };
    delete syncData.session;
    stateRef.set(syncData)
      .then(() => {
        lastFirebaseSaveTime = Date.now();
      })
      .catch((err) => {
        showSyncStatus("error", "同步失敗：" + err.message);
        console.warn("Firebase save failed:", err);
      });
  }
}

function getToday() {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, "0");
  const day = String(now.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function getEmployeeById(id) {
  return state.employees.find((employee) => employee.id === id);
}

function compareEmployeesByScheduleOrder(a, b) {
  const fixedA = fixedEmployeeOrderByName.get(a?.name) || 999;
  const fixedB = fixedEmployeeOrderByName.get(b?.name) || 999;
  if (fixedA !== fixedB) {
    return fixedA - fixedB;
  }

  const roleA = roleOrderForList[a?.role] || 99;
  const roleB = roleOrderForList[b?.role] || 99;
  if (roleA !== roleB) {
    return roleA - roleB;
  }

  return (a?.name || "").localeCompare((b?.name || ""), "zh-Hant");
}

function getRouteById(id) {
  return state.routes.find((route) => route.id === id);
}

function getDefaultRoute(employee) {
  return employee?.defaultRouteId ? getRouteById(employee.defaultRouteId) : null;
}

function getAssignmentsForEmployee(employeeId) {
  return state.assignments
    .filter((assignment) => assignment.employeeId === employeeId)
    .sort((a, b) => a.date.localeCompare(b.date));
}

function getAssignmentByEmployeeDate(employeeId, dateString) {
  return state.assignments.find((assignment) => assignment.employeeId === employeeId && assignment.date === dateString);
}

function getAssignmentsByDate(dateString) {
  return state.assignments
    .filter((assignment) => assignment.date === dateString)
    .sort((a, b) => compareEmployeesByScheduleOrder(getEmployeeById(a.employeeId), getEmployeeById(b.employeeId)));
}

function getRoleUsers(role) {
  return state.employees
    .filter((employee) => employee.role === role && employee.employmentStatus === "active")
    .sort(compareEmployeesByScheduleOrder);
}

function getLabel(group, key) {
  return state.labelSettings[group][key] || key;
}

function displayRouteType(type) {
  return routeTypeLabels[type] || type;
}

function formatDate(dateString) {
  return new Intl.DateTimeFormat("zh-TW", { month: "numeric", day: "numeric", weekday: "short" })
    .format(new Date(`${dateString}T00:00:00`));
}

function formatTimestamp(timestamp) {
  return new Intl.DateTimeFormat("zh-TW", { month: "numeric", day: "numeric", hour: "2-digit", minute: "2-digit" })
    .format(new Date(timestamp));
}

function getMonthRange(baseDateString = getToday(), offsetMonths = 0) {
  const base = new Date(`${baseDateString}T00:00:00`);
  const year = base.getFullYear();
  const month = base.getMonth() + offsetMonths;
  const first = new Date(year, month, 1);
  const last = new Date(year, month + 1, 0);
  const toDateString = (date) => {
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, "0");
    const d = String(date.getDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
  };
  return {
    year: first.getFullYear(),
    month: first.getMonth() + 1,
    startDate: toDateString(first),
    endDate: toDateString(last),
    dates: enumerateDates(toDateString(first), toDateString(last)),
  };
}

function isHoliday(dateString) {
  if (state.companySettings.weekendDaysOff && isWeekend(dateString)) {
    return true;
  }
  return state.companySettings.holidays.includes(dateString);
}

function getDisplayAssignment(employee, dateString) {
  const assignment = getAssignmentByEmployeeDate(employee.id, dateString);
  if (assignment) {
    return assignment;
  }
  const defaultRoute = getDefaultRoute(employee);
  if (defaultRoute && !isHoliday(dateString)) {
    return {
      employeeId: employee.id,
      date: dateString,
      routeId: defaultRoute.id,
      shift: employee.shift,
      status: "scheduled",
      leaveType: "",
      note: "",
      source: "default-preview",
    };
  }
  return null;
}

// ── Monthly Export Constants & Data Builder ──────────────────────────────

const EXPORT_URBAN_ROUTES = [
  "南屯段", "市政段", "黎明段", "中港段", "中工段",
  "松竹段", "文心段", "大雅段", "太平段", "大里段",
];

function getEmployeeByName(name) {
  return state.employees.find((e) => e.name === name && e.employmentStatus === "active");
}

function getRouteByName(name) {
  return state.routes.find((r) => r.name === name);
}

function getRouteOwner(routeName) {
  const seed = routeSeeds.find((r) => r.name === routeName);
  return seed ? getEmployeeByName(seed.owner) : null;
}

function buildMonthlyExportData(startDate, endDate) {
  const workingDates = getWorkingDates(startDate, endDate, state.companySettings);
  const startDateObj = new Date(`${startDate}T00:00:00`);
  const taiwanYear = startDateObj.getFullYear() - 1911;
  const monthNum = startDateObj.getMonth() + 1;

  // Identify team leaders and relief staff
  const teamLeaders = [
    getEmployeeByName("宏擷"),
    getEmployeeByName("凱強"),
  ].filter(Boolean);

  const reliefStaff = [
    getEmployeeByName("志佳"),
    getEmployeeByName("邦漢"),
  ].filter(Boolean);

  const militaryRelief = getEmployeeByName("倉熙");
  const eveningRelief = getEmployeeByName("翔清");

  // Column headers for display
  const teamLeaderHeaders = teamLeaders.map((e) => ({
    label: "組長",
    name: e.name.replace(/^(.).*?(.)$/, "$1x$2"),
  }));
  const reliefHeaders = reliefStaff.map((e) => ({
    label: "抵休",
    name: e.name.replace(/^(.).*?(.)$/, "$1x$2"),
  }));
  const militaryReliefHeader = {
    label: "軍功/抵休",
    name: militaryRelief ? militaryRelief.name.replace(/^(.).*?(.)$/, "$1x$2") : "",
  };
  const eveningReliefHeader = {
    label: "晚班/抵休",
    name: eveningRelief ? eveningRelief.name.replace(/^(.).*?(.)$/, "$1x$2") : "",
  };

  // Urban route owners
  const urbanRouteInfos = EXPORT_URBAN_ROUTES.map((routeName) => {
    const owner = getRouteOwner(routeName);
    return {
      routeName,
      shortName: routeName.replace("段", ""),
      owner,
      ownerShortName: owner ? owner.name.replace(/^(.).*?(.)$/, "$1x$2") : "",
    };
  });

  const rows = [];
  workingDates.forEach((dateStr, idx) => {
    const dayAssignments = getAssignmentsByDate(dateStr);
    const allDateAssignments = state.assignments.filter((a) => a.date === dateStr);

    // 休假狀況: collect who is on leave (2-char short name: 姓+名尾) with leave type
    const leaveEntries = allDateAssignments.filter((a) => a.status === "leave");
    const leaveNames = leaveEntries.map((a) => {
      const emp = getEmployeeById(a.employeeId);
      if (!emp) return null;
      return {
        name: emp.name.replace(/^(.).*?(.)$/, "$1$2"),
        color: a.leaveType === "annual" ? "green" : "yellow",
      };
    }).filter(Boolean);

    // 特殊記載 & 併線
    const specialNotes = [...new Set(allDateAssignments.filter((a) => a.specialNote).map((a) => a.specialNote))];
    const hasMergedLine = allDateAssignments.some((a) => a.isMergedLine);
    const mergedLineRoutes = hasMergedLine ? ["★"] : [];

    // Helper: get what an employee is doing on this date
    const getEmployeeAction = (emp) => {
      if (!emp) return { text: "", color: null };
      const assignment = getAssignmentByEmployeeDate(emp.id, dateStr);
      if (!assignment) return { text: "", color: null };
      if (assignment.status === "leave") {
        const colorType = assignment.leaveType === "annual" ? "green" : "yellow";
        return { text: "X", color: colorType };
      }
      if (assignment.status === "reassigned" || assignment.source === "override") {
        const route = getRouteById(assignment.routeId);
        if (route) {
          const defaultRoute = getDefaultRoute(emp);
          if (!defaultRoute || route.id !== defaultRoute.id) {
            let text = route.name;
            if (assignment.secondaryRouteId) {
              const secRoute = getRouteById(assignment.secondaryRouteId);
              if (secRoute) text = `(上)${route.name}\n(下)${secRoute.name}`;
            }
            return { text, color: "yellow" };
          }
        }
      }
      return { text: "", color: null };
    };

    // Team leader columns
    const teamLeaderCells = teamLeaders.map((tl) => getEmployeeAction(tl));

    // Relief staff columns
    const reliefCells = reliefStaff.map((rs) => getEmployeeAction(rs));

    // Military relief (倉熙)
    const militaryCell = getEmployeeAction(militaryRelief);

    // Evening relief (翔清)
    const eveningCell = getEmployeeAction(eveningRelief);

    // Urban route columns
    const urbanCells = urbanRouteInfos.map((info) => {
      if (!info.owner) return { text: "", color: null };
      const assignment = getAssignmentByEmployeeDate(info.owner.id, dateStr);
      if (!assignment) return { text: "", color: null };

      if (assignment.status === "leave") {
        const colorType = assignment.leaveType === "annual" ? "green" : "yellow";
        return { text: "X", color: colorType };
      }

      // Check if owner is doing something else (reassigned away)
      if (assignment.source === "override" || assignment.status === "reassigned") {
        const route = getRouteById(assignment.routeId);
        const defaultRoute = getDefaultRoute(info.owner);
        if (route && defaultRoute && route.id !== defaultRoute.id) {
          return { text: route.name, color: "yellow" };
        }
      }

      // Check if someone is covering this route (find assignment with this routeId from someone else)
      const coverer = allDateAssignments.find((a) => {
        if (a.employeeId === info.owner.id) return false;
        const r = getRouteById(a.routeId);
        return r && r.name === info.routeName;
      });
      if (coverer) {
        const covEmp = getEmployeeById(coverer.employeeId);
        if (covEmp) {
          // Show covering employee's short name
          return { text: covEmp.name.replace(/^(.).*?(.)$/, "$1x$2"), color: "lightblue" };
        }
      }

      return { text: "", color: null };
    });

    const dateObj = new Date(`${dateStr}T00:00:00`);
    const dateDisplay = `${dateObj.getMonth() + 1}/${dateObj.getDate()}`;
    const weekday = ["日", "一", "二", "三", "四", "五", "六"][dateObj.getDay()];

    rows.push({
      workDay: idx + 1,
      date: `${dateDisplay}(${weekday})`,
      dateRaw: dateStr,
      leaveNames,
      specialNotes,
      mergedLineRoutes,
      teamLeaderCells,
      reliefCells,
      militaryCell,
      eveningCell,
      urbanCells,
    });
  });

  // Calculate max leave count for individual cell layout
  const maxLeaveCount = Math.max(1, ...rows.map((r) => r.leaveNames.length));

  // 值班組長：單數月=宏擷，雙數月=凱強
  const dutyLeader = monthNum % 2 === 1 ? "宏擷" : "凱強";
  const today = new Date();
  const todayTaiwanYear = today.getFullYear() - 1911;
  const updateDate = `${todayTaiwanYear}.${String(today.getMonth() + 1).padStart(2, "0")}.${String(today.getDate()).padStart(2, "0")}`;

  return {
    title: `${taiwanYear}年${monthNum}月`,
    subtitle: `台中辦事處  抵休代班 、 學線  排程表     值班組長：${dutyLeader}   更新日期：${updateDate}`,
    teamLeaderHeaders,
    reliefHeaders,
    militaryReliefHeader,
    eveningReliefHeader,
    urbanRouteInfos,
    maxLeaveCount,
    rows,
  };
}

function logAction({ actorId, action, targetType, targetId, summary, detail }) {
  state.auditLogs.unshift({
    id: makeId("log"),
    timestamp: new Date().toISOString(),
    actorId,
    action,
    targetType,
    targetId,
    summary,
    detail,
  });
}

function buildStatCard(label, value, meta) {
  const node = statCardTemplate.content.firstElementChild.cloneNode(true);
  node.querySelector(".stat-label").textContent = label;
  node.querySelector(".stat-value").textContent = value;
  node.querySelector(".stat-meta").textContent = meta;
  return node;
}

function createSection(title, subtitle = "") {
  const section = document.createElement("section");
  section.className = "card panel";
  const heading = document.createElement("div");
  heading.className = "section-heading";
  heading.innerHTML = `<div><h2>${title}</h2>${subtitle ? `<p class="muted">${subtitle}</p>` : ""}</div>`;
  section.appendChild(heading);
  return section;
}

function employeeOptions({ includeAll = true, includeReliefOnly = false } = {}) {
  return state.employees
    .filter((employee) => employee.employmentStatus === "active")
    .filter((employee) => (includeReliefOnly ? (employee.isRelief || employee.canCoverShift) : true))
    .filter((employee) => (includeAll ? true : employee.role === "operator"))
    .sort(compareEmployeesByScheduleOrder)
    .map((employee) => {
      const tags = [roleLabels[employee.role]];
      if (employee.isRelief) tags.push("抵休");
      if (!employee.isRelief && employee.canCoverShift) tags.push("可支援代班");
      return `<option value="${employee.id}">${employee.name} | ${tags.join(" | ")}</option>`;
    })
    .join("");
}

function routeOptions() {
  return sortedRoutes()
    .map((route) => `<option value="${route.id}">${route.name} | 核定里程 ${route.approvedMileage}</option>`)
    .join("");
}

function labelOptions(group) {
  return Object.entries(state.labelSettings[group])
    .map(([key, value]) => `<option value="${key}">${value}</option>`)
    .join("");
}

function describeAssignment(assignment) {
  const route = getRouteById(assignment.routeId);
  const secondaryRoute = assignment.secondaryRouteId ? getRouteById(assignment.secondaryRouteId) : null;
  const routeText = route ? route.name : "未指定路線";
  const mergedText = secondaryRoute ? ` + 併線：${secondaryRoute.name}` : "";
  return `${getLabel("shifts", assignment.shift)} / ${routeText}${mergedText} / ${getLabel("statuses", assignment.status)}`;
}

function buildRoleHint(role) {
  if (role === "operator") return "運務員可查看自己的今日班別、路線與假別資訊。";
  if (role === "reliefStaff") return "抵休平常不綁固定路線，主要在異動時支援代班。";
  if (role === "teamLeader") return "組長可生成固定班表、登記請假並安排代班。";
  if (role === "adminStaff") return "行政可維護員工、路線、核定里程與公司休假日。";
  return "主管可查看全體班表、異動紀錄並進行全權編修。";
}

function syncSelectors() {
  // Build role dropdown: for protected roles, show each person individually
  let roleOptions = "";
  Object.entries(roleLabels).forEach(([role, label]) => {
    if (protectedRoles.includes(role)) {
      const users = getRoleUsers(role);
      if (users.length > 1) {
        // Show each person as a separate option: "組長 - 宏擷"
        users.forEach((user) => {
          const val = `${role}:${user.id}`;
          const currentVal = `${state.session.role}:${state.session.userId}`;
          const isSelected = (state.session.role === role && state.session.userId === user.id);
          roleOptions += `<option value="${val}" ${isSelected ? "selected" : ""}>${label} - ${user.name}</option>`;
        });
      } else {
        // Only one person, show normally
        roleOptions += `<option value="${role}" ${state.session.role === role ? "selected" : ""}>${label}</option>`;
      }
    } else {
      roleOptions += `<option value="${role}" ${state.session.role === role ? "selected" : ""}>${label}</option>`;
    }
  });
  roleSelect.innerHTML = roleOptions;

  // User dropdown: still show users for the selected role
  const users = getRoleUsers(state.session.role);
  if (!users.some((user) => user.id === state.session.userId)) {
    state.session.userId = users[0]?.id || "";
  }
  userSelect.innerHTML = users
    .map((user) => `<option value="${user.id}" ${state.session.userId === user.id ? "selected" : ""}>${user.name}</option>`)
    .join("");
}

function renderPrototypeNotice() {
  const section = document.createElement("section");
  section.className = "card panel";
  section.innerHTML = `
    <div class="section-heading">
      <div>
        <h2>正式資料排班測試版</h2>
        <p class="muted">目前使用正式路線與人員名單，先驗證固定配置、請假、代班與異動流程；確認穩定後再持續擴充。</p>
      </div>
    </div>
    <div class="tag-row">
      <span class="tag">員工 ${state.employees.filter((e) => e.employmentStatus === "active").length} 人</span>
      <span class="tag">路線 ${state.routes.length} 條</span>
      <span class="tag">抵休 ${state.employees.filter((e) => e.role === "reliefStaff" && e.employmentStatus === "active").length} 人</span>
      <span class="tag">固定配置 + 異動覆蓋</span>
    </div>
  `;
  return section;
}

function renderDailyBoardLauncher() {
  const section = document.createElement("section");
  section.className = "card panel";
  section.innerHTML = `
    <div class="section-heading">
      <div>
        <h2>操作捷徑</h2>
        <p class="muted">可另開視窗查看今日全體班表。</p>
      </div>
      <button type="button" class="secondary" id="openDailyBoardTopButton">今日全體班表</button>
    </div>
  `;
  section.querySelector("#openDailyBoardTopButton").addEventListener("click", () => {
    openDailyBoardWindow();
  });
  return section;
}

function renderManagementLaunchers() {
  const section = document.createElement("section");
  const lastRange = state.session.lastGeneratedRange;
  const defaultStart = lastRange?.startDate || getMonthRange(getToday(), 0).startDate;
  const defaultEnd = lastRange?.endDate || getMonthRange(getToday(), 0).endDate;
  section.className = "card panel";
  section.innerHTML = `
    <div class="section-heading">
      <div>
        <h2>管理捷徑</h2>
        <p class="muted">可查看今日全體班表，或自選日期區間查看班表與休假總表。</p>
      </div>
      <button type="button" class="secondary" id="openDailyBoardTopButton">今日全體班表</button>
    </div>
    <form id="rangeBoardForm" class="form-grid" style="margin-top:12px;">
      <label>開始日期<input name="startDate" type="date" value="${defaultStart}"></label>
      <label>結束日期<input name="endDate" type="date" value="${defaultEnd}"></label>
      <div class="action-row">
        <button type="submit">查看選擇區間班表</button>
        <button type="button" class="secondary" id="openLeaveSummaryButton">查看休假總表</button>
        <button type="button" class="secondary" id="exportMonthlyPrintBtn">📋 匯出排班表 (列印)</button>
        <button type="button" class="secondary" id="exportMonthlyExcelBtn">📊 匯出排班表 (Excel)</button>
      </div>
    </form>
  `;
  section.querySelector("#openDailyBoardTopButton").addEventListener("click", () => {
    openDailyBoardWindow();
  });
  section.querySelector("#rangeBoardForm").addEventListener("submit", (event) => {
    event.preventDefault();
    const fd = new FormData(event.currentTarget);
    const s = fd.get("startDate");
    const e = fd.get("endDate");
    if (!s || !e || s > e) { window.alert("請確認日期區間正確。"); return; }
    openRangeBoardWindow(s, e);
  });
  section.querySelector("#openLeaveSummaryButton").addEventListener("click", () => {
    const form = section.querySelector("#rangeBoardForm");
    const s = form.elements.startDate.value;
    const e = form.elements.endDate.value;
    if (!s || !e || s > e) { window.alert("請確認日期區間正確。"); return; }
    openLeaveSummaryWindow(s, e);
  });
  section.querySelector("#exportMonthlyPrintBtn").addEventListener("click", () => {
    const form = section.querySelector("#rangeBoardForm");
    const s = form.elements.startDate.value;
    const e = form.elements.endDate.value;
    if (!s || !e || s > e) { window.alert("請確認日期區間正確。"); return; }
    openMonthlySchedulePrintWindow(s, e);
  });
  section.querySelector("#exportMonthlyExcelBtn").addEventListener("click", () => {
    const form = section.querySelector("#rangeBoardForm");
    const s = form.elements.startDate.value;
    const e = form.elements.endDate.value;
    if (!s || !e || s > e) { window.alert("請確認日期區間正確。"); return; }
    exportMonthlyScheduleExcel(s, e);
  });
  return section;
}

function renderStats(currentUser) {
  const stats = document.createElement("section");
  stats.className = "stats-grid";
  const todayAssignments = getAssignmentsByDate(getToday());
  const fixedEmployees = state.employees.filter((employee) => employee.defaultRouteId).length;
  const overrides = todayAssignments.filter((assignment) => assignment.source === "override").length;
  const leaves = todayAssignments.filter((assignment) => assignment.status === "leave").length;
  stats.appendChild(buildStatCard("今日班表", `${todayAssignments.length} 筆`, "固定配置 + 異動覆蓋"));
  stats.appendChild(buildStatCard("固定配置人員", `${fixedEmployees} 人`, "平常有固定路線的人員"));
  stats.appendChild(buildStatCard("今日請假", `${leaves} 筆`, "含連續請假與代班安排"));
  stats.appendChild(buildStatCard("今日異動", `${overrides} 筆`, "請假、代班、臨時支援"));
  stats.appendChild(buildStatCard("目前登入", currentUser.name, roleLabels[currentUser.role]));
  return stats;
}

function renderEmployeeHome(currentUser) {
  const wrap = document.createElement("section");
  wrap.className = "panel-grid";
  const todaySection = createSection("我的今日班別", "系統同時顯示固定配置與當日實際安排。");
  const todayAssignment = getAssignmentByEmployeeDate(currentUser.id, getToday());
  const defaultRoute = getDefaultRoute(currentUser);
  const todayBody = document.createElement("div");
  todayBody.className = "split";
  const highlight = document.createElement("div");
  highlight.className = "employee-highlight";

  if (todayAssignment) {
    const route = getRouteById(todayAssignment.routeId);
    const secondaryRoute = todayAssignment.secondaryRouteId ? getRouteById(todayAssignment.secondaryRouteId) : null;
    const isMerged = !!secondaryRoute;
    const differs = !!defaultRoute && route && defaultRoute.id !== route.id;
    highlight.innerHTML = `
      <div class="tag-row">
        <span class="pill brand">${getLabel("shifts", todayAssignment.shift)}</span>
        <span class="pill ${todayAssignment.status === "leave" ? "alert" : ""}">${getLabel("statuses", todayAssignment.status)}</span>
        ${todayAssignment.leaveType ? `<span class="pill alert">${getLabel("leaveTypes", todayAssignment.leaveType)}</span>` : ""}
        <span class="pill">${todayAssignment.source === "default" ? "固定配置" : "異動覆蓋"}</span>
        ${todayAssignment.status === "reassigned" || differs ? `<span class="pill brand">代班 / 支援</span>` : ""}
        ${isMerged ? `<span class="pill alert">併線</span>` : ""}
      </div>
      ${route
        ? `<div class="big-value">${route.name}${isMerged ? ` + ${secondaryRoute.name}` : ""}</div>
           ${isMerged ? `
             <div class="inline-list" style="margin-bottom:6px;">
               <span class="brand">上午：${route.name}</span>
               <span class="brand">下午：${secondaryRoute.name}</span>
             </div>
             <div class="inline-list">
               <span>請參照上下午路線核定里程表</span>
               <span>${todayAssignment.note || "無備註"}</span>
             </div>
           ` : `
             <div class="inline-list">
               <span>${displayRouteType(route.type)}</span>
               <span>核定里程 ${route.approvedMileage} 公里</span>
               <span>${todayAssignment.note || "無備註"}</span>
             </div>
           `}`
        : `<p class="no-route-notice" style="font-size:0.95rem;font-weight:600;margin:4px 0;">未指定路線</p>
           <p class="muted" style="margin:0;font-size:0.85rem;">${todayAssignment.note || "尚未指派路線，請聯繫管理端。"}</p>`
      }
      <div class="notice">${defaultRoute ? `固定配置：${defaultRoute.name}${differs ? `，今日改派：${route?.name || "未指定路線"}${isMerged ? ` + ${secondaryRoute.name}（併線）` : ""}` : ""}` : "目前沒有固定路線，可由管理端安排支援。"}</div>
    `;
  } else if (defaultRoute) {
    highlight.innerHTML = `
      <div class="tag-row">
        <span class="pill brand">${getLabel("shifts", currentUser.shift)}</span>
        <span class="pill">固定配置</span>
      </div>
      <div class="big-value">${defaultRoute.name}</div>
      <div class="inline-list">
        <span>${displayRouteType(defaultRoute.type)}</span>
        <span>核定里程 ${defaultRoute.approvedMileage} 公里</span>
        <span>今日無異動，沿用固定配置</span>
      </div>
      <div class="notice">若有臨時調整，請由組長或主管在管理端異動。</div>
    `;
  } else {
    highlight.innerHTML = `
      <div class="empty-state" style="text-align:left;">
        <p style="font-size:0.95rem;font-weight:600;margin:0 0 4px;color:var(--muted);">今日未指定路線</p>
        <p class="muted" style="margin:0;font-size:0.85rem;">此角色預設不綁固定路線，若需支援請由管理端安排。</p>
      </div>
    `;
  }

  const profile = document.createElement("div");
  profile.className = "card stat-card";
  profile.innerHTML = `
    <p class="stat-label">人員資訊</p>
    <p class="stat-value">${currentUser.name}</p>
    <div class="tag-row">
      <span class="tag">${roleLabels[currentUser.role]}</span>
      <span class="tag">${getLabel("shifts", currentUser.shift)}</span>
      ${currentUser.isRelief ? `<span class="tag">抵休</span>` : ""}
      ${currentUser.canCoverShift && !currentUser.isRelief ? `<span class="tag">可支援代班</span>` : ""}
      ${currentUser.isNightOwner ? `<span class="tag">固定大夜班</span>` : ""}
    </div>
    <p class="muted">固定路線：${defaultRoute ? `${defaultRoute.name} / 核定里程 ${defaultRoute.approvedMileage} 公里` : "未設定"}</p>
    ${(currentUser.supportLineIds && currentUser.supportLineIds.length > 0) ? `
      <div style="margin-top:8px;">
        <p class="muted" style="margin:0 0 6px;font-size:0.85rem;">可支援路線：</p>
        <div class="inline-list">
          ${currentUser.supportLineIds.map((routeId) => {
            const route = getRouteById(routeId);
            return route ? `<span class="brand">${route.name}</span>` : "";
          }).filter(Boolean).join("")}
        </div>
      </div>
    ` : ""}
  `;

  todayBody.append(highlight, profile);
  todaySection.appendChild(todayBody);

  // 機車上下午里程查核表按鈕
  const mileageBtn = document.createElement("button");
  mileageBtn.className = "secondary";
  mileageBtn.textContent = "機車上下午里程查核表";
  mileageBtn.style.cssText = "margin-top:12px;width:100%;";
  mileageBtn.addEventListener("click", () => openMileageTableWindow());
  todaySection.appendChild(mileageBtn);

  const upcomingSection = createSection("我的近期班表", "顯示未來 10 筆班表，包含固定配置與異動覆蓋。");
  const upcomingAssignments = getAssignmentsForEmployee(currentUser.id)
    .filter((assignment) => assignment.date >= getToday())
    .slice(0, 10);
  upcomingSection.appendChild(renderAssignmentTable(upcomingAssignments, false));

  wrap.append(todaySection, upcomingSection);
  return wrap;
}

function renderAssignmentTable(assignments, includeEmployee = true) {
  if (!assignments.length) {
    const empty = document.createElement("div");
    empty.className = "empty-state";
    empty.innerHTML = "<p>目前沒有可顯示的班表資料。</p>";
    return empty;
  }

  const wrap = document.createElement("div");
  wrap.className = "table-card";
  const headers = ["日期", ...(includeEmployee ? ["員工"] : []), "固定配置", "今日路線", "核定里程", "狀態", "假別", "來源", "備註"];
  const rows = assignments.map((assignment) => {
    const employee = getEmployeeById(assignment.employeeId);
    const route = getRouteById(assignment.routeId);
    const secondaryRoute = assignment.secondaryRouteId ? getRouteById(assignment.secondaryRouteId) : null;
    const defaultRoute = getDefaultRoute(employee);
    const routeDisplay = route ? (secondaryRoute ? `上午：${route.name}<br>下午：${secondaryRoute.name}` : route.name) : "-";
    const mileageDisplay = secondaryRoute ? "參照里程表" : (route ? `${route.approvedMileage} 公里` : "-");
    const statusDisplay = getLabel("statuses", assignment.status) + (secondaryRoute ? " / 併線" : "");
    return `
      <tr>
        <td>${formatDate(assignment.date)}</td>
        ${includeEmployee ? `<td>${employee ? employee.name : assignment.employeeId}</td>` : ""}
        <td>${defaultRoute ? defaultRoute.name : "未指定路線"}</td>
        <td>${routeDisplay}</td>
        <td>${mileageDisplay}</td>
        <td>${statusDisplay}</td>
        <td>${assignment.leaveType ? getLabel("leaveTypes", assignment.leaveType) : "-"}</td>
        <td>${assignment.source === "default" ? "固定配置" : "異動覆蓋"}</td>
        <td>${assignment.note || "-"}</td>
      </tr>
    `;
  }).join("");

  wrap.innerHTML = `
    <table>
      <thead>
        <tr>${headers.map((header) => `<th>${header}</th>`).join("")}</tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  `;
  return wrap;
}

function renderSchedulingWorkbench(currentUser) {
  return renderSchedulingWorkbenchV2(currentUser);
}

function renderSchedulingWorkbenchV2(currentUser) {
  const section = createSection("排班與異動工作台", "先生成固定班表，再針對請假、代班或支援做後續調整。");
  const grid = document.createElement("div");
  grid.className = "panel-grid";

  const defaultForm = document.createElement("div");
  defaultForm.className = "card panel";
  defaultForm.innerHTML = `
    <div class="section-heading">
      <div>
        <h3>生成固定班表</h3>
        <p class="muted">請先自行確認日期區間，再產生固定班表。</p>
      </div>
    </div>
    <form id="defaultScheduleFormV2" class="form-grid">
      <label>開始日期<input name="startDate" type="date" value="${getToday()}"></label>
      <label>結束日期<input name="endDate" type="date" value="${getMonthRange(getToday(), 0).endDate}"></label>
      <label>指定員工<select name="employeeId"><option value="">全部固定班表人員</option>${employeeOptions({ includeAll: true })}</select></label>
      <button type="submit">生成指定區間固定班表</button>
      <button type="button" class="secondary" id="fillNextMonthButton">帶入次月日期</button>
    </form>
  `;

  const leaveForm = document.createElement("div");
  leaveForm.className = "card panel";
  leaveForm.innerHTML = `
    <div class="section-heading">
      <div>
        <h3>登記請假</h3>
        <p class="muted">選擇員工與日期區間，一次標記整段休假。代班請到下方另行指派。</p>
      </div>
    </div>
    <form id="leaveOnlyForm" class="form-grid">
      <label>開始日期<input name="startDate" type="date" value="${getToday()}"></label>
      <label>結束日期<input name="endDate" type="date" value="${getToday()}"></label>
      <label>請假員工<select name="employeeId">${employeeOptions({ includeAll: true })}</select></label>
      <label>假別<select name="leaveType">${labelOptions("leaveTypes")}</select></label>
      <label>備註<textarea name="note" placeholder="例如：連續特休、家庭因素"></textarea></label>
      <button type="submit">儲存請假</button>
    </form>
  `;

  const reliefForm = document.createElement("div");
  reliefForm.className = "card panel";
  reliefForm.innerHTML = `
    <div class="section-heading">
      <div>
        <h3>派遣代班</h3>
        <p class="muted">針對已請假的員工，分段指派不同代班人員。可多次操作。</p>
      </div>
    </div>
    <form id="reliefOnlyForm" class="form-grid">
      <label>代班開始日期<input name="startDate" type="date" value="${getToday()}"></label>
      <label>代班結束日期<input name="endDate" type="date" value="${getToday()}"></label>
      <label>代班人員<select name="reliefEmployeeId">${employeeOptions({ includeReliefOnly: true })}</select></label>
      <label>代班路線（上午 / 主要）<select name="routeId">${routeOptions()}</select></label>
      <label>併線<select name="isMergedLine"><option value="no">否</option><option value="yes">是</option></select></label>
      <label class="merged-route-label" style="display:none;">併線路線（下午）<select name="secondaryRouteId"><option value="">無（不併線）</option>${routeOptions()}</select></label>
      <label>特殊記載<textarea name="specialNote" placeholder="例如：預交提出、月底提出"></textarea></label>
      <label>備註<textarea name="note" placeholder="例如：支援大夜班、臨時調度"></textarea></label>
      <button type="submit">儲存代班</button>
    </form>
  `;

  const editPanel = document.createElement("div");
  editPanel.className = "card panel";
  editPanel.innerHTML = `
    <div class="section-heading">
      <div>
        <h3>編輯 / 刪除排班</h3>
        <p class="muted">輸入錯誤的請假或代班，可查詢單日或日期區間後修改或批量刪除。</p>
      </div>
    </div>
    <div class="form-grid">
      <label>開始日期<input type="date" id="editLookupStartDate" value="${getToday()}"></label>
      <label>結束日期<input type="date" id="editLookupEndDate" value="${getToday()}"></label>
      <label>員工<select id="editLookupEmployee">${employeeOptions({ includeAll: true })}</select></label>
      <button type="button" id="editLookupButton">查詢排班</button>
    </div>
    <div id="editResultArea"></div>
  `;

  grid.append(defaultForm, leaveForm, reliefForm, editPanel);
  section.appendChild(grid);

  editPanel.querySelector("#editLookupButton").addEventListener("click", () => {
    const startDate = editPanel.querySelector("#editLookupStartDate").value;
    const endDate = editPanel.querySelector("#editLookupEndDate").value;
    const empId = editPanel.querySelector("#editLookupEmployee").value;
    const resultArea = editPanel.querySelector("#editResultArea");
    if (!startDate || !endDate || !empId) { window.alert("請選擇日期區間與員工。"); return; }
    if (startDate > endDate) { window.alert("結束日期不可早於開始日期。"); return; }

    const emp = getEmployeeById(empId);
    const dates = enumerateDates(startDate, endDate);
    const matches = dates
      .map((d) => ({ date: d, assignment: getAssignmentByEmployeeDate(empId, d) }))
      .filter((item) => item.assignment);

    if (!matches.length) {
      resultArea.innerHTML = `<div class="empty-state" style="margin-top:12px;"><p>${emp?.name || empId} 在 ${formatDate(startDate)} ~ ${formatDate(endDate)} 沒有排班記錄。</p></div>`;
      return;
    }

    const isSingle = matches.length === 1;

    const listHtml = matches.map((item) => {
      const a = item.assignment;
      const route = getRouteById(a.routeId);
      const secRoute = a.secondaryRouteId ? getRouteById(a.secondaryRouteId) : null;
      return `
        <div class="card stat-card" style="margin-top:8px;" data-asg-id="${a.id}">
          <div style="display:flex;align-items:center;justify-content:space-between;gap:8px;flex-wrap:wrap;">
            <div>
              <p class="stat-label">${formatDate(a.date)}</p>
              <div class="tag-row" style="margin:6px 0;">
                <span class="pill brand">${getLabel("shifts", a.shift)}</span>
                <span class="pill ${a.status === "leave" ? "alert" : ""}">${getLabel("statuses", a.status)}</span>
                ${a.leaveType ? `<span class="pill alert">${getLabel("leaveTypes", a.leaveType)}</span>` : ""}
                <span class="pill">${a.source === "default" ? "固定配置" : "異動覆蓋"}</span>
                ${a.isMergedLine ? `<span class="pill alert">併線</span>` : ""}
              </div>
              <p class="muted" style="margin:0;">路線：${route ? route.name : "未指定"}${secRoute ? ` ｜ 併線：${secRoute.name}` : ""}${a.specialNote ? ` ｜ 特殊記載：${a.specialNote}` : ""} ｜ 備註：${a.note || "無"}</p>
            </div>
            <div style="display:flex;gap:6px;flex-shrink:0;">
              <button type="button" class="secondary editSingleButton" data-asg-id="${a.id}">編輯</button>
              <button type="button" class="warn deleteSingleButton" data-asg-id="${a.id}">刪除</button>
            </div>
          </div>
        </div>`;
    }).join("");

    const batchHtml = !isSingle ? `
      <div style="margin-top:14px;display:flex;gap:8px;flex-wrap:wrap;">
        <button type="button" class="warn" id="batchDeleteButton">批量刪除以上 ${matches.length} 筆排班</button>
      </div>` : "";

    resultArea.innerHTML = `
      <div style="margin-top:12px;">
        <p style="margin:0 0 4px;font-weight:700;">${emp?.name || empId}　${formatDate(startDate)} ~ ${formatDate(endDate)}　共 ${matches.length} 筆</p>
        ${listHtml}
        ${batchHtml}
      </div>`;

    resultArea.querySelectorAll(".editSingleButton").forEach((btn) => {
      btn.addEventListener("click", () => {
        const asgId = btn.dataset.asgId;
        const asg = state.assignments.find((a) => a.id === asgId);
        if (!asg) { window.alert("找不到此筆排班。"); return; }
        const asgRoute = getRouteById(asg.routeId);

        const editArea = document.createElement("div");
        editArea.className = "card panel";
        editArea.style.marginTop = "10px";
        editArea.innerHTML = `
          <p class="stat-label" style="margin-bottom:8px;">編輯 ${emp?.name || asg.employeeId} ${formatDate(asg.date)}</p>
          <form class="form-grid editSingleForm">
            <input name="assignmentId" type="hidden" value="${asg.id}">
            <label>狀態<select name="status">${Object.entries(state.labelSettings.statuses).map(([k, v]) => `<option value="${k}" ${asg.status === k ? "selected" : ""}>${v}</option>`).join("")}</select></label>
            <label>假別<select name="leaveType">${Object.entries(state.labelSettings.leaveTypes).map(([k, v]) => `<option value="${k}" ${asg.leaveType === k ? "selected" : ""}>${v}</option>`).join("")}<option value="" ${!asg.leaveType ? "selected" : ""}>無</option></select></label>
            <label>班別<select name="shift">${Object.entries(state.labelSettings.shifts).map(([k, v]) => `<option value="${k}" ${asg.shift === k ? "selected" : ""}>${v}</option>`).join("")}</select></label>
            <label>路線（上午 / 主要）<select name="routeId">${sortedRoutes().map((r) => `<option value="${r.id}" ${asg.routeId === r.id ? "selected" : ""}>${r.name}</option>`).join("")}</select></label>
            <label>併線<select name="isMergedLine"><option value="no" ${!asg.isMergedLine ? "selected" : ""}>否</option><option value="yes" ${asg.isMergedLine ? "selected" : ""}>是</option></select></label>
            <label>併線路線（下午）<select name="secondaryRouteId"><option value="" ${!asg.secondaryRouteId ? "selected" : ""}>無（不併線）</option>${sortedRoutes().map((r) => `<option value="${r.id}" ${asg.secondaryRouteId === r.id ? "selected" : ""}>${r.name}</option>`).join("")}</select></label>
            <label>特殊記載<textarea name="specialNote">${asg.specialNote || ""}</textarea></label>
            <label>備註<textarea name="note">${asg.note || ""}</textarea></label>
            <button type="submit">儲存修改</button>
          </form>`;

        const cardEl = btn.closest("[data-asg-id]");
        const existingEdit = cardEl.nextElementSibling;
        if (existingEdit && existingEdit.classList.contains("panel")) existingEdit.remove();
        cardEl.after(editArea);

        editArea.querySelector(".editSingleForm").addEventListener("submit", (event) => {
          event.preventDefault();
          const fd = new FormData(event.currentTarget);
          const target = state.assignments.find((a) => a.id === fd.get("assignmentId"));
          if (!target) { window.alert("找不到此筆排班。"); return; }
          target.status = fd.get("status");
          target.leaveType = fd.get("status") === "leave" ? fd.get("leaveType") : "";
          target.shift = fd.get("shift");
          target.routeId = fd.get("routeId");
          target.isMergedLine = fd.get("isMergedLine") === "yes";
          target.secondaryRouteId = target.isMergedLine ? (fd.get("secondaryRouteId") || "") : "";
          target.specialNote = (fd.get("specialNote") || "").trim();
          target.note = (fd.get("note") || "").trim();
          target.source = "override";
          logAction({
            actorId: currentUser.id,
            action: "assignment-edit",
            targetType: "assignment",
            targetId: target.id,
            summary: `編輯排班 ${emp?.name || target.employeeId} ${target.date}`,
            detail: describeAssignment(target),
          });
          saveState();
          window.alert("排班已更新。");
          render();
        });
      });
    });

    resultArea.querySelectorAll(".deleteSingleButton").forEach((btn) => {
      btn.addEventListener("click", () => {
        const asgId = btn.dataset.asgId;
        const asg = state.assignments.find((a) => a.id === asgId);
        if (!asg) return;
        const confirmed = window.confirm(`確定要刪除 ${emp?.name || empId} 在 ${formatDate(asg.date)} 的排班嗎？\n\n刪除後該員工當天將回到固定配置（若有），或變成無排班。`);
        if (!confirmed) return;
        state.assignments = state.assignments.filter((a) => a.id !== asgId);
        logAction({
          actorId: currentUser.id,
          action: "assignment-delete",
          targetType: "assignment",
          targetId: asgId,
          summary: `刪除排班 ${emp?.name || empId} ${asg.date}`,
          detail: `已移除該筆排班記錄。`,
        });
        saveState();
        window.alert("排班已刪除。");
        render();
      });
    });

    const batchBtn = resultArea.querySelector("#batchDeleteButton");
    if (batchBtn) {
      batchBtn.addEventListener("click", () => {
        const ids = matches.map((item) => item.assignment.id);
        const confirmed = window.confirm(`確定要批量刪除 ${emp?.name || empId} 在 ${formatDate(startDate)} ~ ${formatDate(endDate)} 的 ${ids.length} 筆排班嗎？\n\n刪除後這些天將回到固定配置（若有），或變成無排班。`);
        if (!confirmed) return;
        const idSet = new Set(ids);
        state.assignments = state.assignments.filter((a) => !idSet.has(a.id));
        logAction({
          actorId: currentUser.id,
          action: "assignment-batch-delete",
          targetType: "assignment",
          targetId: empId,
          summary: `批量刪除 ${emp?.name || empId} ${startDate} ~ ${endDate} 共 ${ids.length} 筆排班`,
          detail: `已移除 ${ids.length} 筆排班記錄。`,
        });
        saveState();
        window.alert(`已批量刪除 ${ids.length} 筆排班。`);
        render();
      });
    }
  });

  defaultForm.querySelector("#defaultScheduleFormV2").addEventListener("submit", (event) => {
    event.preventDefault();
    const changed = generateDefaultAssignments(new FormData(event.currentTarget), currentUser);
    if (changed) render();
  });
  defaultForm.querySelector("#fillNextMonthButton").addEventListener("click", () => {
    applyNextMonthRange(defaultForm.querySelector("#defaultScheduleFormV2"));
  });

  leaveForm.querySelector("#leaveOnlyForm").addEventListener("submit", (event) => {
    event.preventDefault();
    const changed = applyLeaveOnly(new FormData(event.currentTarget), currentUser);
    if (changed) render();
  });

  const mergedToggle = reliefForm.querySelector('[name="isMergedLine"]');
  const mergedRouteLabel = reliefForm.querySelector('.merged-route-label');
  mergedToggle.addEventListener("change", () => {
    mergedRouteLabel.style.display = mergedToggle.value === "yes" ? "" : "none";
    if (mergedToggle.value === "no") {
      mergedRouteLabel.querySelector('[name="secondaryRouteId"]').value = "";
    }
  });

  reliefForm.querySelector("#reliefOnlyForm").addEventListener("submit", (event) => {
    event.preventDefault();
    const changed = applyReliefOnly(new FormData(event.currentTarget), currentUser);
    if (changed) render();
  });

  return section;
}
function openDailyBoardWindow() {
  const popup = window.open("", "daily-board-window", "width=1180,height=820,scrollbars=yes,resizable=yes");
  if (!popup) {
    window.alert("請先允許瀏覽器開啟彈出視窗，才能查看今日全體班表。");
    return;
  }

  const rows = getAssignmentsByDate(getToday()).map((assignment) => {
    const employee = getEmployeeById(assignment.employeeId);
    const route = getRouteById(assignment.routeId);
    const secondaryRoute = assignment.secondaryRouteId ? getRouteById(assignment.secondaryRouteId) : null;
    const defaultRoute = getDefaultRoute(employee);
    const routeDisplay = route ? (secondaryRoute ? `上午：${route.name} / 下午：${secondaryRoute.name}` : route.name) : "-";
    const mileageDisplay = secondaryRoute ? "參照里程表" : (route ? `${route.approvedMileage} 公里` : "-");
    const statusDisplay = getLabel("statuses", assignment.status) + (secondaryRoute ? " / 併線" : "");
    return `
      <tr>
        <td>${employee ? employee.name : assignment.employeeId}</td>
        <td>${roleLabels[employee?.role] || "-"}</td>
        <td>${defaultRoute ? defaultRoute.name : "未指定路線"}</td>
        <td>${routeDisplay}</td>
        <td>${mileageDisplay}</td>
        <td>${statusDisplay}</td>
        <td>${assignment.source === "default" ? "固定配置" : "異動覆蓋"}</td>
        <td>${assignment.note || "-"}</td>
      </tr>
    `;
  }).join("");

  popup.document.open();
  popup.document.write(`
    <!DOCTYPE html>
    <html lang="zh-Hant">
    <head>
      <meta charset="UTF-8">
      <title>今日全體班表</title>
      <style>
        body { font-family: "Segoe UI", "Noto Sans TC", sans-serif; margin: 0; padding: 24px; background: #f6f1e8; color: #2f2418; }
        h1 { margin: 0 0 8px; }
        p { margin: 0 0 18px; color: #6f6254; }
        .meta { display: flex; gap: 10px; flex-wrap: wrap; margin-bottom: 18px; }
        .meta span { background: #fff8ef; border: 1px solid #ead8c2; border-radius: 999px; padding: 8px 12px; font-weight: 600; }
        table { width: 100%; border-collapse: collapse; background: #fffdf9; }
        th, td { border-bottom: 1px solid #eadfd0; padding: 10px 12px; text-align: left; vertical-align: top; }
        th { background: #f1e3d1; }
      </style>
    </head>
    <body>
      <h1>今日全體班表</h1>
      <p>${getToday()} 的全體排班與異動資訊</p>
      <div class="meta">
        <span>總筆數 ${getAssignmentsByDate(getToday()).length}</span>
        <span>異動覆蓋 ${getAssignmentsByDate(getToday()).filter((assignment) => assignment.source === "override").length}</span>
        <span>請假 ${getAssignmentsByDate(getToday()).filter((assignment) => assignment.status === "leave").length}</span>
      </div>
      <table>
        <thead>
          <tr>
            <th>員工</th>
            <th>角色</th>
            <th>固定配置</th>
            <th>今日路線</th>
            <th>核定里程</th>
            <th>狀態</th>
            <th>來源</th>
            <th>備註</th>
          </tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>
    </body>
    </html>
  `);
  popup.document.close();
  popup.focus();
}
function openMileageTableWindow() {
  const canEdit = ["teamLeader", "adminStaff", "supervisor"].includes(state.session.role);
  const data = state.mileageTable || [];
  const rows = data.map((item) => {
    const amClass = item.am === "X" ? ' class="na-cell"' : "";
    const pmClass = item.pm === "X" ? ' class="na-cell"' : "";
    const amVal = item.am === "X" ? "X" : item.am;
    const pmVal = item.pm === "X" ? "X" : item.pm;
    return `
      <tr>
        <td class="center">${item.id}</td>
        <td class="route-name">${item.name}</td>
        <td${amClass}${canEdit && item.am !== "X" ? ` class="editable" data-id="${item.id}" data-field="am"` : ""}>${amVal}</td>
        <td${pmClass}${canEdit && item.pm !== "X" ? ` class="editable" data-id="${item.id}" data-field="pm"` : ""}>${pmVal}</td>
        <td class="total-cell">${item.total}</td>
        <td class="note-cell"${canEdit ? ` data-id="${item.id}" data-field="note"` : ""}>${item.note || ""}</td>
      </tr>
    `;
  }).join("");

  const popup = window.open("", "mileage-table", "width=800,height=900,scrollbars=yes,resizable=yes");
  if (!popup) { window.alert("請先允許瀏覽器開啟彈出視窗。"); return; }

  popup.document.open();
  popup.document.write(`
    <!DOCTYPE html>
    <html lang="zh-Hant">
    <head>
      <meta charset="UTF-8">
      <title>機車上下午里程查核表</title>
      <style>
        body { font-family: "Segoe UI", "Noto Sans TC", sans-serif; margin: 0; padding: 24px; background: #f6f1e8; color: #2f2418; }
        h1 { margin: 0 0 4px; font-size: 1.3rem; }
        .subtitle { margin: 0 0 16px; color: #6f6254; font-size: 0.9rem; }
        .edit-hint { margin: 0 0 16px; padding: 8px 12px; background: #fff0e8; border-radius: 10px; color: #a0401a; font-size: 0.85rem; font-weight: 600; }
        table { width: 100%; border-collapse: collapse; background: #fffdf9; }
        th, td { border: 1px solid #eadfd0; padding: 8px 10px; text-align: center; vertical-align: middle; font-size: 0.9rem; }
        th { background: #f1e3d1; font-weight: 700; white-space: nowrap; }
        .route-name { text-align: left; font-weight: 600; white-space: nowrap; }
        .na-cell { color: #b0a090; }
        .total-cell { font-weight: 700; background: #fff8ef; }
        .note-cell { text-align: left; font-size: 0.85rem; color: #6f6254; }
        .center { text-align: center; }
        .editable { cursor: pointer; position: relative; }
        .editable:hover { background: #ffecd6; }
        .edit-input { width: 50px; padding: 4px; border: 2px solid #bf5b2c; border-radius: 6px; text-align: center; font-size: 0.9rem; outline: none; }
        .note-input { width: 100%; padding: 4px; border: 2px solid #bf5b2c; border-radius: 6px; font-size: 0.85rem; outline: none; }
        .save-msg { position: fixed; bottom: 20px; right: 20px; background: #1d6b63; color: #fff; padding: 10px 18px; border-radius: 10px; font-weight: 600; opacity: 0; transition: opacity 0.3s; z-index: 99; }
        .save-msg.show { opacity: 1; }
        @media print {
          body { padding: 8px; background: #fff; }
          .edit-hint { display: none; }
          .editable:hover { background: none; }
          th, td { padding: 4px 6px; font-size: 10px; }
        }
      </style>
    </head>
    <body>
      <h1>台中辦事處 機車上、下午里程查核表</h1>
      <p class="subtitle">共 ${data.length} 條路線${canEdit ? "（點擊數字可編輯）" : ""}</p>
      ${canEdit ? '<p class="edit-hint">點擊里程數字或備註即可直接編輯，修改後自動儲存同步。</p>' : ""}
      <table>
        <thead>
          <tr>
            <th style="width:40px;">編號</th>
            <th>路線名稱</th>
            <th style="width:75px;">上午里程</th>
            <th style="width:75px;">下午里程</th>
            <th style="width:75px;">總里程</th>
            <th>備註</th>
          </tr>
        </thead>
        <tbody id="mileageBody">${rows}</tbody>
      </table>
      <div class="save-msg" id="saveMsg">已儲存</div>
      ${canEdit ? `
      <script>
        var saveTimeout;
        function showSaved() {
          var msg = document.getElementById("saveMsg");
          msg.classList.add("show");
          clearTimeout(saveTimeout);
          saveTimeout = setTimeout(function() { msg.classList.remove("show"); }, 1500);
        }
        document.getElementById("mileageBody").addEventListener("click", function(e) {
          var td = e.target.closest("td");
          if (!td || !td.classList.contains("editable") && !td.hasAttribute("data-field")) return;
          if (td.querySelector("input")) return;
          var field = td.getAttribute("data-field") || td.dataset.field;
          var id = parseInt(td.getAttribute("data-id") || td.dataset.id);
          if (!field) return;
          var oldVal = td.textContent.trim();
          var isNote = (field === "note");
          var input = document.createElement("input");
          input.className = isNote ? "note-input" : "edit-input";
          input.type = isNote ? "text" : "number";
          input.value = oldVal;
          if (!isNote) { input.min = "0"; input.step = "1"; }
          td.textContent = "";
          td.appendChild(input);
          input.focus();
          input.select();
          function save() {
            var newVal = isNote ? input.value.trim() : (parseInt(input.value) || 0);
            td.textContent = isNote ? newVal : newVal;
            if (window.opener && window.opener.updateMileageItem) {
              window.opener.updateMileageItem(id, field, newVal);
              showSaved();
              // Refresh total
              if (!isNote) {
                var row = td.closest("tr");
                var cells = row.querySelectorAll("td");
                var amText = cells[2].textContent.trim();
                var pmText = cells[3].textContent.trim();
                var am = (amText === "X") ? 0 : parseInt(amText) || 0;
                var pm = (pmText === "X") ? 0 : parseInt(pmText) || 0;
                var total = am + pm;
                cells[4].textContent = total;
                window.opener.updateMileageItem(id, "total", total);
              }
            }
          }
          input.addEventListener("blur", save);
          input.addEventListener("keydown", function(ev) { if (ev.key === "Enter") { input.blur(); } });
        });
      </script>
      ` : ""}
    </body>
    </html>
  `);
  popup.document.close();
  popup.focus();
}

// Global callback for mileage popup to update state
window.updateMileageItem = function(id, field, value) {
  if (!state.mileageTable) return;
  const item = state.mileageTable.find((m) => m.id === id);
  if (item) {
    item[field] = value;
    saveState();
  }
};

function openLeaveSummaryWindow(startDate, endDate) {
  const allDates = enumerateDates(startDate, endDate).filter((d) => !isHoliday(d));
  const leaveAssignments = state.assignments
    .filter((a) => a.date >= startDate && a.date <= endDate && a.status === "leave")
    .sort((a, b) => a.date.localeCompare(b.date) || (getEmployeeById(a.employeeId)?.name || "").localeCompare(getEmployeeById(b.employeeId)?.name || "", "zh-Hant"));

  const popup = window.open("", `leave-summary-${startDate}-${endDate}`, "width=1200,height=850,scrollbars=yes,resizable=yes");
  if (!popup) {
    window.alert("請先允許瀏覽器開啟彈出視窗。");
    return;
  }

  const dateGroups = new Map();
  leaveAssignments.forEach((a) => {
    if (!dateGroups.has(a.date)) dateGroups.set(a.date, []);
    dateGroups.get(a.date).push(a);
  });

  const uniqueEmployees = new Set(leaveAssignments.map((a) => a.employeeId));

  // 統計各假別人次
  const leaveTypeCounts = {};
  leaveAssignments.forEach((a) => {
    const label = getLabel("leaveTypes", a.leaveType);
    leaveTypeCounts[label] = (leaveTypeCounts[label] || 0) + 1;
  });
  const leaveTypeSpans = Object.entries(leaveTypeCounts)
    .map(([label, count]) => `<span>${label} ${count} 人次</span>`)
    .join("");

  let detailRows = "";
  for (const [date, assignments] of dateGroups) {
    assignments.forEach((a, idx) => {
      const emp = getEmployeeById(a.employeeId);
      detailRows += `
        <tr${idx === 0 ? ' class="date-first"' : ""}>
          ${idx === 0 ? `<td rowspan="${assignments.length}" class="date-cell">${formatDate(date)}</td>` : ""}
          <td>${emp ? emp.name : a.employeeId}</td>
          <td><strong>${getLabel("leaveTypes", a.leaveType)}</strong></td>
          <td>${a.note || "-"}</td>
        </tr>
      `;
    });
  }

  if (!leaveAssignments.length) {
    detailRows = `<tr><td colspan="4" style="text-align:center;padding:24px;color:#6f6254;">此區間沒有休假記錄。</td></tr>`;
  }

  popup.document.open();
  popup.document.write(`
    <!DOCTYPE html>
    <html lang="zh-Hant">
    <head>
      <meta charset="UTF-8">
      <title>休假總表 ${startDate} ~ ${endDate}</title>
      <style>
        body { font-family: "Segoe UI", "Noto Sans TC", sans-serif; margin: 0; padding: 24px; background: #f6f1e8; color: #2f2418; }
        h1 { margin: 0 0 8px; }
        p { margin: 0 0 18px; color: #6f6254; }
        .meta { display: flex; gap: 10px; flex-wrap: wrap; margin-bottom: 18px; }
        .meta span { background: #fff8ef; border: 1px solid #ead8c2; border-radius: 999px; padding: 8px 12px; font-weight: 600; }
        .meta span.alert { background: #fff0e8; border-color: #f0c4a8; color: #a0401a; }
        table { width: 100%; border-collapse: collapse; background: #fffdf9; }
        th, td { border-bottom: 1px solid #eadfd0; padding: 10px 12px; text-align: left; vertical-align: top; }
        th { background: #f1e3d1; white-space: nowrap; }
        .date-cell { background: #fff8ef; font-weight: 700; white-space: nowrap; }
        .date-first td { border-top: 2px solid #d9c8b0; }
        @media print {
          body { padding: 8px; background: #fff; }
          .meta span { padding: 4px 8px; font-size: 11px; }
          th, td { padding: 6px 8px; }
        }
      </style>
    </head>
    <body>
      <h1>休假總表</h1>
      <p>${startDate} 至 ${endDate}，僅顯示工作日的休假記錄，方便行政開立假單。</p>
      <div class="meta">
        <span class="alert">休假總人次 ${leaveAssignments.length}</span>
        <span>涉及員工 ${uniqueEmployees.size} 人</span>
        <span>涵蓋工作日 ${allDates.length} 天</span>
        <span>有休假的天數 ${dateGroups.size} 天</span>
        ${leaveTypeSpans}
      </div>
      <table>
        <thead>
          <tr>
            <th>日期</th>
            <th>員工</th>
            <th>假別</th>
            <th>備註</th>
          </tr>
        </thead>
        <tbody>${detailRows}</tbody>
      </table>
    </body>
    </html>
  `);
  popup.document.close();
  popup.focus();
}

// ── Monthly Schedule Export (Print + Excel) ──────────────────────────────

function openMonthlySchedulePrintWindow(startDate, endDate) {
  const data = buildMonthlyExportData(startDate, endDate);
  const popup = window.open("", `monthly-schedule-${startDate}-${endDate}`, "width=1600,height=900,scrollbars=yes,resizable=yes");
  if (!popup) { window.alert("請先允許瀏覽器開啟彈出視窗。"); return; }

  const leaveCols = data.maxLeaveCount;
  const totalCols = 2 + leaveCols + 2 + 1 + data.teamLeaderHeaders.length + data.reliefHeaders.length + 1 + 1 + data.urbanRouteInfos.length;

  const colorMap = { green: "#92D050", yellow: "#FFFF00", orange: "#FFC000", lightblue: "#B4D8E7" };
  const cellStyle = (cell) => cell.color ? `background:${colorMap[cell.color]};-webkit-print-color-adjust:exact;print-color-adjust:exact;` : "";

  // Build header row 2 (group headers) — white bg, black bold text
  const thStyle = 'background:#fff;color:#000;font-weight:bold;';
  let headerRow2 = `<th rowspan="2" style="${thStyle}">工作天</th><th rowspan="2" style="${thStyle}">日期</th>`;
  headerRow2 += `<th colspan="${leaveCols}" rowspan="1" style="${thStyle}">休 假 狀 況</th>`;
  headerRow2 += `<th colspan="2" rowspan="1" style="${thStyle}">特殊記載</th>`;
  headerRow2 += `<th rowspan="2" style="${thStyle}">台中<br>併線</th>`;
  if (data.teamLeaderHeaders.length) headerRow2 += `<th colspan="${data.teamLeaderHeaders.length}" style="${thStyle}">組長</th>`;
  if (data.reliefHeaders.length) headerRow2 += `<th colspan="${data.reliefHeaders.length}" style="${thStyle}">抵休</th>`;
  headerRow2 += `<th style="${thStyle}">軍功/抵休</th>`;
  headerRow2 += `<th style="${thStyle}">晚班/抵休</th>`;
  data.urbanRouteInfos.forEach((info) => { headerRow2 += `<th style="${thStyle}">${info.shortName}</th>`; });

  // Build header row 3 (employee names under group headers) — white bg, black bold text
  let headerRow3 = "";
  for (let i = 0; i < leaveCols; i++) headerRow3 += `<th style="${thStyle}"></th>`;
  headerRow3 += `<th colspan="2" style="${thStyle}"></th>`;
  data.teamLeaderHeaders.forEach((h) => { headerRow3 += `<th style="${thStyle}">${h.name}</th>`; });
  data.reliefHeaders.forEach((h) => { headerRow3 += `<th style="${thStyle}">${h.name}</th>`; });
  headerRow3 += `<th style="${thStyle}">${data.militaryReliefHeader.name}</th>`;
  headerRow3 += `<th style="${thStyle}">${data.eveningReliefHeader.name}</th>`;
  data.urbanRouteInfos.forEach((info) => { headerRow3 += `<th style="${thStyle}">${info.ownerShortName}</th>`; });

  // Build data rows
  let dataRows = "";
  data.rows.forEach((row) => {
    dataRows += "<tr>";
    dataRows += `<td style="text-align:center;font-weight:bold;">${row.workDay}</td>`;
    dataRows += `<td style="white-space:nowrap;font-weight:bold;">${row.date}</td>`;
    // Each leave employee gets their own cell with leave-type color
    for (let i = 0; i < leaveCols; i++) {
      const lv = row.leaveNames[i];
      if (lv) {
        dataRows += `<td style="background:${colorMap[lv.color]};-webkit-print-color-adjust:exact;print-color-adjust:exact;">${lv.name}</td>`;
      } else {
        dataRows += `<td></td>`;
      }
    }
    dataRows += `<td colspan="2">${(row.specialNotes || []).join("、")}</td>`;
    dataRows += `<td>${(row.mergedLineRoutes || []).join("")}</td>`;
    row.teamLeaderCells.forEach((c) => { dataRows += `<td style="${cellStyle(c)}">${c.text}</td>`; });
    row.reliefCells.forEach((c) => { dataRows += `<td style="${cellStyle(c)}">${c.text}</td>`; });
    dataRows += `<td style="${cellStyle(row.militaryCell)}">${row.militaryCell.text}</td>`;
    dataRows += `<td style="${cellStyle(row.eveningCell)}">${row.eveningCell.text}</td>`;
    row.urbanCells.forEach((c) => { dataRows += `<td style="${cellStyle(c)}">${c.text}</td>`; });
    dataRows += "</tr>";
  });

  popup.document.open();
  popup.document.write(`<!DOCTYPE html>
<html lang="zh-Hant">
<head>
  <meta charset="UTF-8">
  <title>${data.title} ${data.subtitle}</title>
  <style>
    * { box-sizing: border-box; }
    body { font-family: "Segoe UI", "Noto Sans TC", "Microsoft JhengHei", sans-serif; margin: 0; padding: 16px; background: #f6f1e8; color: #2f2418; }
    h1 { margin: 0 0 4px; font-size: 18px; }
    .subtitle { margin: 0 0 12px; font-size: 13px; color: #6f6254; }
    table { width: 100%; border-collapse: collapse; background: #fffdf9; font-size: 12px; table-layout: fixed; }
    th, td { border: 1px solid #b0a090; padding: 2px 3px; text-align: center; vertical-align: middle; overflow: hidden; word-break: break-all; }
    th { background: #fff; font-size: 11px; color: #000; font-weight: bold; }
    /* Auto-fit text in cells */
    td .cell-text { display: inline-block; max-width: 100%; white-space: nowrap; }
    .btn-row { margin-bottom: 12px; }
    .btn-row button { padding: 8px 20px; font-size: 14px; cursor: pointer; border: 1px solid #c0a87c; background: #8b6f47; color: #fff; border-radius: 6px; }
    .btn-row button:hover { background: #6e5535; }
    @media print {
      body { padding: 2px 0 0 2px; margin: 0; background: #fff; }
      .btn-row { display: none; }
      table { font-size: 9px; border: 1px solid #b0a090; }
      th, td { padding: 1px 2px; border: 1px solid #b0a090; }
      h1 { font-size: 14px; margin: 0 0 2px; }
      .subtitle { font-size: 10px; margin: 0 0 4px; }
    }
    @page { size: landscape; margin: 6mm; }
  </style>
</head>
<body>
  <div class="btn-row">
    <button onclick="window.print()">🖨️ 列印</button>
  </div>
  <h1>${data.title}   ${data.subtitle}</h1>
  <table>
    <thead>
      <tr>${headerRow2}</tr>
      <tr>${headerRow3}</tr>
    </thead>
    <tbody>${dataRows}
      <tr class="note-row"><td colspan="2" style="font-weight:bold;text-align:center;">備註</td><td colspan="${totalCols - 2}"></td></tr>
    </tbody>
  </table>
  <script>
  (function() {
    // Auto-fit font size in cells: shrink text to fit cell width
    function autoFitCells() {
      var cells = document.querySelectorAll('tbody td, thead th');
      cells.forEach(function(td) {
        var text = td.textContent.trim();
        if (!text || text.length <= 2) return;
        var cellW = td.clientWidth - 4;
        if (cellW <= 0) return;
        // Start from current font size and shrink if needed
        var style = window.getComputedStyle(td);
        var fontSize = parseFloat(style.fontSize);
        var originalSize = fontSize;
        // Create temp span to measure
        var span = document.createElement('span');
        span.style.visibility = 'hidden';
        span.style.position = 'absolute';
        span.style.whiteSpace = 'nowrap';
        span.style.fontFamily = style.fontFamily;
        span.style.fontWeight = style.fontWeight;
        span.textContent = text;
        document.body.appendChild(span);
        while (fontSize > 6) {
          span.style.fontSize = fontSize + 'px';
          if (span.offsetWidth <= cellW) break;
          fontSize -= 0.5;
        }
        document.body.removeChild(span);
        if (fontSize < originalSize) {
          td.style.fontSize = fontSize + 'px';
        }
      });
    }

    // Distribute row heights evenly to fill the page
    function distributeRowHeights() {
      var table = document.querySelector('table');
      var h1 = document.querySelector('h1');
      var subtitle = document.querySelector('.subtitle');
      // Total page height for landscape A4 ≈ 710px at 96dpi with 5mm margins
      var pageH = 710;
      var usedH = (h1 ? h1.offsetHeight : 0) + (subtitle ? subtitle.offsetHeight + 8 : 0);
      var thead = table.querySelector('thead');
      var theadH = thead ? thead.offsetHeight : 0;
      var tbody = table.querySelector('tbody');
      var dataRows = tbody.querySelectorAll('tr:not(.note-row)');
      var noteRow = tbody.querySelector('tr.note-row');
      var noteH = 20;
      var availH = pageH - usedH - theadH - noteH;
      if (dataRows.length > 0 && availH > 0) {
        var rowH = Math.floor(availH / dataRows.length);
        if (rowH < 18) rowH = 18; // minimum
        dataRows.forEach(function(row) {
          row.style.height = rowH + 'px';
        });
      }
      if (noteRow) noteRow.style.height = noteH + 'px';
    }

    setTimeout(function() {
      autoFitCells();
      distributeRowHeights();
    }, 100);
    window.addEventListener('beforeprint', function() {
      autoFitCells();
      distributeRowHeights();
    });
  })();
  </script>
</body>
</html>`);
  popup.document.close();
  popup.focus();
}

function exportMonthlyScheduleExcel(startDate, endDate) {
  if (typeof XLSX === "undefined") {
    window.alert("Excel 匯出套件尚未載入，請重新整理頁面後再試。");
    return;
  }

  const data = buildMonthlyExportData(startDate, endDate);
  const wb = XLSX.utils.book_new();

  const headerStyle = { font: { bold: true, sz: 11, color: { rgb: "000000" } }, alignment: { horizontal: "center", vertical: "center" }, border: { top: { style: "thin" }, bottom: { style: "thin" }, left: { style: "thin" }, right: { style: "thin" } }, fill: { fgColor: { rgb: "FFFFFF" } } };
  const titleStyle = { font: { bold: true, sz: 14 }, alignment: { horizontal: "center", vertical: "center" } };
  const cellBorder = { border: { top: { style: "thin" }, bottom: { style: "thin" }, left: { style: "thin" }, right: { style: "thin" } }, alignment: { horizontal: "center", vertical: "center", wrapText: true }, font: { sz: 10 } };

  const colorFills = {
    green: { ...cellBorder, fill: { fgColor: { rgb: "92D050" } } },
    yellow: { ...cellBorder, fill: { fgColor: { rgb: "FFFF00" } } },
    orange: { ...cellBorder, fill: { fgColor: { rgb: "FFC000" } } },
    lightblue: { ...cellBorder, fill: { fgColor: { rgb: "B4D8E7" } } },
  };

  // Calculate total columns
  const tlCount = data.teamLeaderHeaders.length;
  const rsCount = data.reliefHeaders.length;
  const urbanCount = data.urbanRouteInfos.length;
  const leaveCols = data.maxLeaveCount;
  const totalCols = 2 + leaveCols + 2 + 1 + tlCount + rsCount + 1 + 1 + urbanCount;
  // Row 0: Title
  const row0 = [{ v: `${data.title}　${data.subtitle}`, s: titleStyle }];
  for (let i = 1; i < totalCols; i++) row0.push({ v: "", s: titleStyle });

  // Row 1: Group headers
  const row1 = [
    { v: "工作天", s: headerStyle },
    { v: "日期", s: headerStyle },
  ];
  // 休假狀況 spans leaveCols columns
  row1.push({ v: "休假狀況", s: headerStyle });
  for (let i = 1; i < leaveCols; i++) row1.push({ v: "", s: headerStyle });
  row1.push({ v: "特殊記載", s: headerStyle });
  row1.push({ v: "", s: headerStyle });
  row1.push({ v: "台中併線", s: headerStyle });
  data.teamLeaderHeaders.forEach((h) => row1.push({ v: "組長", s: headerStyle }));
  data.reliefHeaders.forEach((h) => row1.push({ v: "抵休", s: headerStyle }));
  row1.push({ v: "軍功/抵休", s: headerStyle });
  row1.push({ v: "晚班/抵休", s: headerStyle });
  data.urbanRouteInfos.forEach((info) => row1.push({ v: info.shortName, s: headerStyle }));

  // Row 2: Employee names
  const row2 = [
    { v: "", s: headerStyle },
    { v: "", s: headerStyle },
  ];
  for (let i = 0; i < leaveCols; i++) row2.push({ v: "", s: headerStyle });
  row2.push({ v: "", s: headerStyle });
  row2.push({ v: "", s: headerStyle });
  row2.push({ v: "", s: headerStyle });
  data.teamLeaderHeaders.forEach((h) => row2.push({ v: h.name, s: headerStyle }));
  data.reliefHeaders.forEach((h) => row2.push({ v: h.name, s: headerStyle }));
  row2.push({ v: data.militaryReliefHeader.name, s: headerStyle });
  row2.push({ v: data.eveningReliefHeader.name, s: headerStyle });
  data.urbanRouteInfos.forEach((info) => row2.push({ v: info.ownerShortName, s: headerStyle }));

  // Data rows
  const allRows = [row0, row1, row2];
  data.rows.forEach((row) => {
    const r = [
      { v: row.workDay, s: { ...cellBorder, font: { bold: true, sz: 10 } } },
      { v: row.date, s: { ...cellBorder, font: { bold: true, sz: 10 } } },
    ];
    // Each leave employee gets their own cell with leave-type color
    for (let i = 0; i < leaveCols; i++) {
      const lv = row.leaveNames[i];
      if (lv) {
        r.push({ v: lv.name, s: lv.color ? colorFills[lv.color] : cellBorder });
      } else {
        r.push({ v: "", s: cellBorder });
      }
    }
    r.push({ v: (row.specialNotes || []).join("、"), s: { ...cellBorder, font: { sz: 9 } } });
    r.push({ v: "", s: { ...cellBorder, font: { sz: 9 } } });
    r.push({ v: (row.mergedLineRoutes || []).join(""), s: { ...cellBorder, font: { sz: 9 } } });
    row.teamLeaderCells.forEach((c) => r.push({ v: c.text, s: c.color ? colorFills[c.color] : cellBorder }));
    row.reliefCells.forEach((c) => r.push({ v: c.text, s: c.color ? colorFills[c.color] : cellBorder }));
    r.push({ v: row.militaryCell.text, s: row.militaryCell.color ? colorFills[row.militaryCell.color] : cellBorder });
    r.push({ v: row.eveningCell.text, s: row.eveningCell.color ? colorFills[row.eveningCell.color] : cellBorder });
    row.urbanCells.forEach((c) => r.push({ v: c.text, s: c.color ? colorFills[c.color] : cellBorder }));
    allRows.push(r);
  });

  const ws = XLSX.utils.aoa_to_sheet(allRows);

  // Merges: title row spans all columns
  ws["!merges"] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: totalCols - 1 } },
    // Row 1-2: merge 工作天, 日期 vertically
    { s: { r: 1, c: 0 }, e: { r: 2, c: 0 } },
    { s: { r: 1, c: 1 }, e: { r: 2, c: 1 } },
    // 休假狀況 spans leaveCols columns in row 1, and spans rows 1-2 for each sub-column
    { s: { r: 1, c: 2 }, e: { r: 1, c: 2 + leaveCols - 1 } },
  ];
  // Each leave sub-column header spans row 2 (empty, no merge needed individually)
  const snCol = 2 + leaveCols; // 特殊記載 start column
  // 特殊記載 spans 2 columns in row 1
  ws["!merges"].push({ s: { r: 1, c: snCol }, e: { r: 1, c: snCol + 1 } });
  const mlCol = snCol + 2; // 台中併線 column
  // 台中併線 spans rows 1-2
  ws["!merges"].push({ s: { r: 1, c: mlCol }, e: { r: 2, c: mlCol } });
  const colOffset = mlCol + 1; // after 工作天, 日期, 休假狀況xN, 特殊記載x2, 併線
  // Merge 組長 header if 2+ columns
  if (tlCount > 1) {
    ws["!merges"].push({ s: { r: 1, c: colOffset }, e: { r: 1, c: colOffset + tlCount - 1 } });
  }
  // Merge 抵休 header if 2+ columns
  if (rsCount > 1) {
    ws["!merges"].push({ s: { r: 1, c: colOffset + tlCount }, e: { r: 1, c: colOffset + tlCount + rsCount - 1 } });
  }

  // Column widths — auto-fit to landscape A4
  // Target total width ≈ 170 chars for landscape A4
  // Fixed columns: 工作天(5) + 日期(8) + 特殊記載x2(8+8) + 併線(5) = 34
  // Leave cols + remaining cols share the rest
  const remainingCols = tlCount + rsCount + 2 + urbanCount;
  const leaveW = Math.max(5, Math.min(8, Math.floor(50 / leaveCols)));  // each leave col
  const otherW = Math.max(5, Math.min(13, Math.floor((170 - 34 - leaveCols * leaveW) / remainingCols)));
  ws["!cols"] = [
    { wch: 5 },      // 工作天
    { wch: 8 },      // 日期
  ];
  for (let i = 0; i < leaveCols; i++) ws["!cols"].push({ wch: leaveW });
  ws["!cols"].push({ wch: 8 });   // 特殊記載1
  ws["!cols"].push({ wch: 8 });   // 特殊記載2
  ws["!cols"].push({ wch: 5 });   // 台中併線
  for (let i = 0; i < remainingCols; i++) ws["!cols"].push({ wch: otherW });

  // Row heights (matching user-adjusted layout)
  ws["!rows"] = [
    { hpt: 28.05 },  // Row 0: Title
    { hpt: 19.8 },   // Row 1: Group headers
    { hpt: 19.8 },   // Row 2: Employee names
  ];
  for (let i = 0; i < data.rows.length; i++) ws["!rows"].push({ hpt: 19.95 });

  // Add 備註 row at the end
  const noteRowIdx = allRows.length;
  const noteRow = [{ v: "備註", s: { ...cellBorder, font: { bold: true, sz: 10 } } }, { v: "", s: cellBorder }];
  for (let i = 2; i < totalCols; i++) noteRow.push({ v: "", s: cellBorder });
  allRows.push(noteRow);
  // Re-create sheet with updated allRows
  const wsUpdated = XLSX.utils.aoa_to_sheet(allRows);
  // Copy merges and cols/rows to updated sheet
  wsUpdated["!merges"] = [
    ...ws["!merges"],
    { s: { r: noteRowIdx, c: 0 }, e: { r: noteRowIdx, c: 1 } },       // 備註 label spans A-B
    { s: { r: noteRowIdx, c: 2 }, e: { r: noteRowIdx, c: totalCols - 1 } }, // 備註 content spans C-S
  ];
  wsUpdated["!cols"] = ws["!cols"];
  ws["!rows"].push({ hpt: 24 }); // 備註 row height
  wsUpdated["!rows"] = ws["!rows"];

  XLSX.utils.book_append_sheet(wb, wsUpdated, "排班表");

  const fileName = `${data.title.replace(/\s/g, "")}_抵休代班排程表.xlsx`;
  XLSX.writeFile(wb, fileName);
}

function openRangeBoardWindow(startDate, endDate) {
  const allDates = enumerateDates(startDate, endDate);
  const popup = window.open("", `range-board-${startDate}-${endDate}`, "width=1600,height=900,scrollbars=yes,resizable=yes");
  if (!popup) {
    window.alert("請先允許瀏覽器開啟彈出視窗，才能查看班表。");
    return;
  }

  const activeEmployees = state.employees
    .filter((employee) => employee.employmentStatus === "active")
    .sort(compareEmployeesByScheduleOrder);

  const holidayDates = allDates.filter((dateString) => isHoliday(dateString));
  const visibleDates = allDates.filter((dateString) => !isHoliday(dateString));

  const headerCells = visibleDates.map((dateString) => {
    const date = new Date(`${dateString}T00:00:00`);
    const weekday = ["\u65e5", "\u4e00", "\u4e8c", "\u4e09", "\u56db", "\u4e94", "\u516d"][date.getDay()];
    const m = date.getMonth() + 1;
    const d = date.getDate();
    return `<th>${m}/${d}<small>(${weekday})</small></th>`;
  }).join("");

  const rows = activeEmployees.map((employee) => {
    const defaultRoute = getDefaultRoute(employee);
    let hasOverride = false;
    const cells = visibleDates.map((dateString) => {
      const assignment = getDisplayAssignment(employee, dateString);
      const route = assignment ? getRouteById(assignment.routeId) : null;
      const secondaryRoute = assignment?.secondaryRouteId ? getRouteById(assignment.secondaryRouteId) : null;
      const status = assignment ? getLabel("statuses", assignment.status) : "-";
      const leaveType = assignment?.leaveType ? getLabel("leaveTypes", assignment.leaveType) : "";
      const sourceLabel = assignment?.source === "override" ? "\u7570\u52d5" : assignment ? "\u56fa\u5b9a" : "";
      if (assignment?.source === "override") hasOverride = true;
      const cellClass = [assignment?.source === "override" ? "override-cell" : "", secondaryRoute ? "merged-cell" : ""].filter(Boolean).join(" ");
      let routeText;
      if (assignment?.status === "leave") {
        routeText = leaveType || "\u4f11\u5047";
      } else if (secondaryRoute) {
        routeText = `${route?.name || "-"}+${secondaryRoute.name}`;
      } else {
        routeText = route ? route.name : (defaultRoute ? defaultRoute.name : "-");
      }
      const secondaryText = assignment?.status === "leave"
        ? (leaveType ? `${leaveType} / \u4f11\u5047` : "\u4f11\u5047")
        : (secondaryRoute ? `\u4f75\u7dda / ${status}` : status);
      return `
        <td class="${cellClass}">
          <div class="month-route">${routeText}</div>
          <div class="month-meta">${secondaryText}${sourceLabel ? ` / ${sourceLabel}` : ""}</div>
        </td>
      `;
    }).join("");
    return `
      <tr data-has-override="${hasOverride}">
        <th class="sticky-name">
          <div>${employee.name}</div>
          <small>${roleLabels[employee.role] || "-"}</small>
        </th>
        ${cells}
      </tr>
    `;
  }).join("");

  const overrideCount = activeEmployees.filter((emp) => {
    return visibleDates.some((d) => {
      const a = getDisplayAssignment(emp, d);
      return a?.source === "override";
    });
  }).length;

  popup.document.open();
  popup.document.write(`
    <!DOCTYPE html>
    <html lang="zh-Hant">
    <head>
      <meta charset="UTF-8">
      <title>\u73ed\u8868 ${startDate} ~ ${endDate}</title>
      <style>
        body { font-family: "Segoe UI", "Noto Sans TC", sans-serif; margin: 0; padding: 24px; background: #f6f1e8; color: #2f2418; }
        h1 { margin: 0 0 8px; }
        p { margin: 0 0 18px; color: #6f6254; }
        .meta { display: flex; gap: 10px; flex-wrap: wrap; margin-bottom: 12px; }
        .meta span { background: #fff8ef; border: 1px solid #ead8c2; border-radius: 999px; padding: 8px 12px; font-weight: 600; }
        .filter-bar { margin-bottom: 18px; }
        .filter-bar label { cursor: pointer; font-weight: 600; user-select: none; }
        .filter-bar input { margin-right: 6px; transform: scale(1.2); }
        .board-wrap { overflow: auto; border: 1px solid #eadfd0; background: #fffdf9; max-height: calc(100vh - 160px); }
        table { border-collapse: collapse; width: 100%; table-layout: fixed; }
        th, td { border-bottom: 1px solid #eadfd0; border-right: 1px solid #f0e5d7; padding: 6px 6px; text-align: left; vertical-align: top; }
        thead th { position: sticky; top: 0; background: #f1e3d1; z-index: 2; min-width: 70px; font-size: 12px; }
        thead th small { display: block; color: #6f6254; margin-top: 4px; }
        .sticky-name { position: sticky; left: 0; background: #fff8ef; z-index: 1; min-width: 90px; width: 90px; }
        thead th.sticky-name { z-index: 3; }
        .sticky-name small { display: block; color: #6f6254; margin-top: 4px; }
        .override-cell { background: #fff0e8; }
        .merged-cell { background: #fce8f0; }
        .month-route { font-weight: 700; margin-bottom: 2px; font-size: 12px; line-height: 1.25; word-break: break-word; }
        .month-meta { color: #6f6254; font-size: 11px; line-height: 1.25; word-break: break-word; }
        tr.hidden-row { display: none; }
        @media print {
          body { padding: 8px; background: #fff; }
          .meta { margin-bottom: 8px; }
          .meta span { padding: 4px 8px; font-size: 11px; }
          .filter-bar { display: none; }
          th, td { padding: 4px 4px; }
          .month-route { font-size: 10px; }
          .month-meta { font-size: 9px; }
          tr.hidden-row { display: none; }
        }
      </style>
    </head>
    <body>
      <h1>\u73ed\u8868 ${startDate} ~ ${endDate}</h1>
      <p>\u5df2\u81ea\u52d5\u96b1\u85cf\u9031\u672b\u8207\u516c\u53f8\u4f11\u5047\u65e5\u6b04\u4f4d\uff0c\u6bcf\u683c\u986f\u793a\u8def\u7dda\u3001\u72c0\u614b\u3001\u5047\u5225\u3002</p>
      <div class="meta">
        <span>\u54e1\u5de5 ${activeEmployees.length} \u4eba</span>
        <span>\u986f\u793a\u5de5\u4f5c\u65e5 ${visibleDates.length} \u5929</span>
        <span>\u96b1\u85cf\u5047\u65e5 ${holidayDates.length} \u5929</span>
        <span>\u6709\u7570\u52d5\u54e1\u5de5 ${overrideCount} \u4eba</span>
      </div>
      <div class="filter-bar">
        <label><input type="checkbox" id="filterOverrideOnly">\u53ea\u986f\u793a\u6709\u7570\u52d5\uff08\u4f11\u5047/\u4ee3\u73ed/\u652f\u63f4\uff09\u7684\u54e1\u5de5</label>
      </div>
      <div class="board-wrap">
        <table>
          <thead>
            <tr>
              <th class="sticky-name">\u54e1\u5de5</th>
              ${headerCells}
            </tr>
          </thead>
          <tbody id="boardBody">${rows}</tbody>
        </table>
      </div>
      <script>
        document.getElementById("filterOverrideOnly").addEventListener("change", function() {
          var rows = document.querySelectorAll("#boardBody tr");
          for (var i = 0; i < rows.length; i++) {
            if (this.checked && rows[i].getAttribute("data-has-override") === "false") {
              rows[i].classList.add("hidden-row");
            } else {
              rows[i].classList.remove("hidden-row");
            }
          }
        });
      </script>
    </body>
    </html>
  `);
  popup.document.close();
  popup.focus();
}

function generateDefaultAssignments(formData, currentUser) {
  const payload = Object.fromEntries(formData.entries());
  if (payload.startDate > payload.endDate) {
    window.alert("結束日期不可早於開始日期。");
    return false;
  }

  const targetEmployees = payload.employeeId
    ? state.employees.filter((employee) => employee.id === payload.employeeId && employee.defaultRouteId)
    : state.employees.filter((employee) => employee.defaultRouteId);

  if (!targetEmployees.length) {
    window.alert("沒有可生成固定班表的人員。");
    return false;
  }

  const dates = getWorkingDates(payload.startDate, payload.endDate, state.companySettings);
  let created = 0;
  let updated = 0;
  let preservedOverrides = 0;

  dates.forEach((dateString) => {
    targetEmployees.forEach((employee) => {
      const existing = getAssignmentByEmployeeDate(employee.id, dateString);
      if (existing?.source === "override") {
        preservedOverrides += 1;
        return;
      }
      if (existing) {
        existing.routeId = employee.defaultRouteId;
        existing.shift = employee.shift;
        existing.status = "working";
        existing.leaveType = "";
        existing.note = "";
        existing.source = "default";
        updated += 1;
      } else {
        state.assignments.push({
          id: makeId("asg"),
          date: dateString,
          employeeId: employee.id,
          routeId: employee.defaultRouteId,
          shift: employee.shift,
          status: "working",
          leaveType: "",
          note: "",
          source: "default",
        });
        created += 1;
      }
    });
  });

  logAction({
    actorId: currentUser.id,
    action: "default-generate",
    targetType: "assignment",
    targetId: payload.employeeId || "all",
    summary: `生成固定班表 ${payload.startDate} 至 ${payload.endDate}`,
    detail: `新增 ${created} 筆 / 更新 ${updated} 筆 / 保留異動覆蓋 ${preservedOverrides} 筆`,
  });

  state.session.lastGeneratedRange = { startDate: payload.startDate, endDate: payload.endDate };
  saveState();
  window.alert(`固定班表已生成。\n新增 ${created} 筆\n更新 ${updated} 筆\n保留異動覆蓋 ${preservedOverrides} 筆`);
  return true;
}
function applyNextMonthRange(formEl) {
  const nextMonth = getMonthRange(getToday(), 1);
  formEl.elements.startDate.value = nextMonth.startDate;
  formEl.elements.endDate.value = nextMonth.endDate;
  window.alert(`已帶入次月日期區間：${nextMonth.startDate} 至 ${nextMonth.endDate}`);
}

function upsertAssignmentFromForm(formData, currentUser) {
  const payload = Object.fromEntries(formData.entries());
  const route = getRouteById(payload.routeId);
  const employee = getEmployeeById(payload.employeeId);
  if (!route || !employee) return false;

  const routeConflict = state.assignments.find((assignment) =>
    assignment.date === payload.date &&
    assignment.routeId === payload.routeId &&
    assignment.employeeId !== payload.employeeId &&
    assignment.status !== "leave"
  );

  if (routeConflict) {
    const conflictEmployee = getEmployeeById(routeConflict.employeeId);
    const confirmed = window.confirm(
      `${payload.date} 的 ${route?.name || "未指定路線"} 已排給 ${conflictEmployee?.name || routeConflict.employeeId}\n` +
      `既有排班：${describeAssignment(routeConflict)}\n\n是否仍要保留重複路線排班？`
    );
    if (!confirmed) return false;
  }

  const existing = getAssignmentByEmployeeDate(payload.employeeId, payload.date);
  const shift = payload.shift || inferShift(route.name);
  if (existing) {
    existing.routeId = payload.routeId;
    existing.shift = shift;
    existing.status = payload.status;
    existing.leaveType = payload.status === "leave" ? payload.leaveType : "";
    existing.note = (payload.note || "").trim();
    existing.source = "override";
    logAction({
      actorId: currentUser.id,
      action: "assignment-update",
      targetType: "assignment",
      targetId: existing.id,
      summary: `更新排班 ${employee.name} ${payload.date}`,
      detail: describeAssignment(existing),
    });
  } else {
    const assignment = {
      id: makeId("asg"),
      date: payload.date,
      employeeId: payload.employeeId,
      routeId: payload.routeId,
      shift,
      status: payload.status,
      leaveType: payload.status === "leave" ? payload.leaveType : "",
      note: (payload.note || "").trim(),
      source: "override",
    };
    state.assignments.push(assignment);
    logAction({
      actorId: currentUser.id,
      action: "assignment-create",
      targetType: "assignment",
      targetId: assignment.id,
      summary: `新增排班 ${employee.name} ${payload.date}`,
      detail: describeAssignment(assignment),
    });
  }

  saveState();
  return true;
}
function applyLeaveOnly(formData, currentUser) {
  const payload = Object.fromEntries(formData.entries());
  const employee = getEmployeeById(payload.employeeId);
  if (!employee) return false;

  const defaultRoute = getDefaultRoute(employee);
  const startDate = payload.startDate;
  const endDate = payload.endDate;
  if (!startDate || !endDate) {
    window.alert("請選擇開始與結束日期。");
    return false;
  }
  if (startDate > endDate) {
    window.alert("結束日期不可早於開始日期。");
    return false;
  }

  const leaveRouteId = defaultRoute ? defaultRoute.id : "";
  const shift = defaultRoute ? employee.shift : "day";
  const note = (payload.note || "").trim();
  const dates = enumerateDates(startDate, endDate);
  let processed = 0;
  let skippedHoliday = 0;

  dates.forEach((dateString) => {
    if (isHoliday(dateString)) {
      skippedHoliday += 1;
      return;
    }

    let leaveAssignment = getAssignmentByEmployeeDate(employee.id, dateString);
    if (!leaveAssignment) {
      leaveAssignment = {
        id: makeId("asg"),
        date: dateString,
        employeeId: employee.id,
        routeId: leaveRouteId,
        shift,
        status: "leave",
        leaveType: payload.leaveType,
        note,
        source: "override",
      };
      state.assignments.push(leaveAssignment);
    } else {
      leaveAssignment.routeId = leaveRouteId;
      leaveAssignment.shift = shift;
      leaveAssignment.status = "leave";
      leaveAssignment.leaveType = payload.leaveType;
      leaveAssignment.note = note;
      leaveAssignment.source = "override";
    }

    processed += 1;
  });

  if (!processed) {
    window.alert("選擇區間都落在週末或公司休假日，沒有可套用的工作日。");
    return false;
  }

  logAction({
    actorId: currentUser.id,
    action: "leave",
    targetType: "assignment",
    targetId: employee.id,
    summary: `${employee.name} ${startDate} 至 ${endDate} 登記 ${getLabel("leaveTypes", payload.leaveType)}`,
    detail: `套用 ${processed} 天${skippedHoliday ? `，略過假日 ${skippedHoliday} 天` : ""}。`,
  });

  saveState();
  window.alert(`請假已儲存。\n套用工作日 ${processed} 天${skippedHoliday ? `\n略過假日 ${skippedHoliday} 天` : ""}`);
  return true;
}

function applyReliefOnly(formData, currentUser) {
  const payload = Object.fromEntries(formData.entries());
  const reliefEmployee = getEmployeeById(payload.reliefEmployeeId);
  if (!reliefEmployee) { window.alert("請選擇代班人員。"); return false; }

  const startDate = payload.startDate;
  const endDate = payload.endDate;
  if (!startDate || !endDate) {
    window.alert("請選擇開始與結束日期。");
    return false;
  }
  if (startDate > endDate) {
    window.alert("結束日期不可早於開始日期。");
    return false;
  }

  const route = payload.routeId ? getRouteById(payload.routeId) : null;
  if (!route) { window.alert("請選擇代班路線。"); return false; }
  const isMergedLine = payload.isMergedLine === "yes";
  const secondaryRoute = (isMergedLine && payload.secondaryRouteId) ? getRouteById(payload.secondaryRouteId) : null;
  const reliefRouteId = route.id;
  const secondaryRouteId = secondaryRoute ? secondaryRoute.id : "";
  const isMerged = !!secondaryRouteId;
  const shift = inferShift(route.name);
  const specialNote = (payload.specialNote || "").trim();
  let note = (payload.note || "").trim() || "代班";
  if (isMerged && !note.includes("併線")) {
    note += "（併線）";
  }
  const dates = enumerateDates(startDate, endDate);
  let processed = 0;
  let skippedHoliday = 0;

  dates.forEach((dateString) => {
    if (isHoliday(dateString)) {
      skippedHoliday += 1;
      return;
    }

    const reliefExisting = getAssignmentByEmployeeDate(reliefEmployee.id, dateString);
    if (reliefExisting) {
      reliefExisting.routeId = reliefRouteId;
      reliefExisting.secondaryRouteId = secondaryRouteId;
      reliefExisting.isMergedLine = isMergedLine;
      reliefExisting.shift = shift;
      reliefExisting.status = "reassigned";
      reliefExisting.leaveType = "";
      reliefExisting.note = note;
      reliefExisting.specialNote = specialNote;
      reliefExisting.source = "override";
    } else {
      state.assignments.push({
        id: makeId("asg"),
        date: dateString,
        employeeId: reliefEmployee.id,
        routeId: reliefRouteId,
        secondaryRouteId,
        isMergedLine,
        shift,
        status: "reassigned",
        leaveType: "",
        note,
        specialNote,
        source: "override",
      });
    }

    processed += 1;
  });

  if (!processed) {
    window.alert("選擇區間都落在週末或公司休假日，沒有可套用的工作日。");
    return false;
  }

  const mergedInfo = isMerged ? `（併線：上午 ${route.name} / 下午 ${secondaryRoute.name}）` : "";
  logAction({
    actorId: currentUser.id,
    action: "relief",
    targetType: "assignment",
    targetId: reliefEmployee.id,
    summary: `${reliefEmployee.name} 代班，${startDate} 至 ${endDate}${isMergedLine ? "（併線）" : ""}`,
    detail: `路線 ${route.name}${mergedInfo}。套用 ${processed} 天${skippedHoliday ? `，略過假日 ${skippedHoliday} 天` : ""}。${specialNote ? `特殊記載：${specialNote}` : ""}`,
  });

  saveState();
  window.alert(`代班已儲存。\n${reliefEmployee.name} 代班\n路線：${route.name}\n套用工作日 ${processed} 天${skippedHoliday ? `\n略過假日 ${skippedHoliday} 天` : ""}`);
  return true;
}

function renderMasterDataPanel(currentUser) {
  const section = createSection("基本資料與休假日", "行政與主管可維護公司休假日與顯示名稱。");
  const grid = document.createElement("div");
  grid.className = "panel-grid master-grid";

  const holidayPanel = document.createElement("div");
  holidayPanel.className = "card panel holiday-settings master-card";
  holidayPanel.innerHTML = `
    <div class="section-heading">
      <div>
        <h3>休假日設定</h3>
        <p class="muted">週六、週日固定休假；可維護公司加休或補休日期。</p>
      </div>
    </div>
    <div class="notice">目前公司休假日清單（標示「手動」為另外維護）：</div>
    <div class="inline-list holiday-chip-list">${renderCompanyHolidayList()}</div>
    <form id="holidayForm" class="form-grid holiday-form-grid">
      <label>公司休假日（每行一個日期）
        <textarea class="holiday-textarea" name="holidays" rows="8" placeholder="例如：2026-04-04">${state.companySettings.holidays.join("\n")}</textarea>
      </label>
      <button type="submit" class="holiday-save-button">儲存休假日</button>
    </form>
  `;

  const labelsPanel = document.createElement("div");
  labelsPanel.className = "card panel master-card";
  labelsPanel.innerHTML = `
    <div class="section-heading">
      <div>
        <h3>名稱設定</h3>
        <p class="muted">可調整班別、假別、狀態顯示名稱。</p>
      </div>
    </div>
    <form id="labelsForm" class="mini-grid">
      ${renderLabelInputs()}
      <button type="submit">儲存名稱設定</button>
    </form>
  `;

  const employeePanel = document.createElement("div");
  employeePanel.className = "card panel master-card";
  employeePanel.innerHTML = `
    <div class="section-heading">
      <div>
        <h3>員工維護</h3>
        <p class="muted">可新增或編輯員工、角色、固定路線與可支援代班。</p>
      </div>
    </div>
    <form id="employeeForm" class="form-grid">
      <input name="employeeId" type="hidden">
      <label>選擇既有員工
        <select name="existingEmployeeId">
          <option value="">新增員工</option>
          ${[...state.employees].sort(compareEmployeesByScheduleOrder).map((employee) => `<option value="${employee.id}">${employee.name}</option>`).join("")}
        </select>
      </label>
      <div class="action-row">
        <button type="button" class="secondary" id="loadEmployeeButton">載入</button>
        <button type="button" class="ghost" id="resetEmployeeButton">清空</button>
      </div>
      <label>姓名<input name="name" required></label>
      <label>角色<select name="role">${Object.entries(roleLabels).map(([key, label]) => `<option value="${key}">${label}</option>`).join("")}</select></label>
      <label>班別<select name="shift">${labelOptions("shifts")}</select></label>
      <label>固定路線<select name="defaultRouteId"><option value="">未指定路線</option>${routeOptions()}</select></label>
      <label>是否抵休<select name="isRelief"><option value="false">否</option><option value="true">是</option></select></label>
      <label>可支援代班<select name="canCoverShift"><option value="false">否</option><option value="true">是</option></select></label>
      <label>在職狀態
        <select name="employmentStatus">
          <option value="active">在職</option>
          <option value="resigned">已離職</option>
          <option value="unpaidLeave">留職停薪</option>
        </select>
      </label>
      <label>支援路線（可多選）<select name="supportLineIds" multiple>${sortedRoutes().map((route) => `<option value="${route.id}">${route.name}</option>`).join("")}</select></label>
      <div class="action-row">
        <button type="submit">儲存員工</button>
        <button type="button" class="ghost" id="deleteEmployeeButton" style="color:#c0392b;border-color:#c0392b;">刪除員工</button>
      </div>
    </form>
  `;

  const routePanel = document.createElement("div");
  routePanel.className = "card panel master-card";
  routePanel.innerHTML = `
    <div class="section-heading">
      <div>
        <h3>路線維護</h3>
        <p class="muted">可新增或編輯路線名稱、類型與核定里程。</p>
      </div>
    </div>
    <form id="routeForm" class="form-grid">
      <input name="routeId" type="hidden">
      <label>選擇既有路線
        <select name="existingRouteId">
          <option value="">新增路線</option>
          ${sortedRoutes().map((route) => `<option value="${route.id}">${route.name}</option>`).join("")}
        </select>
      </label>
      <div class="action-row">
        <button type="button" class="secondary" id="loadRouteButton">載入</button>
        <button type="button" class="ghost" id="resetRouteButton">清空</button>
      </div>
      <label>路線名稱<input name="name" required></label>
      <label>路線類型<select name="type">${Object.entries(routeTypeLabels).map(([key, label]) => `<option value="${key}">${label}</option>`).join("")}</select></label>
      <label>核定里程<input name="approvedMileage" type="number" min="0" step="1" required></label>
      <button type="submit">儲存路線</button>
    </form>
  `;

  const pinPanel = document.createElement("div");
  if (currentUser.role === "supervisor") {
    // Build individual PIN inputs for protected role employees
    const individualPinInputs = state.employees
      .filter((emp) => protectedRoles.includes(emp.role) && emp.employmentStatus === "active")
      .sort(compareEmployeesByScheduleOrder)
      .map((emp) => {
        const individualPin = (state.pinSettings.individual && state.pinSettings.individual[emp.id]) || "";
        return `<label>${emp.name}（${roleLabels[emp.role]}）<input name="individual_${emp.id}" type="password" value="${individualPin}" maxlength="8" placeholder="未設定則用角色預設"></label>`;
      })
      .join("");

    pinPanel.className = "card panel master-card";
    pinPanel.innerHTML = `
      <div class="section-heading">
        <div>
          <h3>PIN 碼管理</h3>
          <p class="muted">可為每位管理人員設定個人 PIN 碼。未設定個人 PIN 的人員使用角色預設 PIN。</p>
        </div>
      </div>
      <form id="pinForm" class="form-grid">
        <div class="card stat-card">
          <p class="stat-label">角色預設 PIN（通用）</p>
          <label>組長預設 PIN<input name="teamLeader" type="password" value="${state.pinSettings.teamLeader}" maxlength="8" required></label>
          <label>行政預設 PIN<input name="adminStaff" type="password" value="${state.pinSettings.adminStaff}" maxlength="8" required></label>
          <label>主管預設 PIN<input name="supervisor" type="password" value="${state.pinSettings.supervisor}" maxlength="8" required></label>
        </div>
        <div class="card stat-card">
          <p class="stat-label">個人 PIN（優先於角色預設）</p>
          ${individualPinInputs}
        </div>
        <button type="submit">儲存 PIN 碼</button>
      </form>
    `;
  }

  const panels = [holidayPanel, labelsPanel, employeePanel, routePanel];
  if (currentUser.role === "supervisor") panels.push(pinPanel);
  grid.append(...panels);
  section.appendChild(grid);

  const employeeForm = employeePanel.querySelector("#employeeForm");
  employeeForm.addEventListener("submit", (event) => {
    event.preventDefault();
    createEmployee(new FormData(event.currentTarget), currentUser);
    render();
  });
  employeePanel.querySelector("#loadEmployeeButton").addEventListener("click", () => populateEmployeeForm(employeeForm, employeeForm.elements.existingEmployeeId.value));
  employeePanel.querySelector("#resetEmployeeButton").addEventListener("click", () => {
    employeeForm.reset();
    employeeForm.elements.employeeId.value = "";
  });
  employeePanel.querySelector("#deleteEmployeeButton").addEventListener("click", () => {
    const empId = employeeForm.elements.employeeId.value || employeeForm.elements.existingEmployeeId.value;
    if (!empId) { window.alert("請先選擇並載入一位員工，才能執行刪除。"); return; }
    const emp = state.employees.find((e) => e.id === empId);
    if (!emp) { window.alert("找不到該員工。"); return; }
    if (!window.confirm(`確定要刪除員工「${emp.name}」嗎？\n\n此操作會一併刪除該員工所有的班表異動紀錄，且無法復原。`)) return;
    // Remove employee
    state.employees = state.employees.filter((e) => e.id !== empId);
    // Remove related assignments
    state.assignments = state.assignments.filter((a) => a.employeeId !== empId);
    // Audit log
    state.auditLog.push({
      timestamp: new Date().toISOString(),
      user: currentUser.name,
      role: currentUser.role,
      action: "deleteEmployee",
      summary: `刪除員工 ${emp.name}`,
    });
    // Extend echo delay to prevent Firebase listener from restoring deleted data
    lastFirebaseSaveTime = Date.now() + 5000;
    saveState();
    render();
    window.alert(`已刪除員工「${emp.name}」。`);
  });
  const empDetails = document.createElement("details");
  empDetails.className = "collapsible-list";
  const empSummary = document.createElement("summary");
  const empCount = state.employees.filter(e => (e.employmentStatus || "active") === "active").length;
  empSummary.textContent = `目前員工清單（${empCount} 人）`;
  empDetails.appendChild(empSummary);
  empDetails.appendChild(renderEmployeeTable());
  employeePanel.appendChild(empDetails);

  const routeForm = routePanel.querySelector("#routeForm");
  routeForm.addEventListener("submit", (event) => {
    event.preventDefault();
    createRoute(new FormData(event.currentTarget), currentUser);
    render();
  });
  routePanel.querySelector("#loadRouteButton").addEventListener("click", () => populateRouteForm(routeForm, routeForm.elements.existingRouteId.value));
  routePanel.querySelector("#resetRouteButton").addEventListener("click", () => {
    routeForm.reset();
    routeForm.elements.routeId.value = "";
  });
  const routeDetails = document.createElement("details");
  routeDetails.className = "collapsible-list";
  const routeSummary = document.createElement("summary");
  routeSummary.textContent = `目前路線清單（${state.routes.length} 條）`;
  routeDetails.appendChild(routeSummary);
  routeDetails.appendChild(renderRouteTable());
  routePanel.appendChild(routeDetails);

  labelsPanel.querySelector("#labelsForm").addEventListener("submit", (event) => {
    event.preventDefault();
    updateLabels(new FormData(event.currentTarget), currentUser);
    render();
  });

  holidayPanel.querySelector("#holidayForm").addEventListener("submit", (event) => {
    event.preventDefault();
    updateCompanyHolidays(new FormData(event.currentTarget), currentUser);
    render();
  });

  if (currentUser.role === "supervisor" && pinPanel.querySelector("#pinForm")) {
    pinPanel.querySelector("#pinForm").addEventListener("submit", (event) => {
      event.preventDefault();
      const fd = new FormData(event.currentTarget);
      state.pinSettings.teamLeader = fd.get("teamLeader").trim() || "1234";
      state.pinSettings.adminStaff = fd.get("adminStaff").trim() || "1234";
      state.pinSettings.supervisor = fd.get("supervisor").trim() || "0000";
      // Save individual PINs
      if (!state.pinSettings.individual) state.pinSettings.individual = {};
      for (const [key, value] of fd.entries()) {
        if (key.startsWith("individual_")) {
          const empId = key.replace("individual_", "");
          const pin = value.trim();
          if (pin) {
            state.pinSettings.individual[empId] = pin;
          } else {
            delete state.pinSettings.individual[empId];
          }
        }
      }
      logAction({
        actorId: currentUser.id,
        action: "pin-update",
        targetType: "settings",
        targetId: "pin-settings",
        summary: "更新 PIN 碼設定",
        detail: "已更新角色預設 PIN 碼與個人 PIN 碼。",
      });
      // Clear authenticated cache so new PINs take effect
      authenticatedUsers.clear();
      authenticatedRoles.clear();
      clearAuthCache();
      saveState();
      window.alert("PIN 碼已更新。下次切換帳號時將使用新的 PIN 碼。");
    });
  }

  return section;
}
function renderEmployeeTable() {
  const wrap = document.createElement("div");
  wrap.className = "table-card";
  const rows = [...state.employees]
    .sort(compareEmployeesByScheduleOrder)
    .map((employee) => {
      const defaultRoute = getDefaultRoute(employee);
      return `
        <tr>
          <td>${employee.name}</td>
          <td>${roleLabels[employee.role]}</td>
          <td>${defaultRoute ? `${defaultRoute.name} / ${getLabel("shifts", employee.shift)}` : "未指定路線"}</td>
          <td>${employee.canCoverShift ? "是" : "否"}</td>
          <td>${employmentStatusLabels[employee.employmentStatus || "active"]}</td>
        </tr>
      `;
    })
    .join("");

  wrap.innerHTML = `
    <table>
      <thead>
        <tr>
          <th>姓名</th>
          <th>角色</th>
          <th>固定配置</th>
          <th>可支援代班</th>
          <th>狀態</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  `;

  return wrap;
}
function renderRouteTable() {
  const wrap = document.createElement("div");
  wrap.className = "table-card";
  const rows = state.routes
    .sort((a, b) => a.name.localeCompare(b.name, "zh-Hant"))
    .map((route) => {
      const owner = state.employees.find((employee) => employee.defaultRouteId === route.id);
      return `
        <tr>
          <td>${route.name}</td>
          <td>${displayRouteType(route.type)}</td>
          <td>${route.approvedMileage} 公里</td>
          <td>${owner ? owner.name : "未指定"}</td>
        </tr>
      `;
    })
    .join("");

  wrap.innerHTML = `
    <table>
      <thead>
        <tr>
          <th>路線名稱</th>
          <th>類型</th>
          <th>核定里程</th>
          <th>固定人員</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  `;

  return wrap;
}
function renderLabelInputs() {
  const groups = [["shifts", "班別名稱"], ["leaveTypes", "假別名稱"], ["statuses", "狀態名稱"]];
  const keyLabels = {
    shifts: { day: "白班", evening: "晚班", night: "大夜班" },
    leaveTypes: { annual: "特休", personal: "事假", sick: "病假", official: "公假", injury: "公傷", other: "其他" },
    statuses: { working: "上班", leave: "休假", reassigned: "代班", standby: "待命" },
  };
  return groups.map(([group, title]) => `
    <div class="card stat-card">
      <p class="stat-label">${title}</p>
      ${Object.entries(state.labelSettings[group]).map(([key, value]) => `<label>${keyLabels[group][key] || key}<input name="${group}.${key}" value="${value}"></label>`).join("")}
    </div>
  `).join("");
}

function populateEmployeeForm(form, employeeId) {
  const employee = getEmployeeById(employeeId);
  if (!employee) return;
  form.elements.employeeId.value = employee.id;
  form.elements.name.value = employee.name;
  form.elements.role.value = employee.role;
  form.elements.shift.value = employee.shift;
  form.elements.defaultRouteId.value = employee.defaultRouteId || "";
  form.elements.isRelief.value = String(employee.isRelief);
  form.elements.canCoverShift.value = String(!!employee.canCoverShift);
  form.elements.employmentStatus.value = employee.employmentStatus || (employee.active ? "active" : "resigned");
  Array.from(form.elements.supportLineIds.options).forEach((option) => {
    option.selected = employee.supportLineIds.includes(option.value);
  });
}

function populateRouteForm(form, routeId) {
  const route = getRouteById(routeId);
  if (!route) return;
  form.elements.routeId.value = route.id;
  form.elements.type.value = route.type;
  form.elements.name.value = route.name;
  form.elements.approvedMileage.value = route.approvedMileage;
}

function createEmployee(formData, currentUser) {
  const supportLineIds = formData.getAll("supportLineIds");
  const defaultRouteId = formData.get("defaultRouteId") || "";
  const payload = {
    name: (formData.get("name") || "").trim(),
    role: formData.get("role"),
    shift: formData.get("shift"),
    defaultRouteId,
    supportLineIds,
    isRelief: formData.get("isRelief") === "true",
    canCoverShift: formData.get("canCoverShift") === "true",
    employmentStatus: formData.get("employmentStatus") || "active",
    active: (formData.get("employmentStatus") || "active") === "active",
    fixedDuty: defaultRouteId ? (getRouteById(defaultRouteId)?.name || "") : "",
    isNightOwner: getRouteById(defaultRouteId)?.name === "大夜班",
  };
  const employeeId = formData.get("employeeId");
  const existing = employeeId ? getEmployeeById(employeeId) : null;
  if (existing) {
    Object.assign(existing, payload);
    logAction({
      actorId: currentUser.id,
      action: "employee-update",
      targetType: "employee",
      targetId: existing.id,
      summary: `更新員工 ${existing.name}`,
      detail: `${roleLabels[existing.role]} / ${existing.defaultRouteId ? getRouteById(existing.defaultRouteId)?.name || "" : "未指定路線"}`,
    });
  } else {
    const employee = { id: makeId("emp"), ...payload };
    state.employees.push(employee);
    logAction({
      actorId: currentUser.id,
      action: "employee-create",
      targetType: "employee",
      targetId: employee.id,
      summary: `新增員工 ${employee.name}`,
      detail: `${roleLabels[employee.role]} / ${employee.defaultRouteId ? getRouteById(employee.defaultRouteId)?.name || "" : "未指定路線"}`,
    });
  }
  saveState();
}
function createRoute(formData, currentUser) {
  const payload = {
    type: formData.get("type"),
    name: formData.get("name").trim(),
    approvedMileage: Number(formData.get("approvedMileage")),
  };
  const routeId = formData.get("routeId");
  const existing = routeId ? getRouteById(routeId) : null;
  if (existing) {
    Object.assign(existing, payload);
    logAction({
      actorId: currentUser.id,
      action: "route-update",
      targetType: "route",
      targetId: existing.id,
      summary: `更新路線 ${existing.name}`,
      detail: `${displayRouteType(existing.type)} / 核定里程 ${existing.approvedMileage} 公里`,
    });
  } else {
    const route = { id: makeId("route"), ...payload };
    state.routes.push(route);
    logAction({
      actorId: currentUser.id,
      action: "route-create",
      targetType: "route",
      targetId: route.id,
      summary: `新增路線 ${route.name}`,
      detail: `${displayRouteType(route.type)} / 核定里程 ${route.approvedMileage} 公里`,
    });
  }
  saveState();
}

function updateLabels(formData, currentUser) {
  for (const [key, value] of formData.entries()) {
    const [group, itemKey] = key.split(".");
    state.labelSettings[group][itemKey] = value.trim();
  }
  logAction({
    actorId: currentUser.id,
    action: "label-update",
    targetType: "settings",
    targetId: "label-settings",
    summary: "更新顯示名稱設定",
    detail: "已更新班別、假別與狀態名稱。",
  });
  saveState();
}

function updateCompanyHolidays(formData, currentUser) {
  state.companySettings.weekendDaysOff = true;
  state.companySettings.holidays = [...new Set(
    formData.get("holidays")
      .split(/\r?\n/)
      .map((value) => value.trim())
      .filter(Boolean)
  )].sort();

  logAction({
    actorId: currentUser.id,
    action: "holiday-update",
    targetType: "settings",
    targetId: "company-holidays",
    summary: "更新公司休假日",
    detail: `週末固定休假 / 公司休假日 ${state.companySettings.holidays.length} 天`,
  });
  saveState();
}

function renderHistoryQueryPanel(currentUser) {
  const section = createSection("歷史班表查詢", "可指定日期區間與員工，查詢過去或未來的班表記錄。");
  const lastMonth = getMonthRange(getToday(), -1);
  section.innerHTML += `
    <form id="historyQueryForm" class="form-grid">
      <label>開始日期<input name="startDate" type="date" value="${lastMonth.startDate}"></label>
      <label>結束日期<input name="endDate" type="date" value="${lastMonth.endDate}"></label>
      <label>員工篩選<select name="employeeId"><option value="">全部員工</option>${employeeOptions({ includeAll: true })}</select></label>
      <label>狀態篩選
        <select name="statusFilter">
          <option value="">全部狀態</option>
          ${Object.entries(state.labelSettings.statuses).map(([k, v]) => `<option value="${k}">${v}</option>`).join("")}
        </select>
      </label>
      <button type="submit">查詢</button>
    </form>
    <div id="historyQueryResult"></div>
  `;

  section.querySelector("#historyQueryForm").addEventListener("submit", (event) => {
    event.preventDefault();
    const fd = new FormData(event.currentTarget);
    const startDate = fd.get("startDate");
    const endDate = fd.get("endDate");
    const employeeId = fd.get("employeeId");
    const statusFilter = fd.get("statusFilter");
    if (!startDate || !endDate || startDate > endDate) {
      window.alert("請確認日期區間正確。");
      return;
    }
    let results = state.assignments.filter((a) => a.date >= startDate && a.date <= endDate);
    if (employeeId) results = results.filter((a) => a.employeeId === employeeId);
    if (statusFilter) results = results.filter((a) => a.status === statusFilter);
    results.sort((a, b) => a.date.localeCompare(b.date) || (getEmployeeById(a.employeeId)?.name || "").localeCompare(getEmployeeById(b.employeeId)?.name || "", "zh-Hant"));

    const container = section.querySelector("#historyQueryResult");
    if (!results.length) {
      container.innerHTML = `<div class="empty-state" style="margin-top:12px;"><p>查無符合條件的班表記錄。</p></div>`;
      return;
    }
    container.innerHTML = "";
    container.style.marginTop = "12px";
    container.appendChild(renderAssignmentTable(results, true));
  });

  return section;
}

function renderDataManagementPanel(currentUser) {
  const section = createSection("資料管理", "可匯出備份、匯入回復，或重設回初始資料。");
  const grid = document.createElement("div");
  grid.className = "panel-grid";
  grid.innerHTML = `
    <div class="card panel master-card">
      <div class="section-heading">
        <div>
          <h3>備份匯出</h3>
          <p class="muted">將目前所有資料（員工、路線、班表、異動紀錄、設定）匯出為 JSON 檔案。</p>
        </div>
      </div>
      <button type="button" id="exportBackupButton">匯出 JSON 備份</button>
    </div>
    <div class="card panel master-card">
      <div class="section-heading">
        <div>
          <h3>匯入回復</h3>
          <p class="muted">選擇先前匯出的 JSON 備份檔案，覆蓋目前全部資料。此操作不可逆。</p>
        </div>
      </div>
      <label>選擇備份檔案<input type="file" id="importBackupFile" accept=".json"></label>
      <button type="button" class="secondary" id="importBackupButton" disabled>匯入並覆蓋</button>
    </div>
    <div class="card panel master-card">
      <div class="section-heading">
        <div>
          <h3>重設資料</h3>
          <p class="muted">清除所有修改，恢復系統初始的路線、員工與班表資料。</p>
        </div>
      </div>
      <button type="button" class="warn" id="resetDataButton">一鍵重設初始資料</button>
    </div>
  `;
  section.appendChild(grid);

  grid.querySelector("#exportBackupButton").addEventListener("click", () => {
    const blob = new Blob([JSON.stringify(state, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    const now = new Date();
    const ts = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, "0")}${String(now.getDate()).padStart(2, "0")}_${String(now.getHours()).padStart(2, "0")}${String(now.getMinutes()).padStart(2, "0")}`;
    a.href = url;
    a.download = `logistics-backup-${ts}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    logAction({
      actorId: currentUser.id,
      action: "backup-export",
      targetType: "settings",
      targetId: "full-backup",
      summary: "匯出完整備份",
      detail: `檔名 logistics-backup-${ts}.json`,
    });
    saveState();
  });

  const fileInput = grid.querySelector("#importBackupFile");
  const importButton = grid.querySelector("#importBackupButton");
  fileInput.addEventListener("change", () => {
    importButton.disabled = !fileInput.files.length;
  });
  importButton.addEventListener("click", () => {
    const file = fileInput.files[0];
    if (!file) return;
    const confirmed = window.confirm(`確定要匯入「${file.name}」並覆蓋現有資料嗎？\n此操作不可逆，建議先匯出備份。`);
    if (!confirmed) return;
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const imported = JSON.parse(e.target.result);
        if (!imported.employees || !imported.routes || !imported.assignments) {
          window.alert("檔案格式不正確，缺少必要的資料欄位。");
          return;
        }
        Object.assign(state, imported);
        saveState();
        window.alert("資料已成功匯入回復（含同步至雲端）。頁面將重新載入。");
        location.reload();
      } catch (err) {
        window.alert("JSON 解析失敗：" + err.message);
      }
    };
    reader.readAsText(file);
  });

  grid.querySelector("#resetDataButton").addEventListener("click", () => {
    const confirmed = window.confirm("確定要重設所有資料嗎？\n這會清除全部班表、員工修改與異動紀錄，恢復為系統初始狀態。\n\n此操作不可逆，建議先匯出備份。");
    if (!confirmed) return;
    const doubleConfirm = window.confirm("再次確認：真的要重設嗎？");
    if (!doubleConfirm) return;
    localStorage.removeItem(STORAGE_KEY);
    state = buildInitialState();
    // Clear Firebase too
    if (firebaseReady) {
      stateRef.set(state).then(() => {
        window.alert("已恢復初始資料（含雲端）。頁面將重新載入。");
        location.reload();
      }).catch(() => {
        saveState();
        window.alert("已恢復初始資料（雲端清除失敗，下次同步時覆蓋）。頁面將重新載入。");
        location.reload();
      });
    } else {
      saveState();
      window.alert("已恢復初始資料。頁面將重新載入。");
      location.reload();
    }
  });

  return section;
}

function renderAuditPanel() {
  const section = createSection("異動紀錄總覽", "主管可追溯固定班表生成、請假、代班與主資料調整。");
  const totalLogs = state.auditLogs.length;

  const details = document.createElement("details");
  details.className = "collapsible-list";
  const summary = document.createElement("summary");
  summary.textContent = `查看異動紀錄（共 ${totalLogs} 筆）`;
  details.appendChild(summary);

  const timeline = document.createElement("div");
  timeline.className = "timeline";
  state.auditLogs.slice(0, 30).forEach((log) => {
    const actor = getEmployeeById(log.actorId);
    const item = document.createElement("article");
    item.className = "timeline-item";
    item.innerHTML = `
      <p><strong>${log.summary}</strong></p>
      <p>${log.detail}</p>
      <p class="muted">${formatTimestamp(log.timestamp)} ｜ ${actor ? actor.name : "系統"} ｜ ${log.action}</p>
    `;
    timeline.appendChild(item);
  });
  if (totalLogs > 30) {
    const moreNote = document.createElement("p");
    moreNote.className = "muted";
    moreNote.style.textAlign = "center";
    moreNote.style.padding = "12px";
    moreNote.textContent = `僅顯示最近 30 筆，共 ${totalLogs} 筆紀錄`;
    timeline.appendChild(moreNote);
  }
  details.appendChild(timeline);
  section.appendChild(details);
  return section;
}
function render() {
  syncSelectors();
  // Show/hide keep-logged-in toggle based on role
  const isProtectedRole = protectedRoles.includes(state.session.role);
  keepLoggedInLabel.style.display = isProtectedRole ? "" : "none";
  if (isProtectedRole) {
    const cached = loadAuthCache();
    keepLoggedInToggle.checked = !!(cached && cached.keepLoggedIn);
  } else {
    keepLoggedInToggle.checked = false;
  }

  const currentUser = getEmployeeById(state.session.userId) || getRoleUsers(state.session.role)[0];
  if (!currentUser) {
    appEl.innerHTML = `<section class="card empty-state"><h2>找不到可用的登入人員</h2><p>請先到基本資料建立員工，或切換角色後再試一次。</p></section>`;
    return;
  }

  state.session.userId = currentUser.id;
  saveState();
  roleHint.textContent = buildRoleHint(currentUser.role);
  appEl.innerHTML = "";
  appEl.appendChild(renderPrototypeNotice());
  if (["teamLeader", "supervisor", "adminStaff"].includes(currentUser.role)) {
    appEl.appendChild(renderManagementLaunchers());
  }
  appEl.appendChild(renderStats(currentUser));
  appEl.appendChild(renderEmployeeHome(currentUser));

  if (["teamLeader", "supervisor"].includes(currentUser.role)) {
    appEl.appendChild(renderSchedulingWorkbenchV2(currentUser));
  }
  if (currentUser.role === "supervisor") {
    appEl.appendChild(renderMasterDataPanel(currentUser));
  }
  if (["teamLeader", "supervisor"].includes(currentUser.role)) {
    appEl.appendChild(renderHistoryQueryPanel(currentUser));
  }
  if (["adminStaff", "supervisor"].includes(currentUser.role)) {
    appEl.appendChild(renderDataManagementPanel(currentUser));
  }
  if (currentUser.role === "supervisor") {
    appEl.appendChild(renderAuditPanel());
  }
}

roleSelect.addEventListener("change", async (event) => {
  const rawValue = event.target.value;
  let newRole, targetUserId;

  // Parse combined "role:userId" format
  if (rawValue.includes(":")) {
    const parts = rawValue.split(":");
    newRole = parts[0];
    targetUserId = parts[1];
  } else {
    newRole = rawValue;
    targetUserId = getRoleUsers(newRole)[0]?.id || "";
  }

  // Verify PIN for the specific user
  if (protectedRoles.includes(newRole) && targetUserId) {
    if (!(await requireUserPin(targetUserId))) {
      syncSelectors();
      return;
    }
  } else if (!(await requirePin(newRole))) {
    syncSelectors();
    return;
  }

  state.session.role = newRole;
  state.session.userId = targetUserId;
  saveState();
  saveAuthCache();
  render();
});

userSelect.addEventListener("change", async (event) => {
  const newUserId = event.target.value;
  // If switching to a different user in a protected role, verify their PIN
  if (protectedRoles.includes(state.session.role) && !authenticatedUsers.has(newUserId)) {
    if (!(await requireUserPin(newUserId))) {
      event.target.value = state.session.userId; // Revert
      return;
    }
  }
  state.session.userId = newUserId;
  saveState();
  saveAuthCache();
  render();
});

keepLoggedInToggle.addEventListener("change", () => {
  if (keepLoggedInToggle.checked) {
    saveAuthCache();
  } else {
    clearAuthCache();
  }
});

if (protectedRoles.includes(state.session.role)) {
  const cached = loadAuthCache();
  if (cached && cached.keepLoggedIn && cached.role === state.session.role && cached.userId === state.session.userId) {
    (cached.authenticatedRoles || []).forEach(r => authenticatedRoles.add(r));
    (cached.authenticatedUsers || []).forEach(u => authenticatedUsers.add(u));
  } else {
    state.session.role = "operator";
    const users = getRoleUsers("operator");
    state.session.userId = users[0]?.id || "";
    saveState();
    clearAuthCache();
  }
}

// Initial render with localStorage data
render();

// ─── Firebase Initialization (with retry) ───
function applyFirebaseState(firebaseState) {
  const localSession = { ...state.session };
  state = firebaseState;
  state.session = localSession;
  // Run migrations (name rename + role fixes)
  applyEmployeeMigrations(state);
  state.employees.forEach((employee) => {
    if (!employee.employmentStatus) {
      employee.employmentStatus = employee.active ? "active" : "resigned";
    }
    employee.active = employee.employmentStatus === "active";
    if (typeof employee.canCoverShift !== "boolean") {
      employee.canCoverShift = !!employee.isRelief;
    }
  });
  state.session = state.session || {};
  state.session.lastGeneratedRange = state.session.lastGeneratedRange || null;
  state.labelSettings = state.labelSettings || {};
  state.labelSettings.leaveTypes = {
    annual: "特休", personal: "事假", sick: "病假",
    official: "公假", injury: "公傷", other: "其他",
    ...(state.labelSettings.leaveTypes || {}),
  };
  if (!state.pinSettings) {
    state.pinSettings = { teamLeader: "1234", adminStaff: "1234", supervisor: "0000", individual: {} };
  }
  if (!state.pinSettings.individual) {
    state.pinSettings.individual = {};
  }
  if (!state.mileageTable) {
    state.mileageTable = buildInitialState().mileageTable;
  }
  // Apply employee role migrations and save back to Firebase if changed
  const migrated = applyEmployeeMigrations(state);
  if (migrated) {
    saveState();
  } else {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  }
  render();
}

function startFirebaseListener() {
  stateRef.on("value", (snap) => {
    if (Date.now() - lastFirebaseSaveTime < FIREBASE_ECHO_DELAY) return;
    if (!snap.exists()) return;
    const remoteState = ensureArrays(snap.val());
    if (!remoteState.employees || !remoteState.routes) return;
    // Safety: don't accept remote state with zero assignments when local has data
    const remoteAssignments = (remoteState.assignments || []).length;
    const localAssignments = (state.assignments || []).length;
    if (remoteAssignments === 0 && localAssignments > 5) {
      console.warn("Firebase received empty assignments while local has data. Ignoring to prevent data loss.");
      return;
    }
    const localSession = { ...state.session };
    state = remoteState;
    state.session = localSession;
    // Apply role migrations after Firebase sync
    const migrated = applyEmployeeMigrations(state);
    if (migrated) {
      saveState(); // Write corrected roles back to Firebase
    } else {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
    }
    render();
  });
}

async function waitForFirebaseConnection(timeoutMs = 10000) {
  return new Promise((resolve, reject) => {
    const connRef = firebaseDb.ref(".info/connected");
    const timer = setTimeout(() => {
      connRef.off();
      reject(new Error("Firebase connection timeout after " + timeoutMs + "ms"));
    }, timeoutMs);
    connRef.on("value", (snap) => {
      if (snap.val() === true) {
        clearTimeout(timer);
        connRef.off();
        resolve();
      }
    });
  });
}

async function attemptFirebaseInit() {
  await waitForFirebaseConnection();
  const snapshot = await stateRef.once("value");
  if (snapshot.exists()) {
    const firebaseState = ensureArrays(snapshot.val());
    if (firebaseState.employees && firebaseState.routes && firebaseState.assignments) {
      applyFirebaseState(firebaseState);
    }
  } else {
    // First time: upload local state to Firebase (exclude session)
    const syncData = { ...state };
    delete syncData.session;
    await stateRef.set(syncData);
  }
  lastFirebaseSaveTime = Date.now();
  firebaseReady = true;
  firebaseInitialized = true;
  hideLoading();
  startFirebaseListener();
}

(async function initFirebase() {
  const MAX_RETRIES = 3;
  const RETRY_DELAY = 3000; // 3 seconds between retries

  for (let attempt = 1; attempt <= MAX_RETRIES; attempt++) {
    try {
      await attemptFirebaseInit();
      return; // Success, exit
    } catch (err) {
      console.warn(`Firebase init attempt ${attempt}/${MAX_RETRIES} failed:`, err);
      if (attempt < MAX_RETRIES) {
        await new Promise((resolve) => setTimeout(resolve, RETRY_DELAY));
      }
    }
  }
  // All retries failed
  console.warn("Firebase init failed after all retries, using localStorage only.");
  firebaseReady = false;
  hideLoading();
  showSyncStatus("error", "雲端連線失敗，使用本機資料（將自動重試）");
  // Keep trying in background every 10 seconds
  const bgRetry = setInterval(async () => {
    try {
      await attemptFirebaseInit();
      clearInterval(bgRetry);
      showSyncStatus("synced", "雲端連線恢復");
    } catch (e) { /* keep retrying silently */ }
  }, 10000);
})();












