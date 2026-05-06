// College of Agriculture, Balaghat — Organic Commerce Hub
// Google Apps Script backend
// Paste this file as Code.gs in a new Apps Script project.

const APP = {
  name: 'Organic Commerce Hub',
  college: 'College of Agriculture, Balaghat',
  spreadsheetProp: 'ORGANIC_HUB_SPREADSHEET_ID',
  adminPinProp: 'ORGANIC_HUB_ADMIN_PIN',
  adminEmailProp: 'ORGANIC_HUB_ADMIN_EMAIL',
  whatsappUrlProp: 'ORGANIC_HUB_WA_URL',
  whatsappTokenProp: 'ORGANIC_HUB_WA_TOKEN',
  smsUrlProp: 'ORGANIC_HUB_SMS_URL',
  smsTokenProp: 'ORGANIC_HUB_SMS_TOKEN',
  currency: 'INR',
  timezone: 'Asia/Kolkata'
};

const SHEETS = {
  products: 'Products',
  orders: 'Orders',
  sales: 'Sales',
  expenses: 'Expenses',
  customers: 'Customers',
  staff: 'Staff',
  settings: 'Settings',
  audit: 'Audit'
};

const DEFAULT_PRODUCTS = [
  { sku: 'earthworm', nameEn: 'Earthworm', nameHi: 'केंचुआ', category: 'Live Inputs', unit: 'kg', price: 800, stock: 100, active: true, description: 'Premium vermiculture earthworms for rapid composting.', image: '🪱' },
  { sku: 'vermicompost', nameEn: 'Vermicompost', nameHi: 'वर्मी कम्पोस्ट', category: 'Solid Fertilizer', unit: 'kg', price: 10, stock: 2000, active: true, description: 'Sieved organic compost made from earthworm activity.', image: '🌱' },
  { sku: 'enriched-vermicompost', nameEn: 'Enriched Vermicompost', nameHi: 'समृद्ध वर्मी कम्पोस्ट', category: 'Solid Fertilizer', unit: 'kg', price: 14, stock: 1000, active: true, description: 'Nutrient-enriched compost for stronger crop response.', image: '✨' },
  { sku: 'vermiwash', nameEn: 'Vermiwash', nameHi: 'वर्मी वॉश', category: 'Liquid Fertilizer', unit: 'L', price: 120, stock: 500, active: true, description: 'Liquid extract for foliar nutrition and plant vigor.', image: '💧' },
  { sku: 'bga', nameEn: 'BGA', nameHi: 'नील-हरित शैवाल', category: 'Biofertilizer', unit: 'pack', price: 250, stock: 200, active: true, description: 'Blue-green algae culture for paddy systems.', image: '🟦' },
  { sku: 'brahmastra', nameEn: 'Brahmastra', nameHi: 'ब्रह्मास्त्र', category: 'Biopesticide', unit: 'L', price: 220, stock: 250, active: true, description: 'Broad-spectrum botanical pesticide.', image: '🔥' },
  { sku: 'neemastra', nameEn: 'Neemastra', nameHi: 'नीमास्त्र', category: 'Biopesticide', unit: 'L', price: 150, stock: 250, active: true, description: 'Neem-based preventive spray.', image: '🌿' },
  { sku: 'agniastra', nameEn: 'Agniastra', nameHi: 'अग्नास्त्र', category: 'Biopesticide', unit: 'L', price: 250, stock: 150, active: true, description: 'Strong bio-pesticide for heavy infestations.', image: '⚡' },
  { sku: 'panchparni', nameEn: 'Panchparni', nameHi: 'पंचपर्णी', category: 'Plant Extract', unit: 'L', price: 180, stock: 150, active: true, description: '5-leaf decoction for crop protection.', image: '🍃' },
  { sku: 'astparni', nameEn: 'Astparni', nameHi: 'अष्टपर्णी', category: 'Plant Extract', unit: 'L', price: 230, stock: 120, active: true, description: '8-leaf extract for stronger defence.', image: '🌿' },
  { sku: 'dashparni', nameEn: 'Dashparni', nameHi: 'दशपर्णी', category: 'Plant Extract', unit: 'L', price: 280, stock: 120, active: true, description: '10-leaf master extract.', image: '🍀' },
  { sku: 'beejamrit', nameEn: 'Beejamrit', nameHi: 'बीजामृत', category: 'Seed Treatment', unit: 'L', price: 140, stock: 150, active: true, description: 'Seed treatment for healthier seedlings.', image: '🌾' },
  { sku: 'jeevamrit', nameEn: 'Jeevamrit', nameHi: 'जीवामृत', category: 'Soil Booster', unit: 'L', price: 120, stock: 250, active: true, description: 'Living microbial inoculant.', image: '🪴' },
  { sku: 'amrit-pani', nameEn: 'Amrit Pani', nameHi: 'अमृत पानी', category: 'Soil Booster', unit: 'L', price: 160, stock: 180, active: true, description: 'Concentrated soil revitalizer.', image: '🌊' },
  { sku: 'panchgavya', nameEn: 'Panchgavya', nameHi: 'पंचगव्य', category: 'Plant Tonic', unit: 'L', price: 200, stock: 200, active: true, description: 'Cow-based growth promoter and tonic.', image: '🐄' },
  { sku: 'vermicomposting-service', nameEn: 'Vermicomposting Service', nameHi: 'वर्मी कम्पोस्टिंग सेवा', category: 'Service', unit: 'kg', price: 10, stock: 999999, active: true, description: 'Custom vermicomposting processing service.', image: '🧾' }
];

function doGet() {
  ensureSetup_();
  const out = HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(APP.name)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return out;
}

function setup() {
  const db = ensureSetup_();
  return {
    ok: true,
    spreadsheetUrl: db.getUrl(),
    spreadsheetId: db.getId(),
    message: 'Setup complete'
  };
}

function ensureSetup_() {
  const props = PropertiesService.getScriptProperties();
  let ssId = props.getProperty(APP.spreadsheetProp);
  let ss;
  if (ssId) {
    try {
      ss = SpreadsheetApp.openById(ssId);
    } catch (err) {
      ss = SpreadsheetApp.create(APP.name);
      props.setProperty(APP.spreadsheetProp, ss.getId());
    }
  } else {
    ss = SpreadsheetApp.create(APP.name);
    props.setProperty(APP.spreadsheetProp, ss.getId());
  }

  const required = {};
  required[SHEETS.products] = ['SKU','Name (EN)','Name (HI)','Category','Unit','Price','Stock','Active','Description','Image','Updated At'];
  required[SHEETS.orders] = ['Timestamp','Type','Order No','Invoice No','Customer Name','Phone','Email','Address','City','State','Products JSON','Subtotal','Discount','Delivery','Tax','Grand Total','Payment Mode','Payment Status','Transaction Ref','Loyalty Points','Source','Staff Code','Notes'];
  required[SHEETS.sales] = ['Timestamp','Sale No','Student Name','Student Gmail','Student Phone','Buyer Name','Buyer Phone','Buyer Email','Product','SKU','Qty','Unit Price','Subtotal','Payment Mode','Payment Status','Transaction Ref','Notes','Saved By'];
  required[SHEETS.expenses] = ['Timestamp','Category','Amount','Paid To','Purpose','Payment Mode','Approved By','Notes'];
  required[SHEETS.customers] = ['Phone','Name','Email','Address','City','State','Lifetime Value','Orders','Loyalty Points','Last Order At'];
  required[SHEETS.staff] = ['Code','Name','Role','Allowed','Created At','Notes'];
  required[SHEETS.settings] = ['Key','Value'];
  required[SHEETS.audit] = ['Timestamp','Action','Actor','Details JSON'];

  Object.keys(required).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) sh.appendRow(required[name]);
  });

  seedDefaultSettings_(ss);
  seedProductsIfEmpty_(ss);
  seedDefaultStaff_(ss);
  return ss;
}

function seedDefaultSettings_(ss) {
  const sh = ss.getSheetByName(SHEETS.settings);
  const map = getSettingsMap_(sh);
  const defaults = {
    COLLEGE_NAME: APP.college,
    ADMIN_EMAIL: 'coabalaghatorganic@gmail.com',
    ADMIN_PHONE: '7877612427',
    CURRENCY: 'INR',
    WHATSAPP_API_URL: '',
    WHATSAPP_TOKEN: '',
    SMS_API_URL: '',
    SMS_TOKEN: '',
    UPI_ID: '',
    SHOP_OPEN: 'YES',
    LANG_DEFAULT: 'en',
    FREE_SHIPPING_THRESHOLD: '999',
    STOCK_ALERT_QTY: '25',
    ADMIN_PIN: '2022',
    STUDENT_SALE_CODES: 'STU-2022-A,STU-2022-B,STU-2022-C'
  };
  const rows = [];
  Object.keys(defaults).forEach(k => {
    if (!(k in map)) rows.push([k, defaults[k]]);
  });
  if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, 2).setValues(rows);
}

function seedProductsIfEmpty_(ss) {
  const sh = ss.getSheetByName(SHEETS.products);
  if (sh.getLastRow() > 1) return;
  const now = new Date();
  const rows = DEFAULT_PRODUCTS.map(p => [
    p.sku, p.nameEn, p.nameHi, p.category, p.unit, p.price, p.stock, p.active ? 'YES' : 'NO', p.description, p.image, now
  ]);
  sh.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

function seedDefaultStaff_(ss) {
  const sh = ss.getSheetByName(SHEETS.staff);
  if (sh.getLastRow() > 1) return;
  const rows = [
    ['STU-2022-A','Student A','Sales Entry','YES',new Date(),'Default sample code'],
    ['STU-2022-B','Student B','Sales Entry','YES',new Date(),'Default sample code'],
    ['STU-2022-C','Student C','Sales Entry','YES',new Date(),'Default sample code']
  ];
  sh.getRange(2,1,rows.length,rows[0].length).setValues(rows);
}

function apiBootstrap() {
  const ss = ensureSetup_();
  const settings = getSettings_();
  return {
    ok: true,
    app: {
      name: APP.name,
      college: settings.COLLEGE_NAME || APP.college,
      currency: settings.CURRENCY || 'INR',
      freeShippingThreshold: Number(settings.FREE_SHIPPING_THRESHOLD || 999),
      shopOpen: (settings.SHOP_OPEN || 'YES') === 'YES',
      langDefault: settings.LANG_DEFAULT || 'en'
    },
    products: getProducts_(),
    dashboard: getDashboard_(false),
    publicStats: getPublicStats_(),
    categories: getCategories_(),
    settings: {
      adminEmail: settings.ADMIN_EMAIL || '',
      adminPhone: settings.ADMIN_PHONE || '',
      upiId: settings.UPI_ID || ''
    },
    links: {
      sheetUrl: ss.getUrl()
    }
  };
}

function apiLoginAdmin(pin) {
  const settings = getSettings_();
  if (String(pin || '') !== String(settings.ADMIN_PIN || '')) {
    throw new Error('Invalid admin PIN');
  }
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('ADMIN_' + token, '1', 21600);
  audit_('ADMIN_LOGIN', 'system', { ok: true });
  return { ok: true, token: token };
}

function apiLoginStaff(code) {
  const staff = getStaff_().find(s => s.code === String(code || '').trim() && s.allowed === true);
  if (!staff) throw new Error('Invalid staff code');
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('STAFF_' + token, staff.code, 21600);
  return { ok: true, token: token, staff: staff };
}

function apiLogout(token) {
  if (token) {
    CacheService.getScriptCache().remove('ADMIN_' + token);
    CacheService.getScriptCache().remove('STAFF_' + token);
  }
  return { ok: true };
}

function apiGetAdminData(token) {
  requireAdmin_(token);
  return {
    ok: true,
    dashboard: getDashboard_(true),
    products: getProducts_(true),
    orders: getOrders_(50),
    sales: getSales_(50),
    expenses: getExpenses_(50),
    staff: getStaff_(),
    settings: getSettings_(),
    audit: getAudit_(50),
    customers: getCustomers_(50)
  };
}

function apiSaveOrder(payload) {
  ensureSetup_();
  const settings = getSettings_();
  const data = normalizeOrder_(payload, 'Public');
  appendOrder_(data);
  upsertCustomer_(data.customer, data.grandTotal);
  audit_('PUBLIC_ORDER', data.customer.name || 'guest', data);

  sendNotifications_({
    kind: 'order',
    toEmail: settings.ADMIN_EMAIL,
    buyerEmail: data.customer.email,
    phone: data.customer.phone,
    html: buildOrderEmailHtml_(data, settings),
    text: buildOrderEmailText_(data, settings)
  });

  return { ok: true, order: data, dashboard: getDashboard_(true) };
}

function apiSaveStudentSale(payload) {
  ensureSetup_();
  const settings = getSettings_();
  const staffToken = String(payload.staffToken || '');
  const staffCode = CacheService.getScriptCache().get('STAFF_' + staffToken);
  if (!staffCode) throw new Error('Student access code expired or invalid');

  const staff = getStaff_().find(s => s.code === staffCode);
  if (!staff || !staff.allowed) throw new Error('Student is not authorized');

  const data = normalizeSale_(payload, staff.code, staff.name);
  appendSale_(data);
  upsertCustomer_({
    name: data.buyerName || 'Walk-in buyer',
    phone: data.buyerPhone || '',
    email: data.buyerEmail || '',
    address: '',
    city: '',
    state: ''
  }, data.subtotal);

  audit_('STUDENT_SALE', staff.name, data);
  sendNotifications_({
    kind: 'sale',
    toEmail: settings.ADMIN_EMAIL,
    buyerEmail: data.buyerEmail,
    phone: data.buyerPhone,
    html: buildSaleEmailHtml_(data, settings, staff),
    text: buildSaleEmailText_(data, settings, staff)
  });
  return { ok: true, sale: data, dashboard: getDashboard_(true) };
}

function apiSaveExpense(payload, token) {
  requireAdmin_(token);
  const ss = ensureSetup_();
  const sh = ss.getSheetByName(SHEETS.expenses);
  const row = [
    new Date(),
    String(payload.category || 'General'),
    Number(payload.amount || 0),
    String(payload.paidTo || ''),
    String(payload.purpose || ''),
    String(payload.paymentMode || 'Cash'),
    String(payload.approvedBy || 'Admin'),
    String(payload.notes || '')
  ];
  sh.appendRow(row);
  audit_('EXPENSE', 'admin', payload);
  return { ok: true, dashboard: getDashboard_(true) };
}

function apiSaveProduct(payload, token) {
  requireAdmin_(token);
  const ss = ensureSetup_();
  const sh = ss.getSheetByName(SHEETS.products);
  const sku = String(payload.sku || '').trim();
  if (!sku) throw new Error('SKU is required');
  const rows = sh.getDataRange().getValues();
  const idx = rows.findIndex((r, i) => i > 0 && String(r[0]) === sku);
  const now = new Date();
  const row = [
    sku,
    String(payload.nameEn || ''),
    String(payload.nameHi || ''),
    String(payload.category || ''),
    String(payload.unit || ''),
    Number(payload.price || 0),
    Number(payload.stock || 0),
    payload.active ? 'YES' : 'NO',
    String(payload.description || ''),
    String(payload.image || ''),
    now
  ];
  if (idx > 0) {
    sh.getRange(idx + 1, 1, 1, row.length).setValues([row]);
  } else {
    sh.appendRow(row);
  }
  audit_('PRODUCT_SAVE', 'admin', payload);
  return { ok: true, products: getProducts_(true) };
}

function apiSaveSettings(payload, token) {
  requireAdmin_(token);
  const ss = ensureSetup_();
  const sh = ss.getSheetByName(SHEETS.settings);
  const map = getSettingsMap_(sh);
  Object.keys(payload || {}).forEach(k => {
    const v = payload[k];
    if (v === undefined || v === null) return;
    setSetting_(sh, k, String(v));
  });
  audit_('SETTINGS_UPDATE', 'admin', payload);
  return { ok: true, settings: getSettings_() };
}

function apiSaveStaff(payload, token) {
  requireAdmin_(token);
  const ss = ensureSetup_();
  const sh = ss.getSheetByName(SHEETS.staff);
  const code = String(payload.code || '').trim();
  if (!code) throw new Error('Code required');
  const rows = sh.getDataRange().getValues();
  const idx = rows.findIndex((r, i) => i > 0 && String(r[0]) === code);
  const row = [
    code,
    String(payload.name || ''),
    String(payload.role || 'Sales Entry'),
    payload.allowed ? 'YES' : 'NO',
    new Date(),
    String(payload.notes || '')
  ];
  if (idx > 0) {
    sh.getRange(idx + 1, 1, 1, row.length).setValues([row]);
  } else {
    sh.appendRow(row);
  }
  audit_('STAFF_SAVE', 'admin', payload);
  return { ok: true, staff: getStaff_() };
}

function apiSendTestNotification(token) {
  requireAdmin_(token);
  const settings = getSettings_();
  sendNotifications_({
    kind: 'test',
    toEmail: settings.ADMIN_EMAIL,
    html: '<p>Test notification from Organic Commerce Hub.</p>',
    text: 'Test notification from Organic Commerce Hub.'
  });
  return { ok: true };
}

function apiMarkPayment(payload, token) {
  requireAdmin_(token);
  const kind = String(payload.kind || 'order');
  const ref = String(payload.ref || '');
  const status = String(payload.status || 'PAID');
  const ss = ensureSetup_();
  const sheet = ss.getSheetByName(kind === 'sale' ? SHEETS.sales : SHEETS.orders);
  const values = sheet.getDataRange().getValues();
  const rowIdx = values.findIndex((r, i) => i > 0 && String(r[2] || r[1]) === ref || String(r[3]) === ref);
  if (rowIdx > 0) {
    if (kind === 'sale') {
      sheet.getRange(rowIdx + 1, 15).setValue(status);
    } else {
      sheet.getRange(rowIdx + 1, 18).setValue(status);
    }
  }
  return { ok: true };
}

function getPublicStats_() {
  const dash = getDashboard_(false);
  return {
    productsCount: getProducts_().filter(p => p.active).length,
    ordersCount: dash.ordersCount,
    salesCount: dash.salesCount,
    revenue: dash.revenue,
    cashInHand: dash.cashInHand
  };
}

function getCategories_() {
  const cats = {};
  getProducts_().forEach(p => cats[p.category] = true);
  return Object.keys(cats).sort();
}

function getDashboard_(includeRows) {
  const orders = getOrders_();
  const sales = getSales_();
  const expenses = getExpenses_();
  const revenue = orders.concat(sales).reduce((sum, x) => sum + Number(x.grandTotal || x.subtotal || 0), 0);
  const expenseTotal = expenses.reduce((sum, x) => sum + Number(x.amount || 0), 0);
  const cashInHand = revenue - expenseTotal;
  const topProduct = topProduct_();
  const monthly = monthlySeries_();
  return {
    revenue: round_(revenue),
    expenseTotal: round_(expenseTotal),
    cashInHand: round_(cashInHand),
    ordersCount: orders.length,
    salesCount: sales.length,
    customersCount: getCustomers_().length,
    productsCount: getProducts_().length,
    topProduct: topProduct,
    monthly: monthly,
    recentOrders: includeRows ? orders.slice(0, 10) : [],
    recentSales: includeRows ? sales.slice(0, 10) : [],
    recentExpenses: includeRows ? expenses.slice(0, 10) : []
  };
}

function monthlySeries_() {
  const orders = getOrders_();
  const sales = getSales_();
  const all = orders.concat(sales).map(x => ({
    month: Utilities.formatDate(new Date(x.timestamp), APP.timezone, 'yyyy-MM'),
    amount: Number(x.grandTotal || x.subtotal || 0)
  }));
  const map = {};
  all.forEach(x => map[x.month] = (map[x.month] || 0) + x.amount);
  const months = Object.keys(map).sort().slice(-6);
  return months.map(m => ({ month: m, amount: round_(map[m] || 0) }));
}

function topProduct_() {
  const orders = getOrders_();
  const sales = getSales_();
  const tally = {};
  orders.concat(sales).forEach(r => {
    const items = r.items || [];
    items.forEach(it => {
      const k = String(it.name || it.product || it.sku || '');
      if (!k) return;
      tally[k] = (tally[k] || 0) + Number(it.qty || 0);
    });
  });
  const top = Object.keys(tally).sort((a,b) => tally[b] - tally[a])[0];
  return top ? { name: top, qty: tally[top] } : { name: '—', qty: 0 };
}

function getProducts_(all) {
  const ss = ensureSetup_();
  const sh = ss.getSheetByName(SHEETS.products);
  const rows = sh.getDataRange().getValues();
  const out = rows.slice(1).map(r => ({
    sku: String(r[0] || ''),
    nameEn: String(r[1] || ''),
    nameHi: String(r[2] || ''),
    category: String(r[3] || ''),
    unit: String(r[4] || ''),
    price: Number(r[5] || 0),
    stock: Number(r[6] || 0),
    active: String(r[7] || '') === 'YES',
    description: String(r[8] || ''),
    image: String(r[9] || ''),
    updatedAt: r[10] || ''
  }));
  return all ? out : out.filter(p => p.active);
}

function getOrders_(limit) {
  const ss = ensureSetup_();
  const sh = ss.getSheetByName(SHEETS.orders);
  const rows = sh.getDataRange().getValues().slice(1).reverse();
  return rows.slice(0, limit || rows.length).map(r => ({
    timestamp: r[0],
    type: r[1],
    orderNo: r[2],
    invoiceNo: r[3],
    customer: { name: r[4], phone: r[5], email: r[6], address: r[7], city: r[8], state: r[9] },
    items: safeParse_(r[10], []),
    subtotal: Number(r[11] || 0),
    discount: Number(r[12] || 0),
    delivery: Number(r[13] || 0),
    tax: Number(r[14] || 0),
    grandTotal: Number(r[15] || 0),
    paymentMode: r[16],
    paymentStatus: r[17],
    transactionRef: r[18],
    loyaltyPoints: Number(r[19] || 0),
    source: r[20],
    staffCode: r[21],
    notes: r[22]
  }));
}

function getSales_(limit) {
  const ss = ensureSetup_();
  const sh = ss.getSheetByName(SHEETS.sales);
  const rows = sh.getDataRange().getValues().slice(1).reverse();
  return rows.slice(0, limit || rows.length).map(r => ({
    timestamp: r[0],
    saleNo: r[1],
    studentName: r[2],
    studentGmail: r[3],
    studentPhone: r[4],
    buyerName: r[5],
    buyerPhone: r[6],
    buyerEmail: r[7],
    product: r[8],
    sku: r[9],
    qty: Number(r[10] || 0),
    unitPrice: Number(r[11] || 0),
    subtotal: Number(r[12] || 0),
    paymentMode: r[13],
    paymentStatus: r[14],
    transactionRef: r[15],
    notes: r[16],
    savedBy: r[17],
    items: [{ sku: r[9], name: r[8], qty: Number(r[10] || 0), price: Number(r[11] || 0) }],
    grandTotal: Number(r[12] || 0)
  }));
}

function getExpenses_(limit) {
  const ss = ensureSetup_();
  const sh = ss.getSheetByName(SHEETS.expenses);
  const rows = sh.getDataRange().getValues().slice(1).reverse();
  return rows.slice(0, limit || rows.length).map(r => ({
    timestamp: r[0],
    category: r[1],
    amount: Number(r[2] || 0),
    paidTo: r[3],
    purpose: r[4],
    paymentMode: r[5],
    approvedBy: r[6],
    notes: r[7]
  }));
}

function getCustomers_(limit) {
  const ss = ensureSetup_();
  const sh = ss.getSheetByName(SHEETS.customers);
  const rows = sh.getDataRange().getValues().slice(1).reverse();
  return rows.slice(0, limit || rows.length).map(r => ({
    phone: r[0],
    name: r[1],
    email: r[2],
    address: r[3],
    city: r[4],
    state: r[5],
    lifetimeValue: Number(r[6] || 0),
    orders: Number(r[7] || 0),
    loyaltyPoints: Number(r[8] || 0),
    lastOrderAt: r[9]
  }));
}

function getStaff_() {
  const ss = ensureSetup_();
  const sh = ss.getSheetByName(SHEETS.staff);
  return sh.getDataRange().getValues().slice(1).map(r => ({
    code: String(r[0] || ''),
    name: String(r[1] || ''),
    role: String(r[2] || ''),
    allowed: String(r[3] || '') === 'YES',
    createdAt: r[4],
    notes: String(r[5] || '')
  }));
}

function getSettings_() {
  const ss = ensureSetup_();
  const sh = ss.getSheetByName(SHEETS.settings);
  const rows = sh.getDataRange().getValues().slice(1);
  const out = {};
  rows.forEach(r => out[String(r[0])] = String(r[1] || ''));
  return out;
}

function getSettingsMap_(sh) {
  const rows = sh.getDataRange().getValues().slice(1);
  const out = {};
  rows.forEach(r => out[String(r[0])] = String(r[1] || ''));
  return out;
}

function setSetting_(sh, key, value) {
  const rows = sh.getDataRange().getValues();
  const idx = rows.findIndex((r, i) => i > 0 && String(r[0]) === key);
  if (idx > 0) {
    sh.getRange(idx + 1, 2).setValue(value);
  } else {
    sh.appendRow([key, value]);
  }
}

function getAudit_(limit) {
  const ss = ensureSetup_();
  const sh = ss.getSheetByName(SHEETS.audit);
  return sh.getDataRange().getValues().slice(1).reverse().slice(0, limit || 50).map(r => ({
    timestamp: r[0], action: r[1], actor: r[2], details: safeParse_(r[3], {})
  }));
}

function requireAdmin_(token) {
  if (!token) throw new Error('Admin login required');
  const ok = CacheService.getScriptCache().get('ADMIN_' + token);
  if (!ok) throw new Error('Admin session expired');
}

function appendOrder_(data) {
  const ss = ensureSetup_();
  const sh = ss.getSheetByName(SHEETS.orders);
  sh.appendRow([
    data.timestamp,
    data.type,
    data.orderNo,
    data.invoiceNo,
    data.customer.name,
    data.customer.phone,
    data.customer.email,
    data.customer.address,
    data.customer.city,
    data.customer.state,
    JSON.stringify(data.items),
    data.subtotal,
    data.discount,
    data.delivery,
    data.tax,
    data.grandTotal,
    data.paymentMode,
    data.paymentStatus,
    data.transactionRef,
    data.loyaltyPoints,
    data.source,
    data.staffCode,
    data.notes
  ]);
}

function appendSale_(data) {
  const ss = ensureSetup_();
  const sh = ss.getSheetByName(SHEETS.sales);
  sh.appendRow([
    data.timestamp,
    data.saleNo,
    data.studentName,
    data.studentGmail,
    data.studentPhone,
    data.buyerName,
    data.buyerPhone,
    data.buyerEmail,
    data.product,
    data.sku,
    data.qty,
    data.unitPrice,
    data.subtotal,
    data.paymentMode,
    data.paymentStatus,
    data.transactionRef,
    data.notes,
    data.savedBy
  ]);
}

function upsertCustomer_(customer, amount) {
  const phone = String(customer.phone || '').trim();
  if (!phone) return;
  const ss = ensureSetup_();
  const sh = ss.getSheetByName(SHEETS.customers);
  const rows = sh.getDataRange().getValues();
  const idx = rows.findIndex((r, i) => i > 0 && String(r[0]) === phone);
  const newValue = Number(amount || 0);
  if (idx > 0) {
    const prevVal = Number(rows[idx][6] || 0);
    const prevOrders = Number(rows[idx][7] || 0);
    const prevPts = Number(rows[idx][8] || 0);
    sh.getRange(idx + 1, 2).setValue(customer.name || rows[idx][1]);
    if (customer.email) sh.getRange(idx + 1, 3).setValue(customer.email);
    if (customer.address) sh.getRange(idx + 1, 4).setValue(customer.address);
    if (customer.city) sh.getRange(idx + 1, 5).setValue(customer.city);
    if (customer.state) sh.getRange(idx + 1, 6).setValue(customer.state);
    sh.getRange(idx + 1, 7).setValue(round_(prevVal + newValue));
    sh.getRange(idx + 1, 8).setValue(prevOrders + 1);
    sh.getRange(idx + 1, 9).setValue(prevPts + Math.floor(newValue / 100));
    sh.getRange(idx + 1, 10).setValue(new Date());
  } else {
    sh.appendRow([
      phone,
      customer.name || '',
      customer.email || '',
      customer.address || '',
      customer.city || '',
      customer.state || '',
      round_(newValue),
      1,
      Math.floor(newValue / 100),
      new Date()
    ]);
  }
}

function normalizeOrder_(payload, type) {
  const items = Array.isArray(payload.items) ? payload.items.map(normalizeItem_) : [];
  if (!items.length) throw new Error('Cart is empty');
  const subtotal = round_(items.reduce((s, i) => s + (Number(i.price || 0) * Number(i.qty || 0)), 0));
  const discount = round_(Number(payload.discount || 0));
  const delivery = round_(Number(payload.delivery || 0));
  const tax = round_(Number(payload.tax || 0));
  const grandTotal = round_(subtotal - discount + delivery + tax);
  return {
    timestamp: new Date(),
    type: type,
    orderNo: 'ORD-' + Utilities.formatDate(new Date(), APP.timezone, 'yyyyMMdd-HHmmss') + '-' + Math.floor(Math.random() * 900 + 100),
    invoiceNo: 'INV-' + Utilities.formatDate(new Date(), APP.timezone, 'yyMMdd') + '-' + Math.floor(Math.random() * 9000 + 1000),
    customer: {
      name: String(payload.customer?.name || '').trim(),
      phone: String(payload.customer?.phone || '').trim(),
      email: String(payload.customer?.email || '').trim(),
      address: String(payload.customer?.address || '').trim(),
      city: String(payload.customer?.city || '').trim(),
      state: String(payload.customer?.state || '').trim()
    },
    items: items,
    subtotal: subtotal,
    discount: discount,
    delivery: delivery,
    tax: tax,
    grandTotal: grandTotal,
    paymentMode: String(payload.paymentMode || 'UPI'),
    paymentStatus: String(payload.paymentStatus || 'PENDING'),
    transactionRef: String(payload.transactionRef || '').trim(),
    loyaltyPoints: Math.floor(grandTotal / 100),
    source: String(payload.source || 'website'),
    staffCode: String(payload.staffCode || ''),
    notes: String(payload.notes || '')
  };
}

function normalizeSale_(payload, staffCode, staffName) {
  const qty = Number(payload.qty || 0);
  const unitPrice = Number(payload.unitPrice || 0);
  const subtotal = round_(qty * unitPrice);
  return {
    timestamp: new Date(),
    saleNo: 'SAL-' + Utilities.formatDate(new Date(), APP.timezone, 'yyyyMMdd-HHmmss') + '-' + Math.floor(Math.random() * 900 + 100),
    studentName: String(payload.studentName || ''),
    studentGmail: String(payload.studentGmail || ''),
    studentPhone: String(payload.studentPhone || ''),
    buyerName: String(payload.buyerName || ''),
    buyerPhone: String(payload.buyerPhone || ''),
    buyerEmail: String(payload.buyerEmail || ''),
    product: String(payload.product || ''),
    sku: String(payload.sku || ''),
    qty: qty,
    unitPrice: unitPrice,
    subtotal: subtotal,
    paymentMode: String(payload.paymentMode || 'Cash'),
    paymentStatus: String(payload.paymentStatus || 'PENDING'),
    transactionRef: String(payload.transactionRef || ''),
    notes: String(payload.notes || ''),
    savedBy: staffName + ' (' + staffCode + ')'
  };
}

function normalizeItem_(item) {
  return {
    sku: String(item.sku || ''),
    name: String(item.name || ''),
    qty: Number(item.qty || 0),
    price: Number(item.price || 0),
    unit: String(item.unit || '')
  };
}

function sendNotifications_(obj) {
  const settings = getSettings_();
  const subject = obj.kind === 'sale' ? 'New Student Sale Recorded' : obj.kind === 'test' ? 'Test notification' : 'New Customer Order Received';
  try {
    if (obj.toEmail) {
      GmailApp.sendEmail(obj.toEmail, subject, obj.text || '', { htmlBody: obj.html || obj.text || '' });
    }
    if (obj.buyerEmail) {
      GmailApp.sendEmail(obj.buyerEmail, 'Thank you for your order from ' + APP.college, 'Your order has been received.');
    }
  } catch (err) {
    audit_('EMAIL_FAIL', 'system', { error: String(err) });
  }
  // Optional webhook hooks for WhatsApp/SMS providers. You can connect Twilio / WhatsApp Cloud API here.
  try {
    if (settings.WHATSAPP_API_URL && settings.WHATSAPP_TOKEN) {
      UrlFetchApp.fetch(settings.WHATSAPP_API_URL, {
        method: 'post',
        contentType: 'application/json',
        headers: { Authorization: 'Bearer ' + settings.WHATSAPP_TOKEN },
        payload: JSON.stringify({
          to: obj.phone || '',
          message: obj.text || 'Thank you',
          kind: obj.kind
        }),
        muteHttpExceptions: true
      });
    }
  } catch (err) {
    audit_('WHATSAPP_FAIL', 'system', { error: String(err) });
  }
  try {
    if (settings.SMS_API_URL && settings.SMS_TOKEN) {
      UrlFetchApp.fetch(settings.SMS_API_URL, {
        method: 'post',
        contentType: 'application/json',
        headers: { Authorization: 'Bearer ' + settings.SMS_TOKEN },
        payload: JSON.stringify({
          to: obj.phone || '',
          message: obj.text || 'Thank you',
          kind: obj.kind
        }),
        muteHttpExceptions: true
      });
    }
  } catch (err) {
    audit_('SMS_FAIL', 'system', { error: String(err) });
  }
}

function buildOrderEmailHtml_(o, settings) {
  const rows = o.items.map(it => `<tr><td>${escapeHtml_(it.name)}</td><td>${it.qty}</td><td>₹${it.price}</td><td>₹${Number(it.qty) * Number(it.price)}</td></tr>`).join('');
  return `
  <div style="font-family:Arial,sans-serif;line-height:1.55">
    <h2 style="margin:0 0 6px">${escapeHtml_(APP.college)} — New Order</h2>
    <p style="margin:0 0 10px">Invoice: <b>${escapeHtml_(o.invoiceNo)}</b></p>
    <p><b>Customer:</b> ${escapeHtml_(o.customer.name)} | ${escapeHtml_(o.customer.phone)}<br>
       <b>Payment:</b> ${escapeHtml_(o.paymentMode)} / ${escapeHtml_(o.paymentStatus)}</p>
    <table border="1" cellpadding="8" cellspacing="0" style="border-collapse:collapse;width:100%">
      <tr><th>Product</th><th>Qty</th><th>Price</th><th>Total</th></tr>
      ${rows}
    </table>
    <p><b>Grand Total:</b> ₹${o.grandTotal}</p>
    <p>Thank you for supporting student-run organic inputs.</p>
  </div>`;
}

function buildOrderEmailText_(o) {
  return `${APP.college} order ${o.invoiceNo}\nCustomer: ${o.customer.name} / ${o.customer.phone}\nTotal: ₹${o.grandTotal}\nThank you for supporting student-run organic inputs.`;
}

function buildSaleEmailHtml_(o, settings, staff) {
  return `
  <div style="font-family:Arial,sans-serif;line-height:1.55">
    <h2 style="margin:0 0 6px">${escapeHtml_(APP.college)} — Student Sale</h2>
    <p><b>Student:</b> ${escapeHtml_(staff.name)} (${escapeHtml_(staff.code)})<br>
       <b>Buyer:</b> ${escapeHtml_(o.buyerName)} | ${escapeHtml_(o.buyerPhone)}<br>
       <b>Product:</b> ${escapeHtml_(o.product)} × ${o.qty}<br>
       <b>Total:</b> ₹${o.subtotal}</p>
  </div>`;
}

function buildSaleEmailText_(o, settings, staff) {
  return `${APP.college} student sale\nStudent: ${staff.name} (${staff.code})\nBuyer: ${o.buyerName} / ${o.buyerPhone}\nProduct: ${o.product} x ${o.qty}\nTotal: ₹${o.subtotal}`;
}

function round_(n) { return Math.round((Number(n) || 0) * 100) / 100; }

function safeParse_(v, fallback) {
  try {
    if (typeof v === 'object') return v;
    return JSON.parse(v || 'null') ?? fallback;
  } catch (e) {
    return fallback;
  }
}

function escapeHtml_(s) {
  return String(s || '').replace(/[&<>"']/g, m => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]));
}

function audit_(action, actor, details) {
  const ss = ensureSetup_();
  const sh = ss.getSheetByName(SHEETS.audit);
  sh.appendRow([new Date(), action, actor, JSON.stringify(details || {})]);
}