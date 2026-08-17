/*******************************************************
 * CUSTODIO SUPLICO FIREWORKS
 * PHASE 1 — STABLE FOUNDATION
 *
 * Architecture:
 * Google Sheet -> Apps Script -> Index.html
 * Frontend communication uses direct google.script.run.
 *******************************************************/

const APP = {
  NAME: 'CUSTODIO SUPLICO FIREWORKS',
  VERSION: 'PHASE-1',
  INITIAL_ADMIN: {
    email: 'admin@custodiosuplico.com',
    password: 'Admin@12345',
    name: 'System Administrator'
  },
  SHEETS: {
    Users: [
      'ID','UserID','Full Name','Email','Password Hash','Salt',
      'Role','Status','Created Date','Updated Date','Last Login'
    ],
    Settings: [
      'ID','Key','Value','Created Date','Updated Date'
    ],
    ActivityLogs: [
      'ID','User','Action','Module','Details','Timestamp'
    ]
  }
};

/**
 * Entry point for the deployed Web App.
 * NEVER open Index.html directly.
 */
function doGet() {
  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setTitle(APP.NAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Run this once from the Apps Script editor.
 * It creates the Phase 1 database tables and admin account.
 */
function setupSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Please bind this Apps Script project to a Google Sheet.');

  Object.keys(APP.SHEETS).forEach(name => ensureSheet_(ss, name, APP.SHEETS[name]));
  ensureSheetIfMissing_(ss, 'Materials', ['ID','MaterialID','Name','Category','Unit','Purchase Cost','Minimum Stock','Status','Description','Created Date','Updated Date']);
  ensureSheetIfMissing_(ss, 'BookingUsage', ['ID','UsageID','BookingID','Usage Type','ItemID','Item Name','Quantity','Purchase Cost','Total Cost','Used Date','Status','Created Date','Updated Date']);
  ensureSheetIfMissing_(ss, 'BookingItems', ['ID','BookingItemID','BookingID','ProductID','Product Name','Quantity','Unit Price','Subtotal','Created Date','Updated Date']);
  ensureSheetIfMissing_(ss, 'Materials', ['ID','MaterialID','Name','Category','Unit','Purchase Cost','Minimum Stock','Status','Description','Created Date','Updated Date']);
  ensureSheetIfMissing_(ss, 'BookingUsage', ['ID','UsageID','BookingID','Usage Type','ItemID','Item Name','Quantity','Purchase Cost','Total Cost','Used Date','Status','Created Date','Updated Date']);
  ensureSheetIfMissing_(ss, 'BookingCrew', ['ID','CrewID','BookingID','EmployeeID','Employee Name','Role','Show Assignment','Pay','Notes','Created Date','Updated Date']);
  ensureSheetIfMissing_(ss, 'Employees', ['ID','EmployeeID','Employee Name','Role','Employee Type','Contact Number','Salary','Status','Notes','Created Date','Updated Date']);
  ensureSheetIfMissing_(ss, 'EmployeeAdvances', ['ID','AdvanceID','EmployeeID','Employee Name','Type','Amount','Amount Paid','Outstanding Balance','Date','Purpose','Notes','Status','Created Date','Updated Date']);
  ensureSheetIfMissing_(ss, 'BookingExpenses', ['ID','ExpenseID','BookingID','Date','Category','Description','Amount','Created Date','Updated Date']);
  ensureSheetIfMissing_(ss, 'PaymentSchedules', ['ID','ScheduleID','BookingID','Milestone','Due Date','Amount','Paid','Balance','Status','Created Date','Updated Date']);
  ensureSheetIfMissing_(ss, 'AuditLog', ['ID','User','Action','Module','Record ID','Timestamp','Details']);
  ensureSheetIfMissing_(ss, 'Clients', ['ID','ClientID','Name','Phone','Email','Address','Notes','Created Date','Updated Date']);
  ensureSheetIfMissing_(ss, 'Loans', ['ID','LoanID','Lender','Loan Amount','Amount Paid','Outstanding Balance','Date Borrowed','Due Date','Purpose','Notes','Created Date','Updated Date']);
  ensureSheetIfMissing_(ss, 'Bookings', ['ID','BookingID','Client Name','Contact Information','Event Type','Event Name','Event Date','Event Time','Location','Subtotal','Discount Type','Discount Value','Discount Amount','Additional Charges','Final Amount','DownPayment','Paid Amount','Balance','Payment Status','Booking Status','Notes','Created Date','Updated Date']);
  ensureSheetIfMissing_(ss, 'Products', ['ID','ProductID','Name','Source','Category','Unit','Purchase Cost','Selling Price','Minimum Stock','Status','Description','Created Date','Updated Date']);
  ensureSheetIfMissing_(ss, 'Inventory', ['ID','InventoryID','ProductID','Product Name','Current Stock','Reserved Stock','Available Stock','Minimum Stock','Unit Cost','Landed Cost','Selling Price','Inventory Value','Updated Date']);
  ensureSheetIfMissing_(ss, 'Payments', ['ID','PaymentID','BookingID','Payment Date','Milestone','Amount','Method','Reference','Notes','Status','Created Date','Updated Date']);
  ensureSheetIfMissing_(ss, 'Expenses', ['ID','ExpenseID','Date','Amount','Category','SupplierID','BookingID','Description','Approval Status','Created Date','Updated Date']);
  ensureSheetIfMissing_(ss, 'Payroll', ['ID','PayrollID','EmployeeID','Employee Name','Employee Type','Compensation Type','Payroll Period','Contract Amount','Contract Rate','Basic Salary','Contract Pay','Overtime','Allowances','Gas Allowance','Food Allowance','Accommodation Allowance','Deductions','Net Salary','Payment Date','Status','Created Date','Updated Date']);
  ensureSheetIfMissing_(ss, 'CalendarEvents', ['ID','EventID','Type','Title','Start','End','RelatedID','Status','Notes','Created Date','Updated Date']);
  ensureSheetIfMissing_(ss, 'ProductCategories', ['ID','CategoryID','Category Type','Category Name','Status','Created Date','Updated Date']);
  ensureSheetIfMissing_(ss, 'Packages', ['ID','PackageID','Package Name','Description','Selling Price','Estimated Cost','Profit','Profit Margin','Status','Created Date','Updated Date']);
  ensureSheetIfMissing_(ss, 'PackageItems', ['ID','PackageItemID','PackageID','ProductID','Product Name','Quantity','Unit Cost','Unit Price','Subtotal','Created Date','Updated Date']);
  ensureSheetIfMissing_(ss, 'InventoryBatch', ['ID','BatchID','ProductID','Product Name','Batch Number','Purchase Date','Source','Quantity','Remaining Quantity','Unit Cost','Landed Cost','Storage Location','Created Date','Updated Date']);
  ensureSheetIfMissing_(ss, 'StorageLocations', ['ID','LocationID','Location Name','Description','Status','Created Date','Updated Date']);
  ensureSheetIfMissing_(ss, 'StockIn', ['ID','StockInID','ProductID','Product Name','Quantity','Purchase Cost','Shipping Cost','Other Charges','Source','Invoice Number','Batch Number','Storage Location','Transaction Date','Created Date','Updated Date']);
  ensureSheetIfMissing_(ss, 'StockOut', ['ID','StockOutID','ProductID','Product Name','Quantity','Unit Cost','BookingID','Reason','Batch Number','Storage Location','Transaction Date','Created Date','Updated Date']);



  const users = readRows_('Users');

  if (!users.length) {
    const salt = Utilities.getUuid().replace(/-/g, '');
    const row = {
      ID: uid_('USR'),
      UserID: uid_('USR'),
      'Full Name': APP.INITIAL_ADMIN.name,
      Email: APP.INITIAL_ADMIN.email.toLowerCase(),
      'Password Hash': hashPassword_(APP.INITIAL_ADMIN.password, salt),
      Salt: salt,
      Role: 'ADMIN',
      Status: 'Active',
      'Created Date': now_(),
      'Updated Date': now_(),
      'Last Login': ''
    };
    appendRow_('Users', row);
  }

  saveSettingRaw_('company_name', APP.NAME);
  saveSettingRaw_('version', APP.VERSION);

  return {
    ok: true,
    message: 'Phase 1 setup completed.',
    spreadsheetId: ss.getId(),
    adminEmail: APP.INITIAL_ADMIN.email
  };
}

/**
 * Called by the login page to verify the backend is reachable.
 */
function pingServer() {
  return {
    ok: true,
    name: APP.NAME,
    version: APP.VERSION,
    timestamp: now_()
  };
}

/**
 * Login.
 */
function login(email, password, rememberMe) {
  setupSystem();

  email = String(email || '').trim().toLowerCase();
  password = String(password || '');

  if (!email || !password) {
    return { ok: false, message: 'Email and password are required.' };
  }

  const users = readRows_('Users');
  const user = users.find(u =>
    String(u.Email || '').trim().toLowerCase() === email
  );

  if (!user) {
    return { ok: false, message: 'Invalid email or password.' };
  }

  if (String(user.Status || '').toLowerCase() !== 'active') {
    return { ok: false, message: 'This account is inactive.' };
  }

  const expected = String(user['Password Hash'] || '');
  const actual = hashPassword_(password, String(user.Salt || ''));

  if (expected !== actual) {
    return { ok: false, message: 'Invalid email or password.' };
  }

  const token = createSession_(user, !!rememberMe);

  const updated = {
    'Last Login': now_(),
    'Updated Date': now_()
  };
  updateById_('Users', user.ID, updated);

  logActivity_(user, 'LOGIN', 'Authentication', 'Successful login');

  return {
    ok: true,
    token: token,
    user: safeUser_(user)
  };
}

/**
 * Validate an existing session.
 */
function validateSession(token) {
  const session = getSession_(token);
  if (!session) return { ok: false, message: 'Session expired.' };
  return { ok: true, user: session.user };
}

/**
 * Logout.
 */
function logout(token) {
  if (token) CacheService.getScriptCache().remove('session_' + token);
  return { ok: true };
}


/**
 * Simple dashboard test payload.
 * This is intentionally small in Phase 1.
 */
function getDashboard(token) {
  const auth = requireAuth_(token);
  if (!auth.ok) return auth;

  const clients = getOrEmpty_('Clients');
  const suppliers = getOrEmpty_('Suppliers');
  const bookings = getOrEmpty_('Bookings');
  const products = getOrEmpty_('Products');
  const inventory = getOrEmpty_('Inventory');
  const payments = getOrEmpty_('Payments');
  const expenses = getOrEmpty_('Expenses');
  const payroll = getOrEmpty_('Payroll');

  const today = new Date();
  const todayKey = formatDateKey_(today);

  const upcomingBookings = bookings
    .filter(b => {
      const d = toDate_(b['Event Date']);
      return d && d >= startOfDay_(today) && String(b['Booking Status'] || '') !== 'Cancelled';
    })
    .sort((a, b) => toDate_(a['Event Date']) - toDate_(b['Event Date']));

  const todaysBookings = bookings.filter(b => {
    const d = toDate_(b['Event Date']);
    return d && formatDateKey_(d) === todayKey;
  });

  const totalSales = bookings.reduce((sum, b) => sum + number_(b['Final Amount']), 0);
  const amountCollected = payments.reduce((sum, p) => sum + number_(p.Amount), 0);
  const outstandingBalance = bookings.reduce((sum, b) => sum + number_(b.Balance), 0);

  const inventoryValue = inventory.reduce(
    (sum, i) => sum + (
      number_(i['Inventory Value']) ||
      (number_(i['Current Stock']) * number_(i['Landed Cost']))
    ),
    0
  );

  const lowStock = inventory.filter(i =>
    number_(i['Available Stock']) <= number_(i['Minimum Stock'])
  );

  const monthKey = Utilities.formatDate(
    today,
    Session.getScriptTimeZone(),
    'yyyy-MM'
  );

  const monthlyExpenses = expenses.reduce((sum, e) => {
    const d = toDate_(e.Date);
    return d && Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM') === monthKey
      ? sum + number_(e.Amount)
      : sum;
  }, 0);

  const monthlyPayroll = payroll.reduce((sum, p) => {
    const d = toDate_(p['Payment Date']);
    return d && Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM') === monthKey
      ? sum + number_(p['Net Salary'])
      : sum;
  }, 0);

  const currentYear = new Date().getFullYear();
  const monthNames = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

  const monthlyBookings = Array(12).fill(0);
  const monthlySales = Array(12).fill(0);

  bookings.forEach(b => {
    const d = toDate_(b['Event Date'] || b['Booking Date'] || b['Created Date']);
    if (!d || d.getFullYear() !== currentYear) return;

    const m = d.getMonth();
    monthlyBookings[m] += 1;

    // Use the booking's final amount. Older records may use Final Amount,
    // Total Amount Due, or Total Amount.
    monthlySales[m] += number_(
      b['Final Amount'] ??
      b['Total Amount Due'] ??
      b['Total Amount'] ??
      0
    );
  });

  const bookingMonths = monthNames.map((name, i) => ({
    month: name,
    count: monthlyBookings[i]
  }));

  const salesMonths = monthNames.map((name, i) => ({
    month: name,
    sales: monthlySales[i]
  }));

  const totalYearBookings = monthlyBookings.reduce((a,b) => a+b, 0);
  const totalYearSales = monthlySales.reduce((a,b) => a+b, 0);

  const bestBookingIndex = monthlyBookings.reduce(
    (best, value, index) => value > monthlyBookings[best] ? index : best, 0
  );

  const bestSalesIndex = monthlySales.reduce(
    (best, value, index) => value > monthlySales[best] ? index : best, 0
  );

  const worstBookingIndex = monthlyBookings.reduce(
    (worst, value, index) => value < monthlyBookings[worst] ? index : worst, 0
  );

  const worstSalesIndex = monthlySales.reduce(
    (worst, value, index) => value < monthlySales[worst] ? index : worst, 0
  );

  const recentActivity = getOrEmpty_('ActivityLogs')
    .sort((a, b) => toDate_(b.Timestamp) - toDate_(a.Timestamp))
    .slice(0, 8);

  return {
    ok: true,
    user: auth.user,
    metrics: {
      totalBookings: bookings.length,
      upcomingEvents: upcomingBookings.length,
      todaysEvents: todaysBookings.length,
      totalClients: clients.length,
      totalSuppliers: suppliers.length,
      totalProducts: products.length,
      lowStockItems: lowStock.length,
      totalSales,
      amountCollected,
      outstandingBalance,
      inventoryValue,
      monthlyExpenses,
      monthlyPayroll,
      netProfit: totalSales - monthlyExpenses - monthlyPayroll,
      currentYear,
      bookingMonths,
      salesMonths,
      totalYearBookings,
      totalYearSales,
      bestBookingMonth: { month: monthNames[bestBookingIndex], count: monthlyBookings[bestBookingIndex] },
      bestSalesMonth: { month: monthNames[bestSalesIndex], sales: monthlySales[bestSalesIndex] },
      slowestBookingMonth: { month: monthNames[worstBookingIndex], count: monthlyBookings[worstBookingIndex] },
      slowestSalesMonth: { month: monthNames[worstSalesIndex], sales: monthlySales[worstSalesIndex] }
    },
    upcoming: upcomingBookings.slice(0, 8),
    lowStock: lowStock.slice(0, 8),
    recentActivity
  };
}
