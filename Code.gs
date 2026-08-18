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
  VERSION: 'PHASE-1.2-STABLE',
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
  ensureSheetIfMissing_(ss, 'BookingCrew', ['ID','CrewID','BookingID','EmployeeID','Employee Name','Role','Show Assignment','Pay','Notes','Created Date','Updated Date']);
  ensureSheetIfMissing_(ss, 'Employees', ['ID','EmployeeID','Employee Name','Role','Employee Type','Contact Number','Salary','Status','Notes','Created Date','Updated Date']);
  ensureSheetIfMissing_(ss, 'EmployeeAdvances', ['ID','AdvanceID','EmployeeID','Employee Name','Type','Amount','Amount Paid','Outstanding Balance','Date','Purpose','Notes','Status','Created Date','Updated Date']);
  ensureSheetIfMissing_(ss, 'BookingExpenses', ['ID','ExpenseID','BookingID','Date','Category','Description','Amount','Created Date','Updated Date']);
  ensureSheetIfMissing_(ss, 'PaymentSchedules', ['ID','ScheduleID','BookingID','Milestone','Due Date','Amount','Paid','Balance','Status','Created Date','Updated Date']);
  ensureSheetIfMissing_(ss, 'AuditLog', ['ID','User','Action','Module','Record ID','Timestamp','Details']);
  ensureSheetIfMissing_(ss, 'Clients', ['ID','ClientID','Name','Phone','Email','Address','Notes','Status','Created Date','Updated Date']);
  ensureSheetIfMissing_(ss, 'Loans', ['ID','LoanID','Lender','Loan Amount','Amount Paid','Outstanding Balance','Date Borrowed','Due Date','Purpose','Notes','Created Date','Updated Date']);
  ensureSheetIfMissing_(ss, 'Bookings', ['ID','BookingID','ClientID','Client Name','Contact Information','Event Type','Event Name','Event Date','Event Time','Location','Subtotal','Discount Type','Discount Value','Discount Amount','Discount','Additional Charges','Tax Rate','Tax','Final Amount','Total Amount Due','DownPayment','Down Payment','Paid Amount','Balance','Payment Status','Booking Status','Archived','Notes','Created Date','Updated Date']);
  ensureSheetIfMissing_(ss, 'Products', ['ID','ProductID','Name','Source','Category','Unit','Purchase Cost','Selling Cost','Unit Cost','Selling Price','Minimum Stock','Status','Description','Created Date','Updated Date']);
  ensureColumnIfMissing_(ss.getSheetByName('Products'), 'Selling Cost');
  ensureColumnIfMissing_(ss.getSheetByName('Products'), 'Unit Cost');
  ensureSheetIfMissing_(ss, 'Inventory', ['ID','InventoryID','ProductID','Product Name','Current Stock','Reserved Stock','Available Stock','Minimum Stock','Unit Cost','Landed Cost','Selling Price','Inventory Value','Updated Date']);
  ensureSheetIfMissing_(ss, 'Payments', ['ID','PaymentID','BookingID','Payment Date','Milestone','Amount','Method','Reference','Notes','Status','Created Date','Updated Date']);
  ensureSheetIfMissing_(ss, 'Expenses', ['ID','ExpenseID','Date','Amount','Category','SupplierID','BookingID','Description','Approval Status','Created Date','Updated Date']);
  ensureSheetIfMissing_(ss, 'Suppliers', ['ID','SupplierID','Name','Contact Number','Email','Address','Notes','Status','Created Date','Updated Date']);
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
  if (token) PropertiesService.getScriptProperties().deleteProperty('session_' + token);
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
  const bookings = getOrEmpty_('Bookings');
  const products = getOrEmpty_('Products');
  const inventory = getOrEmpty_('Inventory');
  const payments = getOrEmpty_('Payments');
  const expenses = getOrEmpty_('Expenses');
  const payroll = getOrEmpty_('Payroll');

  // Cancelled bookings are retained in the database for audit/history,
  // but are excluded from operational totals and client/show overview metrics.
  const activeBookings = bookings.filter(b =>
    String(b['Booking Status'] || 'Pending').trim().toLowerCase() !== 'cancelled'
  );

  const today = new Date();
  const todayKey = formatDateKey_(today);

  const upcomingBookings = activeBookings
    .filter(b => {
      const d = toDate_(b['Event Date']);
      return d && d >= startOfDay_(today);
    })
    .sort((a, b) => toDate_(a['Event Date']) - toDate_(b['Event Date']));

  const todaysBookings = activeBookings.filter(b => {
    const d = toDate_(b['Event Date']);
    return d && formatDateKey_(d) === todayKey;
  });

  const totalSales = activeBookings.reduce((sum, b) =>
    sum + number_(b['Final Amount'] ?? b['Total Amount Due'] ?? b['Total Amount']), 0
  );

  const activeBookingIds = new Set(
    activeBookings.map(b => String(b.ID || b.BookingID || '').trim()).filter(Boolean)
  );

  const amountCollected = payments.reduce((sum, p) => {
    const bookingId = String(p.BookingID || '').trim();
    return (!bookingId || activeBookingIds.has(bookingId))
      ? sum + number_(p.Amount)
      : sum;
  }, 0);

  const outstandingBalance = activeBookings.reduce((sum, b) =>
    sum + number_(b.Balance ?? b['Outstanding Balance']), 0
  );

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
    const bookingId = String(e.BookingID || '').trim();
    const d = toDate_(e.Date);
    if (bookingId && !activeBookingIds.has(bookingId)) return sum;
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

  // Count unique clients with at least one non-cancelled booking.
  // Multiple active shows for the same client count as one client.
  const activeClientKeys = new Set();
  activeBookings.forEach(function(b){
    const id = String(b.ClientID || '').trim();
    const name = String(b['Client Name'] || '').trim().toLowerCase();
    activeClientKeys.add(id || name);
  });
  const effectiveTotalClients = activeClientKeys.size;

  // Dashboard status counts are derived directly from current active bookings.
  const totalBooked = activeBookings.filter(function(b){
    return String(b['Booking Status'] || '').trim().toLowerCase() === 'booked';
  }).length;
  const totalPending = activeBookings.filter(function(b){
    return String(b['Booking Status'] || 'pending').trim().toLowerCase() === 'pending';
  }).length;

  const currentYear = new Date().getFullYear();
  const monthNames = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

  const monthlyBookings = Array(12).fill(0);
  const monthlySales = Array(12).fill(0);

  activeBookings.forEach(b => {
    const d = toDate_(b['Event Date'] || b['Booking Date'] || b['Created Date']);
    if (!d || d.getFullYear() !== currentYear) return;

    const m = d.getMonth();
    monthlyBookings[m] += 1;
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

  // Calculate best/slowest months with explicit loops so the dashboard has
  // no dependency on reducer accumulator names or closures.
  let bestBookingIndex = 0;
  let worstBookingIndex = 0;
  let bestSalesIndex = 0;
  let worstSalesIndex = 0;

  for (let i = 1; i < 12; i++) {
    if (monthlyBookings[i] > monthlyBookings[bestBookingIndex]) bestBookingIndex = i;
    if (monthlyBookings[i] < monthlyBookings[worstBookingIndex]) worstBookingIndex = i;
    if (monthlySales[i] > monthlySales[bestSalesIndex]) bestSalesIndex = i;
    if (monthlySales[i] < monthlySales[worstSalesIndex]) worstSalesIndex = i;
  }

  const recentActivity = getOrEmpty_('ActivityLogs')
    .sort((a, b) => toDate_(b.Timestamp) - toDate_(a.Timestamp))
    .slice(0, 8);

  return {
    ok: true,
    user: auth.user,
    metrics: {
      totalBookings: activeBookings.length,
      totalBooked: totalBooked,
      totalPending: totalPending,
      upcomingEvents: upcomingBookings.length,
      todaysEvents: todaysBookings.length,
      totalClients: effectiveTotalClients,
      totalSuppliers: getOrEmpty_('Suppliers').filter(s =>
        String(s.Status || 'Active').toLowerCase() !== 'inactive'
      ).length,
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



/* =========================
   CLIENTS API
   ========================= */

function getClients(token) {
  const auth = requireAuth_(token);
  if (!auth.ok) return auth;
  return { ok:true, clients:getOrEmpty_('Clients') };
}

function saveClient(token, payload) {
  const auth = requireAuth_(token);
  if (!auth.ok) return auth;
  payload = payload || {};

  const name = String(payload.name || '').trim();
  const phone = String(payload.phone || '').trim();
  const email = String(payload.email || '').trim();
  const address = String(payload.address || '').trim();
  const notes = String(payload.notes || '').trim();

  if (!name) return {ok:false, message:'Client name is required.'};
  if (email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    return {ok:false, message:'Please enter a valid email address.'};
  }

  const id = payload.id ? String(payload.id) : makeId_('CLI');
  const existing = payload.id ? findRowById_('Clients', id) : null;
  const now = new Date();

  const duplicate = getOrEmpty_('Clients').find(c =>
    String(c.Name || '').trim().toLowerCase() === name.toLowerCase() &&
    String(c.Phone || '').trim() === phone &&
    String(c.ID || '') !== id
  );
  if (duplicate) return {ok:false, message:'A client with the same name and phone already exists.'};

  const record = {
    ID:id,
    ClientID:id,
    Name:name,
    Phone:phone,
    Email:email,
    Address:address,
    Notes:notes,
    'Created Date':existing ? existing['Created Date'] : now,
    'Updated Date':now
  };

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    upsertRow_('Clients', record);
    audit_(auth.user, existing ? 'UPDATE' : 'CREATE', 'Clients', id, record);
    return {
      ok:true,
      message:existing ? 'Client updated successfully.' : 'Client created successfully.',
      client:record
    };
  } finally {
    lock.releaseLock();
  }
}

function deactivateClient(token, clientId) {
  const auth = requireAuth_(token);
  if (!auth.ok) return auth;

  const id = String(clientId || '').trim();
  if (!id) return {ok:false, message:'Client ID is required.'};

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const client = findRowById_('Clients', id);
    if (!client) return {ok:false, message:'Client not found.'};

    client.Status = 'Inactive';
    client['Updated Date'] = new Date();
    upsertRow_('Clients', client);
    audit_(auth.user, 'DEACTIVATE', 'Clients', id, client);

    return {ok:true, message:'Client marked inactive.'};
  } finally {
    lock.releaseLock();
  }
}




/* =========================
   CLIENT WORKSPACE
   One central entry point for client/show information.
   ========================= */

function saveClientWorkspace(token, payload) {
  const auth = requireAuth_(token);
  if (!auth.ok) return auth;
  payload = payload || {};

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    const now = new Date();

    // Client
    const clientName = String(payload.clientName || '').trim();
    if (!clientName) return {ok:false, message:'Client name is required.'};

    const clientId = String(payload.clientId || '') || makeId_('CLI');
    const existingClient = findRowById_('Clients', clientId);
    const client = {
      ID:clientId,
      ClientID:clientId,
      Name:clientName,
      Phone:String(payload.phone || '').trim(),
      Email:String(payload.email || '').trim(),
      Address:String(payload.address || '').trim(),
      Notes:String(payload.clientNotes || '').trim(),
      'Created Date':existingClient ? existingClient['Created Date'] : now,
      'Updated Date':now
    };
    upsertRow_('Clients', client);

    // Booking
    const bookingId = String(payload.bookingId || '') || makeId_('BKG');
    const existingBooking = findRowById_('Bookings', bookingId);

    const itemRows = Array.isArray(payload.products) ? payload.products : [];
    let subtotal = 0;
    itemRows.forEach(i => {
      const qty = Math.max(0, number_(i.quantity));
      const price = Math.max(0, number_(i.unitPrice));
      subtotal += qty * price;
    });

    const discount = Math.max(0, number_(payload.discount));
    const additionalCharges = 0; // intentionally removed from the business workflow
    const taxableBase = Math.max(0, subtotal - discount);
    const tax = taxableBase * 0.05;
    const finalAmount = taxableBase + tax;

    const payments = Math.max(0, number_(payload.payments));
    const balance = Math.max(0, finalAmount - payments);

    const booking = {
      ID:bookingId,
      BookingID:bookingId,
      ClientID:clientId,
      'Client Name':clientName,
      'Contact Information':client.Phone,
      'Event Type':String(payload.eventType || ''),
      'Event Name':String(payload.eventName || ''),
      'Event Date':String(payload.eventDate || ''),
      'Event Time':normalizeEventTime_(payload.eventTime),
      Location:String(payload.location || ''),
      'Subtotal':subtotal,
      Discount:discount,
      'Additional Charges':additionalCharges,
      'Tax Rate':0.05,
      Tax:tax,
      'Total Amount Due':finalAmount,
      'Down Payment':payments,
      Balance:balance,
      'Payment Status':balance <= 0 ? 'Paid' : payments > 0 ? 'Partial' : 'Unpaid',
      'Booking Status':String(payload.bookingStatus || 'Pending'),
      Notes:String(payload.notes || ''),
      'Created Date':existingBooking ? existingBooking['Created Date'] : now,
      'Updated Date':now
    };

    // If an older installation has a different booking header, upsertRow_ only writes known headers.
    upsertRow_('Bookings', booking);
    setBookingEventTimeText_(bookingId, booking['Event Time']);
    recalculateCrewPayForBooking_(bookingId);

    // Booking items: replace linked items on save to avoid duplicate rows.
    const itemSheet = getSheet_('BookingItems');
    if (itemSheet) {
      const all = getOrEmpty_('BookingItems');
      all.filter(r => String(r.BookingID || '') === bookingId)
        .forEach(r => deleteRowById_('BookingItems', String(r.ID || r.BookingItemID || '')));

      itemRows.forEach(i => {
        const qty = Math.max(0, number_(i.quantity));
        const unitPrice = Math.max(0, number_(i.unitPrice));
        if (!i.productId || qty <= 0) return;

        upsertRow_('BookingItems', {
          ID:makeId_('BIT'),
          BookingItemID:makeId_('BIT'),
          BookingID:bookingId,
          ProductID:String(i.productId),
          'Product Name':String(i.productName || ''),
          Quantity:qty,
          'Unit Price':unitPrice,
          Subtotal:qty * unitPrice,
          'Created Date':now,
          'Updated Date':now
        });
      });
    }

    audit_(auth.user, existingBooking ? 'UPDATE' : 'CREATE', 'Client Workspace', bookingId, {
      clientId:clientId,
      bookingId:bookingId,
      subtotal:subtotal,
      tax:tax,
      total:finalAmount,
      balance:balance
    });

    return {
      ok:true,
      message:existingBooking ? 'Client/show updated successfully.' : 'Client/show created successfully.',
      client:client,
      booking:booking,
      calculations:{
        subtotal:subtotal,
        discount:discount,
        taxRate:0.05,
        tax:tax,
        finalAmount:finalAmount,
        payments:payments,
        balance:balance
      }
    };
  } finally {
    lock.releaseLock();
  }
}

function getClientWorkspace(token, bookingId) {
  const auth = requireAuth_(token);
  if (!auth.ok) return auth;

  const booking = findRowById_('Bookings', String(bookingId || ''));
  if (!booking) return {ok:false, message:'Booking/show not found.'};

  const client = findRowById_('Clients', String(booking.ClientID || ''));
  const items = getOrEmpty_('BookingItems').filter(r => String(r.BookingID || '') === String(booking.ID || booking.BookingID || ''));

  return {ok:true, client:client || {}, booking:booking, products:items};
}

/* =========================
   CLIENT OPTIONS
   ========================= */

function getClientOptions(token) {
  const auth = requireAuth_(token);
  if (!auth.ok) return auth;
  const clients = getOrEmpty_('Clients')
    .filter(c => String(c.Name || '').trim())
    .map(c => ({
      ID: String(c.ID || c.ClientID || ''),
      Name: String(c.Name || ''),
      Phone: String(c.Phone || ''),
      Email: String(c.Email || ''),
      Address: String(c.Address || ''),
      Notes: String(c.Notes || '')
    }));
  return {ok:true, clients:clients};
}



/* =========================
   SHARED SHEET HELPERS
   ========================= */

function getSpreadsheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('No active spreadsheet found. Make sure this Apps Script project is bound to your ERP Google Sheet.');
  }
  return ss;
}

function getSheet_(sheetName) {
  const sheet = getSpreadsheet_().getSheetByName(sheetName);
  return sheet || null;
}

function getHeaders_(sheet) {
  if (!sheet || sheet.getLastColumn() < 1) return [];
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

function deleteRowById_(sheetName, id) {
  const sheet = getSheet_(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return false;

  const headers = getHeaders_(sheet);
  const idCol = headers.indexOf('ID') + 1;
  if (!idCol) return false;

  const values = sheet.getRange(2, idCol, sheet.getLastRow() - 1, 1).getValues().flat();
  const idx = values.findIndex(v => String(v) === String(id));
  if (idx < 0) return false;

  sheet.deleteRow(idx + 2);
  return true;
}

/* =========================
   LOANS API
   ========================= */

function getLoans(token) {
  const auth = requireAuth_(token);
  if (!auth.ok) return auth;
  const loans = getOrEmpty_('Loans').map(l => {
    const amount = number_(l['Loan Amount']);
    const paid = number_(l['Amount Paid']);
    return Object.assign({}, l, {
      'Loan Amount': amount,
      'Amount Paid': paid,
      'Outstanding Balance': Math.max(0, amount - paid)
    });
  });
  return {ok:true, loans:loans};
}

function saveLoan(token, payload) {
  const auth = requireAuth_(token);
  if (!auth.ok) return auth;
  payload = payload || {};

  const lender = String(payload.lender || '').trim();
  const amount = number_(payload.amount);
  const dateBorrowed = payload.dateBorrowed ? new Date(payload.dateBorrowed) : new Date();

  if (!lender) return {ok:false, message:'Lender / source of loan is required.'};
  if (amount <= 0) return {ok:false, message:'Loan amount must be greater than zero.'};

  const id = payload.id ? String(payload.id) : makeId_('LOAN');
  const existing = payload.id ? findRowById_('Loans', id) : null;
  const paid = Math.max(0, number_(payload.amountPaid));
  if (paid > amount) return {ok:false, message:'Amount paid cannot exceed the loan amount.'};

  const now = new Date();
  const record = {
    ID:id,
    LoanID:id,
    Lender:lender,
    'Loan Amount':amount,
    'Amount Paid':paid,
    'Outstanding Balance':Math.max(0, amount - paid),
    'Date Borrowed':dateBorrowed,
    'Due Date':payload.dueDate ? new Date(payload.dueDate) : '',
    Purpose:String(payload.purpose || ''),
    Notes:String(payload.notes || ''),
    'Created Date':existing ? existing['Created Date'] : now,
    'Updated Date':now
  };

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    if (!getSheet_('Loans')) {
      ensureSheet_(getSpreadsheet_(), 'Loans',
        ['ID','LoanID','Lender','Loan Amount','Amount Paid','Outstanding Balance','Date Borrowed','Due Date','Purpose','Notes','Created Date','Updated Date']
      );
    }
    upsertRow_('Loans', record);
    audit_(auth.user, existing ? 'UPDATE' : 'CREATE', 'Loans', id, record);
    return sanitizeForClient_({ok:true, message:existing ? 'Loan updated successfully.' : 'Loan recorded successfully.', loan:record});
  } finally {
    lock.releaseLock();
  }
}

function recordLoanPayment(token, payload) {
  const auth = requireAuth_(token);
  if (!auth.ok) return auth;
  payload = payload || {};

  const loanId = String(payload.loanId || '').trim();
  const payment = number_(payload.amount);
  if (!loanId) return {ok:false, message:'Loan is required.'};
  if (payment <= 0) return {ok:false, message:'Payment amount must be greater than zero.'};

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const loan = findRowById_('Loans', loanId);
    if (!loan) return {ok:false, message:'Loan not found.'};

    const amount = number_(loan['Loan Amount']);
    const paid = number_(loan['Amount Paid']);
    const outstanding = Math.max(0, amount - paid);

    if (payment > outstanding) {
      return {ok:false, message:'Payment cannot exceed the outstanding balance.'};
    }

    loan['Amount Paid'] = paid + payment;
    loan['Outstanding Balance'] = Math.max(0, amount - loan['Amount Paid']);
    loan['Updated Date'] = new Date();
    upsertRow_('Loans', loan);

    // Reuse Payments sheet as the financial transaction history.
    const paymentId = makeId_('LPAY');
    const pSheet = getSpreadsheet_().getSheetByName('Payments');
    if (pSheet) {
      appendRow_('Payments', {
        ID:paymentId,
        PaymentID:paymentId,
        BookingID:'',
        'Payment Date':payload.date ? new Date(payload.date) : new Date(),
        Milestone:'Loan Payment',
        Amount:payment,
        Method:String(payload.method || ''),
        Reference:loanId,
        Notes:'Loan payment to ' + loan.Lender,
        Status:'Completed',
        'Created Date':new Date(),
        'Updated Date':new Date()
      });
    }

    audit_(auth.user, 'LOAN_PAYMENT', 'Loans', loanId, {amount:payment, remaining:loan['Outstanding Balance']});
    return sanitizeForClient_({ok:true, message:'Loan payment recorded successfully.', loan:loan});
  } finally {
    lock.releaseLock();
  }
}




/* =========================
   CLIENT/SHOW LIST API
   ========================= */

function getBookings(token) {
  const auth = requireAuth_(token);
  if (!auth.ok) return auth;
  const bookings = getOrEmpty_('Bookings').map(function(row){
    const copy = Object.assign({}, row);
    copy['Event Time'] = getPersistedBookingEventTime_(String(row.ID || row.BookingID || '')) ||
      normalizeEventTime_(row['Event Time'] || row.EventTime || row.Time || '');
    return copy;
  });
  return {ok:true, bookings:bookings};
}

function createClientShow(token, payload) {
  const auth = requireAuth_(token);
  if (!auth.ok) return auth;
  payload = payload || {};

  const clientName = String(payload.clientName || '').trim();
  if (!clientName) return {ok:false, message:'Client name is required.'};
  if (!String(payload.eventDate || '').trim()) return {ok:false, message:'Event date is required.'};

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    const now = new Date();
    const clientId = makeId_('CLI');
    const bookingId = makeId_('BKG');

    const client = {
      ID:clientId,
      ClientID:clientId,
      Name:clientName,
      Phone:String(payload.phone || '').trim(),
      Email:String(payload.email || '').trim(),
      Address:String(payload.address || '').trim(),
      Notes:String(payload.notes || '').trim(),
      'Created Date':now,
      'Updated Date':now
    };

    const booking = {
      ID:bookingId,
      BookingID:bookingId,
      ClientID:clientId,
      'Client Name':clientName,
      'Contact Information':client.Phone,
      'Event Type':String(payload.eventType || 'Other'),
      'Event Name':String(payload.eventName || ''),
      'Event Date':String(payload.eventDate || ''),
      'Event Time':normalizeEventTime_(payload.eventTime),
      Location:String(payload.location || ''),
      Subtotal:0,
      'Discount Type':'Fixed',
      'Discount Value':0,
      'Discount Amount':0,
      'Additional Charges':0,
      Tax:0,
      'Tax Rate':0.05,
      'Final Amount':0,
      DownPayment:0,
      'Paid Amount':0,
      Balance:0,
      'Payment Status':'Unpaid',
      'Booking Status':String(payload.bookingStatus || 'Pending'),
      Notes:String(payload.notes || ''),
      'Created Date':now,
      'Updated Date':now
    };

    upsertRow_('Clients', client);
    upsertRow_('Bookings', booking);
    setBookingEventTimeText_(bookingId, payload.eventTime);
    booking['Event Time'] = getPersistedBookingEventTime_(bookingId) || normalizeEventTime_(payload.eventTime || payload.event_time || '');
    recalculateCrewPayForBooking_(bookingId);

    audit_(auth.user, 'CREATE', 'Client / Show', bookingId, {
      clientId:clientId,
      clientName:clientName
    });

    // google.script.run cannot safely return raw Date objects to the browser.
    // Convert the response to JSON-safe values so a successful database write
    // is reported as success instead of appearing as a frontend failure.
    return {
      ok:true,
      client:{
        ID:String(client.ID || ''),
        ClientID:String(client.ClientID || client.ID || ''),
        Name:String(client.Name || ''),
        Phone:String(client.Phone || ''),
        Email:String(client.Email || ''),
        Address:String(client.Address || ''),
        Notes:String(client.Notes || ''),
        'Created Date':safeValue_(client['Created Date']),
        'Updated Date':safeValue_(client['Updated Date'])
      },
      booking:{
        ID:String(booking.ID || ''),
        BookingID:String(booking.BookingID || booking.ID || ''),
        ClientID:String(booking.ClientID || ''),
        'Client Name':String(booking['Client Name'] || ''),
        'Contact Information':String(booking['Contact Information'] || ''),
        'Event Type':String(booking['Event Type'] || ''),
        'Event Name':String(booking['Event Name'] || ''),
        'Event Date':String(booking['Event Date'] || ''),
        'Event Time':String(booking['Event Time'] || ''),
        Location:String(booking.Location || ''),
        Subtotal:number_(booking.Subtotal),
        'Discount Type':String(booking['Discount Type'] || ''),
        'Discount Value':number_(booking['Discount Value']),
        'Discount Amount':number_(booking['Discount Amount']),
        'Additional Charges':number_(booking['Additional Charges']),
        Tax:number_(booking.Tax),
        'Tax Rate':number_(booking['Tax Rate']),
        'Final Amount':number_(booking['Final Amount']),
        DownPayment:number_(booking.DownPayment),
        'Paid Amount':number_(booking['Paid Amount']),
        Balance:number_(booking.Balance),
        'Payment Status':String(booking['Payment Status'] || ''),
        'Booking Status':String(booking['Booking Status'] || 'Pending'),
        Notes:String(booking.Notes || ''),
        'Created Date':safeValue_(booking['Created Date']),
        'Updated Date':safeValue_(booking['Updated Date'])
      }
    };
  } finally {
    lock.releaseLock();
  }
}

function recalculateCrewPayForBooking_(bookingId) {
  const id = String(bookingId || '').trim();
  if (!id) return;

  const booking = findRowById_('Bookings', id);
  if (!booking) return;

  const showTotal = number_(
    booking['Final Amount'] ??
    booking['Total Amount Due'] ??
    booking['Total Amount'] ??
    0
  );

  const crew = getOrEmpty_('BookingCrew')
    .filter(c => String(c.BookingID || '') === id);

  crew.forEach(c => {
    // Only apply the automatic 5% rule when the show total is above ₱60,000.
    // For ₱60,000 and below, preserve the manually entered show pay.
    if (showTotal > 60000) {
      const rate = number_(c['Pay Rate']) || 5;
      c['Pay Rate'] = rate;
      c.Pay = showTotal * rate / 100;
      c['Pay Mode'] = 'Percentage';
    } else {
      c['Pay Mode'] = 'Manual';
      if (c.Pay === '' || c.Pay === null || c.Pay === undefined) c.Pay = 0;
    }
    c['Updated Date'] = new Date();
    upsertRow_('BookingCrew', c);
  });
}


function getCalendarData(token, year, month) {
  const auth = requireAuth_(token);
  if (!auth.ok) return auth;

  const y = Number(year);
  const m = Number(month);
  if (!Number.isInteger(y) || !Number.isInteger(m) || m < 0 || m > 11) {
    return { ok:false, message:'Invalid calendar period.' };
  }

  const bookings = getOrEmpty_('Bookings');
  const schedules = getOrEmpty_('PaymentSchedules');
  const calendarEvents = getOrEmpty_('CalendarEvents');
  const clients = getOrEmpty_('Clients');

  const clientMap = new Map();
  clients.forEach(function(c){
    const id = String(c.ID || c.ClientID || '').trim();
    if (id) clientMap.set(id, c);
  });

  const start = new Date(y, m, 1);
  const end = new Date(y, m + 1, 1);
  const events = [];

  bookings.forEach(function(b){
    const d = toDate_(b['Event Date']);
    if (!d || d < start || d >= end) return;
    if (String(b['Booking Status'] || '').toLowerCase() === 'cancelled') return;

    const client = clientMap.get(String(b.ClientID || '').trim()) || {};
    events.push({
      id: String(b.ID || b.BookingID || ''),
      type: 'booking',
      title: String(b['Event Name'] || b['Event Type'] || 'Booking'),
      date: Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      // Calendar intentionally uses DATE ONLY. Event time stays in Client/Show Overview.
      time: '',
      client: String(b['Client Name'] || client.Name || ''),
      status: String(b['Booking Status'] || 'Pending'),
      location: String(b.Location || ''),
      bookingId: String(b.ID || b.BookingID || ''),
      amount: number_(b['Final Amount'] ?? b['Total Amount Due'] ?? b['Total Amount'] ?? 0),
      details: 'Client / Show'
    });
  });

  schedules.forEach(function(s){
    const d = toDate_(s['Due Date']);
    if (!d || d < start || d >= end) return;
    const relatedBookingId = String(s.BookingID || '').trim();
    const booking = bookings.find(function(b){
      return String(b.ID || b.BookingID || '').trim() === relatedBookingId;
    }) || {};
    const client = clientMap.get(String(booking.ClientID || '').trim()) || {};
    events.push({
      id: String(s.ID || s.ScheduleID || ''),
      type: 'payment',
      title: String(s.Milestone || 'Payment Due'),
      date: Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      time: '',
      client: String(booking['Client Name'] || client.Name || ''),
      status: String(s.Status || (String(s.Paid || '').toLowerCase() === 'true' ? 'Paid' : 'Due')),
      location: '',
      bookingId: relatedBookingId,
      amount: number_(s.Amount),
      balance: number_(s.Balance),
      details: 'Payment Schedule'
    });
  });

  calendarEvents.forEach(function(e){
    const d = toDate_(e.Start);
    if (!d || d < start || d >= end) return;
    events.push({
      id: String(e.ID || e.EventID || ''),
      type: 'operation',
      title: String(e.Title || 'Event'),
      date: Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      time: Utilities.formatDate(d, Session.getScriptTimeZone(), 'h:mm a'),
      client: '',
      status: String(e.Status || 'Scheduled'),
      location: '',
      bookingId: String(e.RelatedID || ''),
      amount: 0,
      details: String(e.Notes || 'Calendar Event')
    });
  });

  events.sort(function(a,b){
    const da = new Date(a.date + (a.time ? 'T' + a.time : 'T00:00:00'));
    const db = new Date(b.date + (b.time ? 'T' + b.time : 'T00:00:00'));
    return da - db;
  });

  return { ok:true, year:y, month:m, events:events };
}

function getShowWorkspace(token, bookingId, skipRecalculate) {
  const auth = requireAuth_(token);
  if (!auth.ok) return auth;

  const id = String(bookingId || '').trim();
  if (!id) return {ok:false, message:'Booking / show ID is required.'};

  const booking = findRowById_('Bookings', id);
  if (!booking) return {ok:false, message:'Client/show not found.'};

  // Skip recalculation for view-only loads when requested. This avoids a costly
  // write/read round-trip every time the user opens or switches tabs.
  if (!skipRecalculate) {
    recalculateBookingAmountFromProductUsage_(id);
    const refreshedBooking = findRowById_('Bookings', id);
    if (refreshedBooking) Object.assign(booking, refreshedBooking);
  }

  const persistedEventTime = getPersistedBookingEventTime_(id) ||
    normalizeEventTime_(
      booking['Event Time'] ||
      booking.EventTime ||
      booking['Event Time '] ||
      booking.Time ||
      booking['event time'] ||
      booking['EVENT TIME'] ||
      ''
    );
  booking['Event Time'] = persistedEventTime;
  booking.EventTime = persistedEventTime;
  booking.Time = persistedEventTime;
  if (persistedEventTime) setBookingEventTimeText_(id, persistedEventTime);

  // Resolve the client by the booking's ClientID FIRST.
  // This is the authoritative relationship and prevents multiple bookings
  // with the same phone/name from opening the wrong client.
  const clientRows = getOrEmpty_('Clients');
  const bookingClientId = String(booking.ClientID || booking.ClientId || '').trim();
  const bookingPhone = String(booking['Contact Information'] || '').trim();
  const bookingName = String(booking['Client Name'] || '').trim().toLowerCase();
  let client = null;

  if (bookingClientId) {
    client = clientRows.find(function(row){
      return String(row.ID || '').trim() === bookingClientId ||
             String(row.ClientID || '').trim() === bookingClientId;
    }) || null;
  }

  // Legacy fallback only when ClientID cannot resolve.
  if (!client && bookingPhone) {
    client = clientRows.find(function(row){
      return String(row.Phone || '').trim() === bookingPhone;
    }) || null;
  }

  if (!client && bookingName) {
    client = clientRows.find(function(row){
      return String(row.Name || '').trim().toLowerCase() === bookingName;
    }) || null;
  }

  client = client || {};

  // Always surface the booking's saved contact as a fallback so Overview reflects
  // what was entered in the Client/Show form even if a legacy client row is incomplete.
  if (!String(client.Phone || '').trim() && bookingPhone) client.Phone = bookingPhone;
  if (!String(client.Name || '').trim() && booking['Client Name']) client.Name = String(booking['Client Name']);

  // If we resolved a better client record, repair the booking relationship.
  const resolvedClientId = String(client.ID || client.ClientID || '').trim();
  if (resolvedClientId && resolvedClientId !== bookingClientId) {
    booking.ClientID = resolvedClientId;
  }

  const bookingKey = String(booking.ID || booking.BookingID || id);

  let usage = getOrEmpty_('BookingUsage')
    .filter(function(u){
      return String(u.BookingID || '').trim() === bookingKey;
    });
  usage = decorateUsageStockFields_(usage);
  const usageTotals = calculateShowUsageTotals_(usage);

  const showTotal = number_(
    booking['Final Amount'] ??
    booking['Total Amount Due'] ??
    booking['Total Amount'] ??
    0
  );

  const crew = getOrEmpty_('BookingCrew')
    .filter(function(c){
      return String(c.BookingID || '').trim() === bookingKey;
    })
    .map(function(c){
      const row = Object.assign({}, c);
      const storedPay = number_(c.Pay);
      const rate = number_(c['Pay Rate']) || 5;

      row.Pay = showTotal > 60000
        ? showTotal * rate / 100
        : storedPay;

      row['Pay Mode'] = showTotal > 60000
        ? 'Percentage'
        : 'Manual';

      return row;
    });

  const crewCost = crew.reduce(function(sum, c){ return sum + number_(c.Pay); }, 0);

  const payments = getOrEmpty_('Payments')
    .filter(function(p){ return String(p.BookingID || '').trim() === bookingKey; })
    .sort(function(a,b){ return (toDate_(b['Payment Date']) || 0) - (toDate_(a['Payment Date']) || 0); });

  const expenses = getOrEmpty_('Expenses')
    .filter(function(e){ return String(e.BookingID || '').trim() === bookingKey; })
    .sort(function(a,b){ return (toDate_(b.Date) || 0) - (toDate_(a.Date) || 0); });

  const expenseTotal = expenses.reduce(function(sum, e){ return sum + number_(e.Amount); }, 0);

  // Event Time is already canonical 12-hour text. Do not convert it back to HH:mm.

  return {
    ok:true,
    client:client,
    booking:booking,
    eventTime:persistedEventTime || '',
    usage:usage,
    crew:crew,
    payments:payments,
    expenses:expenses,
    totals:Object.assign({}, usageTotals, {
      crewCost:crewCost,
      expenseTotal:expenseTotal
    })
  };
}



function decorateUsageStockFields_(usageRows) {
  const rows = Array.isArray(usageRows) ? usageRows.map(function(r){ return Object.assign({}, r); }) : [];
  const stockIns = getOrEmpty_('StockIn');
  const stockOuts = getOrEmpty_('StockOut');
  const usageOrder = rows.map(function(row, index){
    return {
      index:index,
      itemId:String(row.ItemID || '').trim(),
      usedAt:toDate_(row['Used Date'] || row['Created Date']) || null
    };
  });

  rows.forEach(function(row, index){
    if (String(row['Usage Type'] || '') !== 'Product') return;

    const itemId = String(row.ItemID || '').trim();
    const usedAt = usageOrder[index].usedAt;
    const qty = number_(row.Quantity);

    let beginning = null;

    if (itemId && usedAt) {
      const stockInBefore = stockIns
        .filter(function(x){
          return String(x.ProductID || '').trim() === itemId &&
            (toDate_(x['Transaction Date'] || x['Created Date']) || new Date(0)) <= usedAt;
        })
        .reduce(function(sum, x){ return sum + number_(x.Quantity); }, 0);

      const stockOutBefore = stockOuts
        .filter(function(x){
          return String(x.ProductID || '').trim() === itemId &&
            (toDate_(x['Transaction Date'] || x['Created Date']) || new Date(0)) < usedAt;
        })
        .reduce(function(sum, x){ return sum + number_(x.Quantity); }, 0);

      beginning = Math.max(0, stockInBefore - stockOutBefore);

      // Multiple usage rows can share the same timestamp. Apply earlier
      // rows from this same usage payload in display order.
      for (let i = 0; i < index; i++) {
        const prev = usageOrder[i];
        if (prev.itemId === itemId && prev.usedAt && prev.usedAt.getTime() === usedAt.getTime()) {
          beginning = Math.max(0, beginning - number_(rows[i].Quantity));
        }
      }
    }

    if (beginning === null) {
      // Legacy fallback when a usage row has no usable timestamp.
      const laterUsage = rows
        .filter(function(x, i){
          return i >= index && String(x['Usage Type'] || '') === 'Product' && String(x.ItemID || '').trim() === itemId;
        })
        .reduce(function(sum, x){ return sum + number_(x.Quantity); }, 0);
      const inv = getOrEmpty_('Inventory').find(function(x){
        return String(x.ProductID || '').trim() === itemId || String(x.ItemID || '').trim() === itemId;
      });
      beginning = Math.max(0, number_(inv && (inv['Available Stock'] ?? inv['Current Stock'])) + laterUsage - qty);
    }

    row['Beginning Stock'] = beginning;
    row['Remaining Stock'] = Math.max(0, beginning - qty);
  });

  return rows;
}

function getBookingCrew(token, bookingId) {
  const auth = requireAuth_(token);
  if (!auth.ok) return auth;

  const id = String(bookingId || '').trim();
  if (!id) return {ok:false, message:'Booking / show ID is required.'};

  const booking = findRowById_('Bookings', id);
  const showTotal = booking ? number_(
    booking['Final Amount'] ??
    booking['Total Amount Due'] ??
    booking['Total Amount'] ??
    0
  ) : 0;

  const crew = getOrEmpty_('BookingCrew')
    .filter(r => String(r.BookingID || '').trim() === id)
    .map(function(r){
      const row = Object.assign({}, r);
      const storedPay = number_(r.Pay);
      const rate = number_(r['Pay Rate']) || 5;
      row.Pay = showTotal > 60000
        ? showTotal * rate / 100
        : storedPay;
      row['Pay Mode'] = showTotal > 60000 ? 'Percentage' : 'Manual';
      return row;
    });

  return {ok:true, crew:crew};
}

/* =========================
   EMPLOYEES / CREW / PAYROLL
   ========================= */

function getEmployees(token) {
  const ss = getSpreadsheet_();
  ensureSheetIfMissing_(ss, 'Employees', ['ID','EmployeeID','Employee Name','Role','Employee Type','Contact Number','Salary','Status','Notes','Created Date','Updated Date']);
  ensureSheetIfMissing_(ss, 'EmployeeAdvances', ['ID','AdvanceID','EmployeeID','Employee Name','Type','Amount','Amount Paid','Outstanding Balance','Date','Purpose','Notes','Status','Created Date','Updated Date']);

  const auth = requireAuth_(token);
  if (!auth.ok) return auth;

  const employees = getOrEmpty_('Employees');
  const crews = getOrEmpty_('BookingCrew');
  const advances = getOrEmpty_('EmployeeAdvances');
  const bookings = getOrEmpty_('Bookings');
  const now = new Date();

  const result = employees.map(e => {
    const employeeId = String(e.EmployeeID || e.ID || '').trim();

    const assignments = crews.filter(c =>
      String(c.EmployeeID || '').trim() === employeeId
    );

    const validAssignments = assignments.filter(c => {
      const booking = bookings.find(b => String(b.ID || b.BookingID || '') === String(c.BookingID || ''));
      if (!booking) return false;
      return String(booking['Booking Status'] || 'Booked') !== 'Cancelled';
    });

    const showsThisMonth = validAssignments.filter(c => {
      const booking = bookings.find(b => String(b.ID || b.BookingID || '') === String(c.BookingID || ''));
      const d = booking ? toDate_(booking['Event Date']) : null;
      return d && d.getMonth() === now.getMonth() && d.getFullYear() === now.getFullYear();
    }).length;

    const showPayThisMonth = validAssignments.reduce((sum, c) => {
      const booking = bookings.find(b => String(b.ID || b.BookingID || '') === String(c.BookingID || ''));
      if (!booking) return sum;
      const d = toDate_(booking['Event Date']);
      if (!d || d.getMonth() !== now.getMonth() || d.getFullYear() !== now.getFullYear()) return sum;
      const showTotal = number_(booking['Final Amount'] ?? booking['Total Amount Due'] ?? booking['Total Amount'] ?? 0);
      const rate = number_(c['Pay Rate']) || (showTotal > 60000 ? 5 : 0);
      const pay = showTotal > 60000 ? showTotal * rate / 100 : number_(c.Pay);
      return sum + pay;
    }, 0);

    const employeeAdvances = advances.filter(a => String(a.EmployeeID || '') === employeeId);
    const outstanding = employeeAdvances.reduce((sum, a) =>
      sum + Math.max(0, number_(a.Amount) - number_(a['Amount Paid'])), 0
    );

    return Object.assign({}, e, {
      ShowsThisMonth:showsThisMonth,
      ShowPayThisMonth:showPayThisMonth,
      Salary:number_(e.Salary),
      Loans:outstanding,
      AdvanceLoanOutstanding:outstanding
    });
  });

  return {ok:true, employees:result};
}

function saveEmployee(token, payload) {
  const ss = getSpreadsheet_();
  ensureSheetIfMissing_(ss, 'Employees', ['ID','EmployeeID','Employee Name','Role','Employee Type','Contact Number','Salary','Status','Notes','Created Date','Updated Date']);

  const auth = requireAuth_(token);
  if (!auth.ok) return auth;
  payload = payload || {};

  const name = String(payload.name || '').trim();
  const role = String(payload.role || '').trim();
  const type = String(payload.type || 'Regular').trim() === 'On Call' ? 'On Call' : 'Regular';
  const contact = String(payload.contact || '').trim();
  const salary = Math.max(0, number_(payload.salary));
  const status = String(payload.status || 'Active').trim() === 'Inactive' ? 'Inactive' : 'Active';
  const notes = String(payload.notes || '').trim();

  if (!name) return {ok:false, message:'Employee name is required.'};
  if (!role) return {ok:false, message:'Employee role is required.'};

  const id = payload.id ? String(payload.id).trim() : makeId_('EMP');
  const existing = payload.id ? findRowById_('Employees', id) : null;
  const now = new Date();

  const duplicate = getOrEmpty_('Employees').find(e =>
    String(e['Employee Name'] || '').trim().toLowerCase() === name.toLowerCase() &&
    String(e.Role || '').trim().toLowerCase() === role.toLowerCase() &&
    String(e.ID || '') !== id
  );
  if (duplicate) return {ok:false, message:'An employee with the same name and role already exists.'};

  const record = {
    ID:id,
    EmployeeID:id,
    'Employee Name':name,
    Role:role,
    'Employee Type':type,
    'Contact Number':contact,
    Salary:salary,
    Status:status,
    Notes:notes,
    'Created Date':existing ? existing['Created Date'] : now,
    'Updated Date':now
  };

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    upsertRow_('Employees', record);
    audit_(auth.user, existing ? 'UPDATE' : 'CREATE', 'Employees', id, record);
    return sanitizeForClient_({ok:true, message:existing ? 'Employee updated successfully.' : 'Employee added successfully.', employee:record});
  } finally {
    lock.releaseLock();
  }
}

function deleteEmployee(token, employeeId) {
  const auth = requireAuth_(token);
  if (!auth.ok) return auth;
  const id = String(employeeId || '').trim();
  if (!id) return {ok:false, message:'Employee ID is required.'};

  const employee = findRowById_('Employees', id);
  if (!employee) return {ok:false, message:'Employee not found.'};

  const crew = getOrEmpty_('BookingCrew').filter(c => String(c.EmployeeID || '') === id);
  if (crew.length) {
    return {ok:false, message:'Cannot delete this employee because they are assigned to one or more shows. Set the employee to Inactive instead.'};
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    deleteRowById_('Employees', id);
    audit_(auth.user, 'DELETE', 'Employees', id, employee);
    return {ok:true, message:'Employee deleted successfully.'};
  } finally {
    lock.releaseLock();
  }
}

function assignEmployeeToShow(token, payload) {
  const auth = requireAuth_(token);
  if (!auth.ok) return auth;

  payload = payload || {};

  const bookingId = String(payload.bookingId || '').trim();
  const name = String(payload.employeeName || payload.name || '').trim();
  const role = String(payload.role || '').trim();
  const type = String(payload.employeeType || 'Regular').trim() === 'On Call'
    ? 'On Call'
    : 'Regular';
  const contact = String(payload.contact || '').trim();
  const salaryInput = Math.max(0, number_(payload.salary));
  const notes = String(payload.notes || '').trim();
  const assignment = String(payload.showAssignment || payload.assignment || '').trim();

  if (!bookingId) return {ok:false, message:'Booking / show is required.'};
  if (!name) return {ok:false, message:'Employee name is required.'};
  if (!role) return {ok:false, message:'Employee role is required.'};

  const booking = findRowById_('Bookings', bookingId);
  if (!booking) return {ok:false, message:'The selected client/show could not be found.'};

  const showTotal = Math.max(0, number_(
    booking['Final Amount'] ??
    booking['Total Amount Due'] ??
    booking['Total Amount'] ??
    payload.bookingTotal ??
    0
  ));

  const percentageMode = showTotal > 60000;
  const payRate = 5;
  const pay = percentageMode
    ? showTotal * payRate / 100
    : salaryInput;

  const ss = getSpreadsheet_();
  ensureSheetIfMissing_(ss, 'Employees', [
    'ID','EmployeeID','Employee Name','Role','Employee Type',
    'Contact Number','Salary','Status','Notes','Created Date','Updated Date'
  ]);
  ensureSheetIfMissing_(ss, 'BookingCrew', [
    'ID','CrewID','BookingID','EmployeeID','Employee Name','Role',
    'Show Assignment','Pay','Notes','Created Date','Updated Date'
  ]);

  try {
    /*
     * Reuse an existing employee with the same name + role.
     * This prevents duplicate master employee records when the same
     * person is assigned to another show.
     */
    let employee = getOrEmpty_('Employees').find(function(e){
      return String(e['Employee Name'] || '').trim().toLowerCase() === name.toLowerCase() &&
             String(e.Role || '').trim().toLowerCase() === role.toLowerCase();
    }) || null;

    if (!employee) {
      const employeeId = makeId_('EMP');
      employee = {
        ID: employeeId,
        EmployeeID: employeeId,
        'Employee Name': name,
        Role: role,
        'Employee Type': type,
        'Contact Number': contact,
        Salary: salaryInput,
        Status: 'Active',
        Notes: notes,
        'Created Date': new Date(),
        'Updated Date': new Date()
      };

      upsertRow_('Employees', employee);
    } else {
      // Keep the master employee current without overwriting an existing
      // salary with zero when the Crew dialog is used for a new show.
      const employeeUpdate = {
        ID: employee.ID,
        EmployeeID: employee.EmployeeID || employee.ID,
        'Employee Name': name,
        Role: role,
        'Employee Type': employee['Employee Type'] || type,
        'Contact Number': contact || employee['Contact Number'] || '',
        Salary: salaryInput > 0 ? salaryInput : number_(employee.Salary),
        Status: employee.Status || 'Active',
        Notes: notes || employee.Notes || '',
        'Created Date': employee['Created Date'] || new Date(),
        'Updated Date': new Date()
      };

      upsertRow_('Employees', employeeUpdate);
      employee = Object.assign({}, employee, employeeUpdate);
    }

    const employeeId = String(employee.EmployeeID || employee.ID);
    const bookingKey = String(booking.ID || booking.BookingID || bookingId);

    // Never assign the same employee twice to the same show.
    const duplicate = getOrEmpty_('BookingCrew').find(function(c){
      return String(c.BookingID || '').trim() === bookingKey &&
             String(c.EmployeeID || '').trim() === employeeId;
    });

    if (duplicate) {
      return {
        ok:false,
        message: name + ' is already assigned to this show.'
      };
    }

    const crewId = makeId_('CREW');

    const crew = {
      ID: crewId,
      CrewID: crewId,
      BookingID: bookingKey,
      EmployeeID: employeeId,
      'Employee Name': name,
      Role: role,
      'Show Assignment': assignment || booking['Event Name'] || booking['Event Type'] || 'Show',
      Pay: pay,
      Notes: notes,
      'Created Date': new Date(),
      'Updated Date': new Date()
    };

    // Persist the show assignment in BookingCrew. This is the record
    // that makes the employee reappear after leaving/reopening the show.
    upsertRow_('BookingCrew', crew);

    // Audit failure must never undo a successful assignment.
    try {
      audit_(auth.user, 'ASSIGN', 'Crew Assignment', crewId, {
        BookingID: bookingKey,
        EmployeeID: employeeId,
        EmployeeName: name,
        Pay: pay,
        PayMode: percentageMode ? 'Percentage' : 'Manual'
      });
    } catch (auditError) {
      console.warn('Crew assignment audit warning: ' + auditError.message);
    }

    // IMPORTANT: google.script.run cannot return Date objects to the browser.
    // The write above has already succeeded; return only JSON-safe scalar values
    // so the client receives a real success response instead of a serialization
    // failure after the row has been saved.
    return {
      ok:true,
      message: percentageMode
        ? 'Employee assigned. Show pay is automatically 5% of the booking amount.'
        : 'Employee assigned. Manual show pay was recorded.',
      employee:{
        ID:String(employee.ID || ''),
        EmployeeID:String(employee.EmployeeID || employee.ID || ''),
        Name:String(employee.Name || employee['Employee Name'] || name),
        Position:String(employee.Position || role),
        'Employee Type':String(employee['Employee Type'] || employee.EmployeeType || 'Regular'),
        Status:String(employee.Status || 'Active')
      },
      crew:{
        ID:String(crew.ID || ''),
        CrewID:String(crew.CrewID || crew.ID || ''),
        BookingID:String(crew.BookingID || bookingKey),
        EmployeeID:String(crew.EmployeeID || employeeId),
        'Employee Name':String(crew['Employee Name'] || name),
        Role:String(crew.Role || role),
        'Show Assignment':String(crew['Show Assignment'] || assignment || booking['Event Name'] || booking['Event Type'] || 'Show'),
        Pay:Number(pay || 0),
        Notes:String(crew.Notes || notes || '')
      },
      showTotal:Number(showTotal || 0),
      pay:Number(pay || 0)
    };

  } catch (error) {
    return {
      ok:false,
      message:'Unable to save the crew assignment: ' + error.message
    };
  }
}

// Backward-compatible alias for any older frontend call.
function saveBookingCrew(token, payload) {
  return assignEmployeeToShow(token, payload);
}

function deleteBookingCrew(token, crewId) {
  const auth = requireAuth_(token);
  if (!auth.ok) return auth;
  const id = String(crewId || '').trim();
  if (!id) return {ok:false, message:'Crew assignment ID is required.'};

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const crew = findRowById_('BookingCrew', id);
    if (!crew) return {ok:false, message:'Crew assignment not found.'};
    deleteRowById_('BookingCrew', id);
    audit_(auth.user, 'DELETE', 'Crew Assignment', id, crew);
    return {ok:true, message:'Employee removed from the show.'};
  } finally {
    lock.releaseLock();
  }
}

function getEmployeeAdvances(token, employeeId) {
  const auth = requireAuth_(token);
  if (!auth.ok) return auth;
  ensureSheetIfMissing_(getSpreadsheet_(), 'EmployeeAdvances', ['ID','AdvanceID','EmployeeID','Employee Name','Type','Amount','Amount Paid','Outstanding Balance','Date','Purpose','Notes','Status','Created Date','Updated Date']);
  const id = String(employeeId || '').trim();
  let rows = getOrEmpty_('EmployeeAdvances');
  if (id) rows = rows.filter(r => String(r.EmployeeID || '') === id);
  rows = rows.map(r => Object.assign({}, r, {
    Amount:number_(r.Amount),
    'Amount Paid':number_(r['Amount Paid']),
    'Outstanding Balance':Math.max(0, number_(r.Amount) - number_(r['Amount Paid']))
  }));
  return {ok:true, advances:rows};
}

function saveEmployeeAdvance(token, payload) {
  const auth = requireAuth_(token);
  if (!auth.ok) return auth;
  ensureSheetIfMissing_(getSpreadsheet_(), 'EmployeeAdvances', ['ID','AdvanceID','EmployeeID','Employee Name','Type','Amount','Amount Paid','Outstanding Balance','Date','Purpose','Notes','Status','Created Date','Updated Date']);
  payload = payload || {};

  const employeeId = String(payload.employeeId || '').trim();
  const employee = findRowById_('Employees', employeeId);
  if (!employee) return {ok:false, message:'Employee not found.'};

  const type = String(payload.type || 'Salary Advance').trim() === 'Employee Loan' ? 'Employee Loan' : 'Salary Advance';
  const amount = number_(payload.amount);
  if (amount <= 0) return {ok:false, message:'Amount must be greater than zero.'};

  const now = new Date();
  const id = payload.id ? String(payload.id) : makeId_('EADV');
  const existing = payload.id ? findRowById_('EmployeeAdvances', id) : null;
  const paid = existing ? number_(existing['Amount Paid']) : 0;
  const record = {
    ID:id,
    AdvanceID:id,
    EmployeeID:employeeId,
    'Employee Name':String(employee['Employee Name'] || ''),
    Type:type,
    Amount:amount,
    'Amount Paid':Math.min(paid, amount),
    'Outstanding Balance':Math.max(0, amount - Math.min(paid, amount)),
    Date:payload.date ? new Date(payload.date) : now,
    Purpose:String(payload.purpose || ''),
    Notes:String(payload.notes || ''),
    Status:(amount - Math.min(paid, amount)) <= 0 ? 'Fully Paid' : 'Active',
    'Created Date':existing ? existing['Created Date'] : now,
    'Updated Date':now
  };

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    upsertRow_('EmployeeAdvances', record);
    audit_(auth.user, existing ? 'UPDATE' : 'CREATE', 'Employee Advances', id, record);
    return sanitizeForClient_({ok:true, message:existing ? 'Employee advance/loan updated.' : 'Employee advance/loan recorded.', advance:record});
  } finally {
    lock.releaseLock();
  }
}

function recordEmployeeAdvancePayment(token, payload) {
  const auth = requireAuth_(token);
  if (!auth.ok) return auth;
  payload = payload || {};
  const id = String(payload.advanceId || '').trim();
  const payment = number_(payload.amount);
  if (!id) return {ok:false, message:'Advance / loan is required.'};
  if (payment <= 0) return {ok:false, message:'Payment amount must be greater than zero.'};

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const row = findRowById_('EmployeeAdvances', id);
    if (!row) return {ok:false, message:'Advance / loan not found.'};
    const amount = number_(row.Amount);
    const paid = number_(row['Amount Paid']);
    const outstanding = Math.max(0, amount - paid);
    if (payment > outstanding) return {ok:false, message:'Payment cannot exceed the outstanding balance.'};
    row['Amount Paid'] = paid + payment;
    row['Outstanding Balance'] = Math.max(0, amount - row['Amount Paid']);
    row.Status = row['Outstanding Balance'] <= 0 ? 'Fully Paid' : 'Active';
    row['Updated Date'] = new Date();
    upsertRow_('EmployeeAdvances', row);
    audit_(auth.user, 'EMPLOYEE_ADVANCE_PAYMENT', 'Employee Advances', id, {amount:payment, remaining:row['Outstanding Balance']});
    return sanitizeForClient_({ok:true, message:'Employee advance/loan payment recorded.', advance:row});
  } finally {
    lock.releaseLock();
  }
}

function getPayrollSummary(token, payrollPeriod) {
  const auth = requireAuth_(token);
  if (!auth.ok) return auth;

  const period = String(
    payrollPeriod ||
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM')
  ).trim();

  if (!/^\d{4}-\d{2}$/.test(period)) {
    return {ok:false, message:'Invalid payroll period.'};
  }

  const [yearText, monthText] = period.split('-');
  const year = Number(yearText);
  const month = Number(monthText) - 1;

  // Read each source table once. Do not create sheets, acquire locks,
  // write audit records, or perform nested sheet reads while loading Payroll.
  const employees = getOrEmpty_('Employees');
  const crews = getOrEmpty_('BookingCrew');
  const bookings = getOrEmpty_('Bookings');
  const advances = getOrEmpty_('EmployeeAdvances');
  const payrollRows = getOrEmpty_('Payroll');

  const bookingMap = new Map();
  bookings.forEach(b => {
    const id = String(b.ID || b.BookingID || '').trim();
    if (id) bookingMap.set(id, b);
  });

  const crewByEmployee = new Map();
  crews.forEach(c => {
    const employeeId = String(c.EmployeeID || '').trim();
    if (!employeeId) return;
    if (!crewByEmployee.has(employeeId)) crewByEmployee.set(employeeId, []);
    crewByEmployee.get(employeeId).push(c);
  });

  const advancesByEmployee = new Map();
  advances.forEach(a => {
    const employeeId = String(a.EmployeeID || '').trim();
    if (!employeeId) return;
    const amount = number_(a.Amount);
    const paid = number_(a['Amount Paid']);
    const outstanding = Math.max(0, amount - paid);
    if (outstanding <= 0) return;
    if (!advancesByEmployee.has(employeeId)) advancesByEmployee.set(employeeId, []);
    advancesByEmployee.get(employeeId).push({
      id:String(a.ID || a.AdvanceID || ''),
      type:String(a.Type || 'Salary Advance'),
      outstanding:outstanding,
      date:toDate_(a.Date) || toDate_(a['Created Date']) || new Date(0)
    });
  });

  const processedByEmployee = new Map();
  payrollRows.forEach(p => {
    const employeeId = String(p.EmployeeID || '').trim();
    const pPeriod = String(p['Payroll Period'] || '').trim();
    if (employeeId && pPeriod === period) processedByEmployee.set(employeeId, p);
  });

  const rows = [];

  employees.forEach(employee => {
    const employeeId = String(employee.ID || employee.EmployeeID || '').trim();
    if (!employeeId) return;

    const employeeCrews = crewByEmployee.get(employeeId) || [];
    const showBreakdown = [];

    employeeCrews.forEach(c => {
      const booking = bookingMap.get(String(c.BookingID || '').trim());
      if (!booking) return;

      const status = String(booking['Booking Status'] || 'Booked');
      if (status === 'Cancelled') return;

      const eventDate = toDate_(booking['Event Date']);
      if (!eventDate) return;
      if (eventDate.getFullYear() !== year || eventDate.getMonth() !== month) return;

      const showTotal = number_(
        booking['Final Amount'] ??
        booking['Total Amount Due'] ??
        booking['Total Amount'] ??
        booking.Subtotal ??
        0
      );

      const manualPay = number_(c.Pay ?? c['Show Pay'] ?? 0);
      const configuredRate = number_(c['Pay Rate'] ?? 5) || 5;
      const pay = showTotal > 60000
        ? showTotal * configuredRate / 100
        : manualPay;

      showBreakdown.push({
        bookingId:String(c.BookingID || ''),
        eventName:String(booking['Event Name'] || booking['Event Type'] || 'Show'),
        eventDate:eventDate,
        bookingAmount:showTotal,
        pay:pay,
        payMode:showTotal > 60000 ? '5% Automatic' : 'Manual'
      });
    });

    const gross = showBreakdown.reduce((sum, x) => sum + number_(x.pay), 0);
    const employeeAdvances = (advancesByEmployee.get(employeeId) || [])
      .slice()
      .sort((a,b) => a.date - b.date);
    const outstanding = employeeAdvances.reduce((sum, a) => sum + a.outstanding, 0);

    const existingPayroll = processedByEmployee.get(employeeId) || null;
    const deduction = existingPayroll
      ? number_(existingPayroll.Deductions)
      : Math.min(gross, outstanding);
    const net = Math.max(0, gross - deduction);

    if (showBreakdown.length || outstanding > 0) {
      rows.push({
        EmployeeID:employeeId,
        'Employee Name':String(employee['Employee Name'] || employee.Name || ''),
        Role:String(employee.Role || employee.Position || ''),
        'Employee Type':String(employee['Employee Type'] || 'Regular'),
        Shows:showBreakdown.length,
        'Show Breakdown':showBreakdown,
        'Gross Pay':gross,
        'Advance / Loan Outstanding':outstanding,
        'Suggested Deduction':deduction,
        'Net Pay':net,
        Status:existingPayroll
          ? String(existingPayroll.Status || 'Processed')
          : (showBreakdown.length ? 'Ready' : 'No Shows'),
        Processed:!!existingPayroll,
        PayrollID:existingPayroll
          ? String(existingPayroll.PayrollID || existingPayroll.ID || '')
          : ''
      });
    }
  });

  return sanitizeForClient_({
    ok:true,
    period:period,
    rows:rows,
    totals:{
      gross:rows.reduce((s,r) => s + number_(r['Gross Pay']), 0),
      deductions:rows.reduce((s,r) => s + number_(r['Suggested Deduction']), 0),
      net:rows.reduce((s,r) => s + number_(r['Net Pay']), 0),
      outstanding:rows.reduce((s,r) => s + number_(r['Advance / Loan Outstanding']), 0),
      employees:rows.length,
      processed:rows.filter(r => r.Processed).length
    }
  });
}

function savePayroll(token, payload) {
  const auth = requireAuth_(token);
  if (!auth.ok) return auth;

  payload = payload || {};

  const employeeId = String(payload.employeeId || payload.EmployeeID || '').trim();
  const period = String(payload.payrollPeriod || payload['Payroll Period'] || '').trim();
  const paymentDateValue = payload.paymentDate || payload['Payment Date'];

  if (!employeeId) return {ok:false, message:'Employee is required.'};
  if (!/^\d{4}-\d{2}$/.test(period)) return {ok:false, message:'Payroll period is required.'};

  const employee = findRowById_('Employees', employeeId);
  if (!employee) return {ok:false, message:'Employee not found.'};

  ensureSheetIfMissing_(getSpreadsheet_(), 'Payroll', [
    'ID','PayrollID','EmployeeID','Employee Name','Employee Type',
    'Compensation Type','Payroll Period','Contract Amount','Contract Rate',
    'Basic Salary','Contract Pay','Overtime','Allowances','Gas Allowance',
    'Food Allowance','Accommodation Allowance','Deductions','Net Salary',
    'Payment Date','Status','Created Date','Updated Date'
  ]);

  const existing = getOrEmpty_('Payroll').find(p =>
    String(p.EmployeeID || '') === employeeId &&
    String(p['Payroll Period'] || '') === period
  );
  if (existing) {
    return {
      ok:false,
      message:'Payroll has already been processed for this employee for ' + period + '.',
      payroll:existing
    };
  }

  const summary = getPayrollSummary(token, period);
  if (!summary.ok) return summary;

  const row = summary.rows.find(r => String(r.EmployeeID) === employeeId);
  if (!row) return {ok:false, message:'No payrollable earnings found for this employee in ' + period + '.'};

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    // Re-check duplicate after acquiring the lock.
    const latestPayroll = getOrEmpty_('Payroll').find(p =>
      String(p.EmployeeID || '') === employeeId &&
      String(p['Payroll Period'] || '') === period
    );
    if (latestPayroll) {
      return {ok:false, message:'Payroll has already been processed for this employee for ' + period + '.'};
    }

    const now = new Date();
    const payrollId = makeId_('PAY');
    const paymentDate = paymentDateValue ? new Date(paymentDateValue) : now;
    const deduction = Math.max(0, Math.min(row['Gross Pay'], number_(payload.deductions || row['Suggested Deduction'])));

    const payroll = {
      ID:payrollId,
      PayrollID:payrollId,
      EmployeeID:employeeId,
      'Employee Name':String(employee['Employee Name'] || ''),
      'Employee Type':String(employee['Employee Type'] || 'Regular'),
      'Compensation Type':'Show Based',
      'Payroll Period':period,
      'Contract Amount':row['Gross Pay'],
      'Contract Rate':5,
      'Basic Salary':row['Gross Pay'],
      'Contract Pay':row['Gross Pay'],
      Overtime:number_(payload.overtime),
      Allowances:number_(payload.allowances),
      'Gas Allowance':number_(payload.gasAllowance),
      'Food Allowance':number_(payload.foodAllowance),
      'Accommodation Allowance':number_(payload.accommodationAllowance),
      Deductions:deduction,
      'Net Salary':Math.max(0, row['Gross Pay'] - deduction),
      'Payment Date':paymentDate,
      Status:'Processed',
      'Created Date':now,
      'Updated Date':now
    };

    appendRow_('Payroll', payroll);

    // Apply payroll deduction to outstanding salary advances / employee loans,
    // oldest first, so the employee balance decreases automatically.
    let remainingDeduction = deduction;
    if (remainingDeduction > 0) {
      const advanceRows = getOrEmpty_('EmployeeAdvances')
        .filter(a => String(a.EmployeeID || '') === employeeId)
        .sort((a,b) => {
          const da = toDate_(a.Date) || toDate_(a['Created Date']) || new Date(0);
          const db = toDate_(b.Date) || toDate_(b['Created Date']) || new Date(0);
          return da - db;
        });

      for (const advance of advanceRows) {
        if (remainingDeduction <= 0) break;
        const amount = number_(advance.Amount);
        const paid = number_(advance['Amount Paid']);
        const outstanding = Math.max(0, amount - paid);
        if (outstanding <= 0) continue;

        const applied = Math.min(remainingDeduction, outstanding);
        advance['Amount Paid'] = paid + applied;
        advance['Outstanding Balance'] = Math.max(0, amount - advance['Amount Paid']);
        advance.Status = advance['Outstanding Balance'] <= 0 ? 'Fully Paid' : 'Active';
        advance['Updated Date'] = now;
        upsertRow_('EmployeeAdvances', advance);
        remainingDeduction -= applied;
      }
    }

    audit_(auth.user, 'PROCESS', 'Payroll', payrollId, {
      employeeId:employeeId,
      payrollPeriod:period,
      grossPay:row['Gross Pay'],
      deduction:deduction,
      netPay:payroll['Net Salary']
    });

    return sanitizeForClient_({
      ok:true,
      message:'Payroll processed successfully.',
      payroll:payroll
    });
  } finally {
    lock.releaseLock();
  }
}


function getMaterials(token) {
  const auth = requireAuth_(token);
  if (!auth.ok) return auth;
  return {ok:true, materials:getOrEmpty_('Materials')};
}

function saveMaterial(token, payload) {
  const auth = requireAuth_(token);
  if (!auth.ok) return auth;
  payload = payload || {};

  const name = String(payload.name || '').trim();
  const category = String(payload.category || 'Materials').trim();
  const unit = String(payload.unit || '').trim();
  const purchaseCost = number_(payload.purchaseCost);
  const minimumStock = Math.max(0, Math.floor(number_(payload.minimumStock)));

  if (!name) return {ok:false, message:'Material name is required.'};
  if (!unit) return {ok:false, message:'Unit is required.'};
  if (purchaseCost < 0) return {ok:false, message:'Purchase cost cannot be negative.'};

  const id = payload.id ? String(payload.id) : makeId_('MAT');
  const existing = payload.id ? findRowById_('Materials', id) : null;
  const now = new Date();

  const record = {
    ID:id,
    MaterialID:id,
    Name:name,
    Category:category,
    Unit:unit,
    'Purchase Cost':purchaseCost,
    'Minimum Stock':minimumStock,
    Status:'Active',
    Description:String(payload.description || ''),
    'Created Date':existing ? existing['Created Date'] : now,
    'Updated Date':now
  };

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    upsertRow_('Materials', record);
    audit_(auth.user, existing ? 'UPDATE' : 'CREATE', 'Materials', id, record);
    return {ok:true, message:existing ? 'Material updated successfully.' : 'Material created successfully.', material:record};
  } finally {
    lock.releaseLock();
  }
}



function saveMaterialAndSeedStock(token, payload) {
  const auth = requireAuth_(token);
  if (!auth.ok) return auth;
  payload = payload || {};

  const name = String(payload.name || '').trim();
  const unit = String(payload.unit || 'pcs').trim();
  const purchaseCost = number_(payload.purchaseCost);
  const quantity = number_(payload.quantity);

  if (!name) return {ok:false, message:'Material name is required.'};
  if (quantity <= 0) return {ok:false, message:'Initial stock must be greater than zero.'};
  if (purchaseCost < 0) return {ok:false, message:'Purchase cost cannot be negative.'};

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const now = new Date();
    const materialId = makeId_('MAT');

    const material = {
      ID:materialId, MaterialID:materialId, Name:name,
      Category:'Material', Unit:unit, 'Purchase Cost':purchaseCost,
      'Minimum Stock':Math.max(0, Math.floor(number_(payload.minimumStock))),
      Status:'Active', Description:'',
      'Created Date':now, 'Updated Date':now
    };
    upsertRow_('Materials', material);

    // Inventory is unified by ItemID; ProductID is also populated for legacy compatibility.
    const invId = makeId_('INV');
    appendRow_('Inventory', {
      ID:invId, InventoryID:invId, ProductID:materialId, ItemID:materialId,
      'Product Name':name, 'Current Stock':quantity, 'Reserved Stock':0,
      'Available Stock':quantity, 'Minimum Stock':material['Minimum Stock'],
      'Purchase Cost':purchaseCost, 'Unit Cost':purchaseCost,
      'Selling Price':0, 'Inventory Value':quantity * purchaseCost,
      'Updated Date':now
    });

    audit_(auth.user, 'CREATE', 'Materials', materialId, {quantity, purchaseCost});
    return sanitizeForClient_({ok:true, message:'Material created with initial stock.', material:material});
  } finally {
    lock.releaseLock();
  }
}


/* =========================
   SHOW PRODUCTS / MATERIALS USAGE
   ========================= */

function getShowUsageData(token, bookingId) {
  const auth = requireAuth_(token);
  if (!auth.ok) return auth;

  const booking = findRowById_('Bookings', String(bookingId || ''));
  if (!booking) return {ok:false, message:'Booking / show not found.'};

  // Read-only workspace load: do not rebuild/sync the entire inventory here.
  // Inventory is synchronized after stock/usage write operations and when the Inventory module loads.

  const products = getOrEmpty_('Products').filter(p => String(p.Status || 'Active') !== 'Inactive');
  const materials = getOrEmpty_('Materials').filter(m => String(m.Status || 'Active') !== 'Inactive');
  const inventory = getOrEmpty_('Inventory');

  const allItems = []
    .concat(products.map(p => ({
      ID:String(p.ID || p.ProductID || ''),
      Name:String(p.Name || ''),
      Type:'Product',
      Category:String(p.Category || ''),
      Unit:String(p.Unit || ''),
      PurchaseCost:number_(p['Purchase Cost']),
      UnitCost:number_(p['Unit Cost'] ?? p['Purchase Cost']),
      SellingCost:number_(p['Selling Cost']),
      SellingPrice:number_(p['Selling Price'])
    })))
    .concat(materials.map(m => ({
      ID:String(m.ID || m.MaterialID || ''),
      Name:String(m.Name || ''),
      Type:'Material',
      Category:String(m.Category || 'Material'),
      Unit:String(m.Unit || ''),
      PurchaseCost:number_(m['Purchase Cost']),
      SellingPrice:0
    })));

  allItems.forEach(item => {
    const inv = inventory.find(i =>
      String(i.ProductID || '') === item.ID ||
      String(i.ItemID || '') === item.ID
    );
    item.CurrentStock = inv ? number_(inv['Current Stock']) : 0;
    item.ReservedStock = inv ? number_(inv['Reserved Stock']) : 0;
    item.AvailableStock = inv ? number_(inv['Available Stock']) : 0;
    item.MinimumStock = inv ? number_(inv['Minimum Stock']) : 0;
    item.StockStatus = stockStatus_(item.AvailableStock, item.MinimumStock);
  });

  const usage = getOrEmpty_('BookingUsage')
    .filter(u => String(u.BookingID || '') === String(bookingId));

  return {
    ok:true,
    booking:booking,
    items:allItems,
    usage:usage,
    totals:{
      productCost:usage.filter(u => String(u['Usage Type']) === 'Product')
        .reduce((s,u)=>s+number_(u['Total Cost']),0),
      materialCost:usage.filter(u => String(u['Usage Type']) === 'Material')
        .reduce((s,u)=>s+number_(u['Total Cost']),0),
      totalCost:usage.reduce((s,u)=>s+number_(u['Total Cost']),0)
    }
  };
}


function getShowProductData(token, bookingId, skipRecalculate) {
  const auth = requireAuth_(token);
  if (!auth.ok) return auth;

  const id = String(bookingId || '').trim();
  if (!id) return {ok:false, message:'Booking / show ID is required.'};
  const booking = findRowById_('Bookings', id);
  if (!booking) return {ok:false, message:'Booking / show not found.'};

  if (!skipRecalculate) {
    recalculateBookingAmountFromProductUsage_(id);
    const refreshedBooking = findRowById_('Bookings', id);
    if (refreshedBooking) Object.assign(booking, refreshedBooking);
  }

  const products = getOrEmpty_('Products').filter(p => String(p.Status || 'Active') !== 'Inactive');
  const inventory = getOrEmpty_('Inventory');
  const usageAll = getOrEmpty_('BookingUsage');
  const usage = decorateUsageStockFields_(usageAll.filter(u => String(u.BookingID || '') === id && String(u['Usage Type'] || '') === 'Product')).map(function(u){
    const product = products.find(function(p){
      return String(p.ID || p.ProductID || '').trim() === String(u.ItemID || '').trim();
    });
    const unitCost = product
      ? number_(product['Unit Cost'] ?? product['Purchase Cost'])
      : number_(u['Unit Cost'] ?? u['Purchase Cost']);
    const sellingPrice = product
      ? number_(product['Selling Price'])
      : number_(u['Selling Price']);
    const row = Object.assign({}, u);
    row['Unit Cost'] = unitCost;
    row['Total Cost'] = unitCost * number_(row.Quantity);
    row['Selling Price'] = sellingPrice;
    row['Total Selling Price'] = sellingPrice * number_(row.Quantity);
    return row;
  });
  const materialCost = usageAll.filter(u => String(u.BookingID || '') === id && String(u['Usage Type'] || '') === 'Material')
    .reduce((s,u)=>s+number_(u['Total Cost']),0);

  const items = products.map(function(p){
    const item = {
      ID:String(p.ID || p.ProductID || ''),
      Name:String(p.Name || ''),
      Type:'Product',
      Category:String(p.Category || ''),
      Unit:String(p.Unit || ''),
      PurchaseCost:number_(p['Purchase Cost']),
      UnitCost:number_(p['Unit Cost'] ?? p['Purchase Cost']),
      SellingCost:number_(p['Selling Cost']),
      SellingPrice:number_(p['Selling Price'])
    };
    const inv = inventory.find(i => String(i.ProductID || '') === item.ID || String(i.ItemID || '') === item.ID);
    item.CurrentStock = inv ? number_(inv['Current Stock']) : 0;
    item.ReservedStock = inv ? number_(inv['Reserved Stock']) : 0;
    item.AvailableStock = inv ? number_(inv['Available Stock']) : 0;
    item.MinimumStock = inv ? number_(inv['Minimum Stock']) : 0;
    item.StockStatus = stockStatus_(item.AvailableStock, item.MinimumStock);
    return item;
  });

  const totals = calculateShowUsageTotals_(usage);

  return {
    ok:true,
    booking:booking,
    items:items,
    usage:usage,
    totals:totals
  };
}

function getShowMaterialData(token, bookingId) {
  const auth = requireAuth_(token);
  if (!auth.ok) return auth;

  const id = String(bookingId || '').trim();
  if (!id) return {ok:false, message:'Booking / show ID is required.'};
  const booking = findRowById_('Bookings', id);
  if (!booking) return {ok:false, message:'Booking / show not found.'};

  const materials = getOrEmpty_('Materials').filter(m => String(m.Status || 'Active') !== 'Inactive');
  const inventory = getOrEmpty_('Inventory');
  const usageAll = getOrEmpty_('BookingUsage');
  const usage = usageAll.filter(u => String(u.BookingID || '') === id && String(u['Usage Type'] || '') === 'Material');
  const productCost = usageAll.filter(u => String(u.BookingID || '') === id && String(u['Usage Type'] || '') === 'Product')
    .reduce((s,u)=>s+number_(u['Total Cost']),0);

  const items = materials.map(function(m){
    const item = {
      ID:String(m.ID || m.MaterialID || ''),
      Name:String(m.Name || ''),
      Type:'Material',
      Category:String(m.Category || 'Material'),
      Unit:String(m.Unit || ''),
      PurchaseCost:number_(m['Purchase Cost']),
      SellingPrice:0
    };
    const inv = inventory.find(i => String(i.ProductID || '') === item.ID || String(i.ItemID || '') === item.ID);
    item.CurrentStock = inv ? number_(inv['Current Stock']) : 0;
    item.ReservedStock = inv ? number_(inv['Reserved Stock']) : 0;
    item.AvailableStock = inv ? number_(inv['Available Stock']) : 0;
    item.MinimumStock = inv ? number_(inv['Minimum Stock']) : 0;
    item.StockStatus = stockStatus_(item.AvailableStock, item.MinimumStock);
    return item;
  });

  return {
    ok:true,
    booking:booking,
    items:items,
    usage:usage,
    totals:{
      productCost:productCost,
      materialCost:usage.reduce((s,u)=>s+number_(u['Total Cost']),0),
      totalCost:productCost+usage.reduce((s,u)=>s+number_(u['Total Cost']),0)
    }
  };
}



function isMaterialProductCategory_(product) {
  if (!product) return false;
  return String(product.Category || '').trim().toLowerCase() === 'materials';
}

function calculateShowUsageTotals_(usageRows) {
  const usage = Array.isArray(usageRows) ? usageRows : [];
  const products = getOrEmpty_('Products');

  let productCost = 0;
  let materialCost = 0;
  let productSellingTotal = 0;

  usage.forEach(function(u){
    const type = String(u['Usage Type'] || '').trim();
    const itemId = String(u.ItemID || '').trim();
    const qty = number_(u.Quantity);

    if (type === 'Material') {
      materialCost += number_(u['Total Cost']);
      return;
    }

    if (type !== 'Product') return;

    const product = products.find(function(p){
      return String(p.ID || p.ProductID || '').trim() === itemId;
    });
    const isMaterial = isMaterialProductCategory_(product);

    const unitCost = product
      ? number_(product['Unit Cost'] ?? product['Purchase Cost'])
      : number_(u['Unit Cost'] ?? u['Purchase Cost']);
    const usageCost = unitCost * qty;

    if (isMaterial) {
      materialCost += usageCost;
      const sellingPrice = product
        ? number_(product['Selling Price'])
        : number_(u['Selling Price']);
      productSellingTotal += sellingPrice * qty;
      return;
    }

    productCost += usageCost;
    const sellingPrice = product
      ? number_(product['Selling Price'])
      : number_(u['Selling Price']);
    productSellingTotal += sellingPrice * qty;
  });

  return {
    productCost: productCost,
    materialCost: materialCost,
    productSellingTotal: productSellingTotal,
    totalCost: productCost + materialCost
  };
}

function recalculateBookingAmountFromProductUsage_(bookingId) {
  const id = String(bookingId || '').trim();
  if (!id) return null;

  const booking = findRowById_('Bookings', id);
  if (!booking) return null;

  const usages = getOrEmpty_('BookingUsage').filter(function(u){
    return String(u.BookingID || '').trim() === id && String(u['Usage Type'] || '') === 'Product';
  });
  const products = getOrEmpty_('Products');

  const productTotalAmount = usages.reduce(function(sum, u){
    const product = products.find(function(p){
      return String(p.ID || p.ProductID || '').trim() === String(u.ItemID || '').trim();
    });

    const sellingPrice = product ? number_(product['Selling Price']) : number_(u['Selling Price']);
    return sum + (sellingPrice * number_(u.Quantity));
  }, 0);

  // Product Total Amount becomes the booking subtotal / booking amount.
  const discount = Math.max(0, number_(booking.Discount ?? booking['Discount Amount']));
  const taxableBase = Math.max(0, productTotalAmount - discount);
  const tax = taxableBase * 0.05;
  const finalAmount = taxableBase + tax;
  const paid = getOrEmpty_('Payments')
    .filter(function(p){ return String(p.BookingID || '').trim() === id; })
    .reduce(function(sum, p){ return sum + number_(p.Amount); }, 0) ||
    Math.max(0, number_(booking['Paid Amount'] ?? booking['Down Payment'] ?? booking.DownPayment));
  const balance = Math.max(0, finalAmount - paid);

  booking.Subtotal = productTotalAmount;
  booking.Discount = discount;
  booking['Tax Rate'] = 0.05;
  booking.Tax = tax;
  booking['Final Amount'] = finalAmount;
  booking['Total Amount Due'] = finalAmount;
  booking['Paid Amount'] = paid;
  booking['Down Payment'] = paid;
  booking.Balance = balance;
  booking['Payment Status'] = balance <= 0 ? 'Paid' : paid > 0 ? 'Partial' : 'Unpaid';
  booking['Updated Date'] = new Date();

  upsertRow_('Bookings', booking);
  recalculateCrewPayForBooking_(id);

  return {
    productTotalAmount: productTotalAmount,
    tax: tax,
    finalAmount: finalAmount,
    paid: paid,
    balance: balance
  };
}

function saveShowUsage(token, payload) {
  const auth = requireAuth_(token);
  if (!auth.ok) return auth;

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    return saveShowUsageOne_(auth, payload || {}, false);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Saves several product/material usages in a single request. Adding products
 * one at a time (one google.script.run call per product) meant every extra
 * product paid for its own full network round-trip AND its own booking-total
 * recalculation. Batching them into one call means only one round-trip and
 * one final recalculation for the whole batch.
 */
function saveShowUsageBatch(token, items) {
  const auth = requireAuth_(token);
  if (!auth.ok) return auth;

  const list = Array.isArray(items) ? items : [];
  if (!list.length) return {ok:false, message:'No items to save.'};

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const saved = [];
    let bookingId = '';
    let recalcNeeded = false;

    // Every item in a batch shares the same show, and almost always the same
    // Products/Materials catalog and Inventory snapshot. Fetching those once
    // and sharing them across all items (instead of re-reading the whole
    // Bookings/Products/Materials/Inventory sheets for every single product)
    // is what actually made multi-product saves slow even after batching the
    // network round-trip into one call.
    const ctx = {
      bookingId: '',
      booking: null,
      productsRows: getOrEmpty_('Products'),
      materialsRows: getOrEmpty_('Materials'),
      inventoryRows: getOrEmpty_('Inventory'),
      usageRows: getOrEmpty_('BookingUsage'),
      pendingUsageRows: [],
      pendingStockOutRows: [],
      auditEntries: []
    };

    for (let i = 0; i < list.length; i++) {
      const payload = list[i] || {};
      bookingId = String(payload.bookingId || bookingId);

      const thisBookingId = String(payload.bookingId || '').trim();
      if (thisBookingId && thisBookingId !== ctx.bookingId) {
        ctx.booking = findRowById_('Bookings', thisBookingId);
        ctx.bookingId = thisBookingId;
      }

      const result = saveShowUsageOne_(auth, payload, true, ctx);

      if (!result.ok) {
        // Earlier items in this batch already built their BookingUsage /
        // StockOut rows - flush those before reporting the failure so a
        // partial batch isn't silently lost.
        appendRows_('BookingUsage', ctx.pendingUsageRows);
        appendRows_('StockOut', ctx.pendingStockOutRows);
        if (ctx.auditEntries.length) {
          audit_(auth.user, 'SHOW_USAGE_BATCH', 'Client Workspace', bookingId, {
            bookingId:bookingId,
            items:ctx.auditEntries
          });
        }
        if (saved.length && bookingId) recalculateBookingAmountFromProductUsage_(bookingId);
        return sanitizeForClient_({
          ok:false,
          message: result.message ||
            ('Unable to save ' + (payload.itemId || 'an item') + '.'),
          savedCount: saved.length
        });
      }

      saved.push(result);
      if (String(payload.usageType || '') === 'Product') recalcNeeded = true;
    }

    // Write every BookingUsage / StockOut row for the whole batch in one
    // setValues() call each, instead of one appendRow_() round trip per item.
    appendRows_('BookingUsage', ctx.pendingUsageRows);
    appendRows_('StockOut', ctx.pendingStockOutRows);

    // One consolidated audit/activity log entry for the whole batch instead
    // of one pair of log rows per item.
    if (ctx.auditEntries.length) {
      audit_(auth.user, 'SHOW_USAGE_BATCH', 'Client Workspace', bookingId, {
        bookingId:bookingId,
        items:ctx.auditEntries
      });
    }

    const bookingTotals = recalcNeeded && bookingId
      ? recalculateBookingAmountFromProductUsage_(bookingId)
      : null;

    return sanitizeForClient_({
      ok:true,
      message: saved.length + ' item' + (saved.length === 1 ? '' : 's') + ' saved successfully.',
      results: saved,
      bookingTotals: bookingTotals
    });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Core logic shared by saveShowUsage() and saveShowUsageBatch(). Does not
 * acquire its own lock - callers must already hold the script lock.
 * skipRecalc lets a batch caller defer the booking-total recalculation
 * (which rescans BookingUsage/Products/Payments and crew pay) until after
 * every item in the batch has been written, instead of once per item.
 *
 * ctx (optional) lets a batch caller share the Bookings/Products/Materials/
 * Inventory data it already fetched across every item in the batch, instead
 * of every single item re-reading those same sheets from scratch. When
 * called for a single save (ctx is undefined), each lookup falls back to
 * fetching fresh, exactly as before.
 */
function saveShowUsageOne_(auth, payload, skipRecalc, ctx) {
  payload = payload || {};
  ctx = ctx || {};

  const bookingId = String(payload.bookingId || '').trim();
  const usageType = String(payload.usageType || '').trim();
  const itemId = String(payload.itemId || '').trim();
  const quantity = number_(payload.quantity);

  if (!bookingId) return {ok:false, message:'Booking / show is required.'};
  if (usageType !== 'Product' && usageType !== 'Material') {
    return {ok:false, message:'Usage type must be Product or Material.'};
  }
  if (!itemId) return {ok:false, message:'Select an item.'};
  if (quantity <= 0) return {ok:false, message:'Quantity must be greater than zero.'};

  {
    const booking = ctx.booking || findRowById_('Bookings', bookingId);
    if (!booking) return {ok:false, message:'Booking / show not found.'};

    const masterSheet = usageType === 'Product' ? 'Products' : 'Materials';
    const masterRows = usageType === 'Product'
      ? (ctx.productsRows || getOrEmpty_('Products'))
      : (ctx.materialsRows || getOrEmpty_('Materials'));
    const item = masterRows.find(function(r){
      return String(r.ID || r.ProductID || r.MaterialID || '').trim() === itemId;
    }) || findRowById_(masterSheet, itemId);
    if (!item) return {ok:false, message:usageType + ' not found.'};

    const purchaseCost = usageType === 'Product'
      ? number_(item['Unit Cost'] ?? item['Purchase Cost'])
      : number_(item['Purchase Cost']);
    const sellingCost = usageType === 'Product'
      ? number_(item['Selling Cost'])
      : 0;

    // Do NOT call syncInventoryRecords_() here - it rescans every product and
    // every Stock In / Stock Out row on every single usage save, which made
    // adding a product to a show noticeably slow. The block below already
    // ensures an inventory row exists for this specific item, and the write
    // further down keeps it up to date, so a full resync is unnecessary.

    // Inventory uses ProductID for historical compatibility, and ItemID when
    // the row came from the Materials master. Reuse the caller's cached
    // Inventory snapshot when given, so a batch of N items reads the
    // Inventory sheet once instead of N times.
    const inventoryRows = ctx.inventoryRows || getOrEmpty_('Inventory');
    let inv = inventoryRows.find(i =>
      String(i.ProductID || '') === itemId ||
      String(i.ItemID || '') === itemId
    );

    if (!inv) {
      // First use: create an inventory row for the material/product.
      const now = new Date();
      inv = {
        ID:makeId_('INV'),
        InventoryID:makeId_('INVREF'),
        ProductID:itemId,
        ItemID:itemId,
        'Product Name':String(item.Name || ''),
        'Current Stock':0,
        'Reserved Stock':0,
        'Available Stock':0,
        'Minimum Stock':usageType === 'Product'
          ? number_(item['Minimum Stock'])
          : number_(item['Minimum Stock']),
        'Purchase Cost':purchaseCost,
        'Unit Cost':purchaseCost,
        'Selling Price':usageType === 'Product' ? number_(item['Selling Price']) : 0,
        'Inventory Value':0,
        'Updated Date':now
      };
      appendRow_('Inventory', inv);
      // Keep the shared cache in sync so a later item in the same batch that
      // uses this same product sees it as already existing.
      inventoryRows.push(inv);
    }

    const available = number_(inv['Available Stock']);
    if (quantity > available) {
      return {
        ok:false,
        message:'Insufficient stock for ' + String(item.Name || '') +
          '. Available: ' + available + ' ' + String(item.Unit || '') + '.'
      };
    }

    const now = new Date();
    const sellingPriceForItem = usageType === 'Product' ? number_(item['Selling Price']) : 0;

    // If this item already has a usage line for this show, add to it instead
    // of creating a duplicate row every time the same product is added again.
    const usageRows = ctx.usageRows || getOrEmpty_('BookingUsage');
    const existingUsage = usageRows.find(function(u){
      return String(u.BookingID || '').trim() === bookingId &&
        String(u.ItemID || '').trim() === itemId &&
        String(u['Usage Type'] || '') === usageType;
    });

    // totalCost below always refers to the cost of THIS addition (used for
    // the audit trail and the return value) - it's distinct from the usage
    // row's own 'Total Cost' field, which reflects the full merged quantity.
    const totalCost = purchaseCost * quantity;
    let usageId;

    if (existingUsage) {
      usageId = String(existingUsage.ID || existingUsage.UsageID || makeId_('USE'));
      const mergedQuantity = number_(existingUsage.Quantity) + quantity;

      existingUsage.Quantity = mergedQuantity;
      existingUsage['Purchase Cost'] = purchaseCost;
      existingUsage['Unit Cost'] = purchaseCost;
      existingUsage['Selling Price'] = sellingPriceForItem;
      existingUsage['Total Selling Price'] = sellingPriceForItem * mergedQuantity;
      existingUsage['Total Cost'] = purchaseCost * mergedQuantity;
      existingUsage['Updated Date'] = now;
      // Keep the original 'Used Date' - Beginning Stock stays anchored to
      // when the item was first added, and Remaining Stock (beginning minus
      // the now-merged total quantity) still reflects the current total use.
      upsertRow_('BookingUsage', existingUsage);
    } else {
      usageId = makeId_('USE');

      const usageRow = {
        ID:usageId,
        UsageID:usageId,
        BookingID:bookingId,
        'Usage Type':usageType,
        ItemID:itemId,
        'Item Name':String(item.Name || ''),
        Quantity:quantity,
        'Purchase Cost':purchaseCost,
        'Unit Cost':purchaseCost,
        'Selling Price':sellingPriceForItem,
        'Total Selling Price':sellingPriceForItem * quantity,
        'Total Cost':totalCost,
        'Used Date':now,
        Status:'Used',
        'Created Date':now,
        'Updated Date':now
      };

      // A batch caller (saveShowUsageBatch) collects new rows and writes all
      // of them for the whole batch in one setValues() call each, instead of
      // one appendRow_() round trip per item. A single-item save (ctx has no
      // pending arrays) writes immediately, same as before.
      if (ctx.pendingUsageRows) {
        ctx.pendingUsageRows.push(usageRow);
      } else {
        appendRow_('BookingUsage', usageRow);
      }
      // Keep the shared cache in sync so a later item in the same batch that
      // reuses this same product merges into it instead of duplicating it.
      usageRows.push(usageRow);
    }

    // Always record a fresh physical stock-out transaction - it represents
    // this specific stock movement even when the usage line above was merged
    // into an existing one.
    const stockOutId = makeId_('SOUT');
    const stockOut = {
      ID:stockOutId,
      StockOutID:stockOutId,
      ItemType:usageType,
      ItemID:itemId,
      ProductID:itemId,
      'Product Name':String(item.Name || ''),
      Quantity:quantity,
      'Unit Cost':purchaseCost,
      BookingID:bookingId,
      UsageType:usageType,
      Reason:'Show usage',
      'Batch Number':'',
      'Storage Location':'',
      'Transaction Date':now,
      'Created Date':now,
      'Updated Date':now
    };

    if (ctx.pendingStockOutRows) {
      ctx.pendingStockOutRows.push(stockOut);
    } else {
      appendRow_('StockOut', stockOut);
    }

    // Update the Inventory row in place. `inv` is either the row we just
    // looked up or the one just created above, so this always reflects any
    // earlier items in the same batch that already touched this product.
    inv['Current Stock'] = Math.max(0, number_(inv['Current Stock']) - quantity);
    inv['Available Stock'] = Math.max(
      0,
      number_(inv['Current Stock']) - number_(inv['Reserved Stock'])
    );
    inv['Purchase Cost'] = number_(item['Purchase Cost']);
    inv['Unit Cost'] = number_(item['Unit Cost'] ?? item['Purchase Cost']);
    inv['Inventory Value'] =
      number_(inv['Current Stock']) * number_(inv['Unit Cost']);
    inv['Updated Date'] = now;
    upsertRow_('Inventory', inv);
    const currentInv = inv;

    const bookingTotals = (!skipRecalc && usageType === 'Product')
      ? recalculateBookingAmountFromProductUsage_(bookingId)
      : null;

    // A batch caller writes one consolidated audit/activity entry for the
    // whole batch after the loop instead of one pair of log rows per item
    // (each audit_() call appends to both AuditLog and ActivityLogs).
    if (ctx.auditEntries) {
      ctx.auditEntries.push({
        usageType:usageType,
        itemId:itemId,
        itemName:String(item.Name || ''),
        quantity:quantity,
        totalCost:totalCost
      });
    } else {
      audit_(auth.user, 'SHOW_USAGE', usageType, usageId, {
        bookingId:bookingId,
        itemId:itemId,
        quantity:quantity,
        totalCost:totalCost
      });
    }

    return {
      ok:true,
      message:String(item.Name || '') + ' usage recorded.',
      stock:{
        current:currentInv ? number_(currentInv['Current Stock']) : Math.max(0, available - quantity),
        available:currentInv ? number_(currentInv['Available Stock']) : Math.max(0, available - quantity),
        minimum:currentInv ? number_(currentInv['Minimum Stock']) : number_(item['Minimum Stock']),
        status:currentInv
          ? stockStatus_(number_(currentInv['Available Stock']), number_(currentInv['Minimum Stock']))
          : stockStatus_(Math.max(0, available - quantity), number_(item['Minimum Stock']))
      },
      usage:{
        id:usageId,
        bookingId:bookingId,
        usageType:usageType,
        itemId:itemId,
        itemName:String(item.Name || ''),
        quantity:quantity,
        purchaseCost:purchaseCost,
        totalCost:totalCost
      },
      bookingTotals:bookingTotals
    };
  }
}

function removeShowUsage(token, usageId) {
  const auth = requireAuth_(token);
  if (!auth.ok) return auth;

  const id = String(usageId || '').trim();
  if (!id) return {ok:false, message:'Usage record is required.'};

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    const usage = findRowById_('BookingUsage', id);
    if (!usage) return {ok:false, message:'Usage record not found.'};

    // Reverse every physical stock-out transaction tied to this usage line.
    // A single BookingUsage row can represent several "Add Product" actions
    // merged together (see saveShowUsageOne_), each of which logged its own
    // StockOut row, so remove all of them instead of looking for one row
    // whose quantity happens to match the merged total.
    const stockOutIds = getOrEmpty_('StockOut')
      .filter(r =>
        String(r.BookingID || '') === String(usage.BookingID || '') &&
        String(r.ProductID || r.ItemID || '') === String(usage.ItemID || '') &&
        String(r.Reason || '') === 'Show usage'
      )
      .map(r => String(r.ID || ''))
      .filter(Boolean);

    const inv = getOrEmpty_('Inventory').find(i =>
      String(i.ProductID || i.ItemID || '') === String(usage.ItemID || '')
    );

    const now = new Date();
    if (inv) {
      inv['Current Stock'] = number_(inv['Current Stock']) + number_(usage.Quantity);
      inv['Available Stock'] = Math.max(
        0,
        number_(inv['Current Stock']) - number_(inv['Reserved Stock'])
      );
      inv['Inventory Value'] =
        number_(inv['Current Stock']) * number_(inv['Unit Cost'] ?? usage['Purchase Cost']);
      inv['Updated Date'] = now;
      upsertRow_('Inventory', inv);
    }

    stockOutIds.forEach(function(sid){ deleteRowById_('StockOut', sid); });
    deleteRowById_('BookingUsage', id);

    const bookingTotals = String(usage['Usage Type'] || '') === 'Product'
      ? recalculateBookingAmountFromProductUsage_(String(usage.BookingID || ''))
      : null;

    audit_(auth.user, 'REMOVE_SHOW_USAGE', usage['Usage Type'], id, usage);
    return {ok:true, message:'Usage removed and stock restored.', bookingTotals:bookingTotals};
  } finally {
    lock.releaseLock();
  }
}

function stockStatus_(available, minimum) {
  available = number_(available);
  minimum = number_(minimum);
  if (available <= 0) return 'Out of Stock';
  if (available <= minimum) return 'Low Stock';
  return 'In Stock';
}



/* =========================
   CLIENT / SHOW STATUS
   ========================= */

function updateClientShowStatus(token, bookingId, status) {
  const auth = requireAuth_(token);
  if (!auth.ok) return auth;

  const id = String(bookingId || '').trim();
  const nextStatus = String(status || '').trim();

  if (!id) return {ok:false, message:'Client/show ID is required.'};
  if (!['Pending','Booked','Cancelled'].includes(nextStatus)) {
    return {ok:false, message:'Invalid booking status.'};
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    const booking = findRowById_('Bookings', id);
    if (!booking) return {ok:false, message:'Client/show not found.'};

    const previous = String(booking['Booking Status'] || 'Pending');
    booking['Booking Status'] = nextStatus;
    booking['Updated Date'] = new Date();
    upsertRow_('Bookings', booking);
    recalculateCrewPayForBooking_(id);

    audit_(auth.user, 'STATUS_CHANGE', 'Clients', id, {
      from: previous,
      to: nextStatus
    });

    return sanitizeForClient_({
      ok:true,
      booking:booking,
      message:
        nextStatus === 'Booked' ? 'Client is now booked.' :
        nextStatus === 'Cancelled' ? 'Client/show cancelled.' :
        'Client/show moved to pending.'
    });
  } finally {
    lock.releaseLock();
  }
}

function getPendingClients(token) {
  const auth = requireAuth_(token);
  if (!auth.ok) return auth;

  const rows = getOrEmpty_('Bookings').filter(b =>
    String(b['Booking Status'] || 'Pending') === 'Pending'
  );

  return {ok:true, bookings:rows};
}


/* =========================
   CLIENT EDIT / DELETE
   ========================= */

function updateClientShow(token, payload) {
  const auth = requireAuth_(token);
  if (!auth.ok) return auth;
  payload = payload || {};

  const bookingId = String(payload.bookingId || '').trim();
  if (!bookingId) return {ok:false, message:'Client/show ID is required.'};

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    const booking = findRowById_('Bookings', bookingId);
    if (!booking) return {ok:false, message:'Client/show not found.'};

    let clientId = String(booking.ClientID || booking.ClientId || '').trim();
    let client = null;
    const clientRows = getOrEmpty_('Clients');

    // 1) Normal/current relationship: Booking.ClientID -> Clients.ID
    if (clientId) {
      client = clientRows.find(function(row){
        return String(row.ID || '').trim() === clientId;
      }) || null;
    }

    // 2) Legacy relationship: Booking.ClientID -> Clients.ClientID
    if (!client && clientId) {
      client = clientRows.find(function(row){
        return String(row.ClientID || '').trim() === clientId;
      }) || null;
    }

    // 3) Frontend may provide the exact client identifier.
    const payloadClientId = String(payload.clientId || '').trim();
    if (!client && payloadClientId) {
      client = clientRows.find(function(row){
        return String(row.ID || '').trim() === payloadClientId ||
               String(row.ClientID || '').trim() === payloadClientId;
      }) || null;
      if (client) clientId = String(client.ID || client.ClientID || payloadClientId).trim();
    }

    // 4) Legacy/older records sometimes lost the relationship ID.
    // Match the client using the booking's stored Client Name, then Phone.
    if (!client) {
      const bookingClientName = String(booking['Client Name'] || '').trim().toLowerCase();
      const bookingPhone = String(booking['Contact Information'] || '').trim();
      if (bookingClientName) {
        client = clientRows.find(function(row){
          return String(row.Name || '').trim().toLowerCase() === bookingClientName;
        }) || null;
      }
      if (!client && bookingPhone) {
        client = clientRows.find(function(row){
          return String(row.Phone || '').trim() === bookingPhone;
        }) || null;
      }
      if (client) {
        clientId = String(client.ID || client.ClientID || '').trim();
        // Repair the booking relationship so future edits always work.
        if (clientId) {
          booking.ClientID = clientId;
          booking.ClientId = clientId;
        }
      }
    }

    if (!client) return {ok:false, message:'Client record not found. The booking is not linked to a client record.'};
    if (!clientId) clientId = String(client.ID || client.ClientID || '').trim();

    const now = new Date();

    client.Name = String(payload.clientName || '').trim();
    client.Phone = String(payload.phone || '').trim();
    client.Email = String(payload.email || '').trim();
    client.Address = String(payload.address || '').trim();
    client.Notes = String(payload.notes || '').trim();
    client['Updated Date'] = now;
    upsertRow_('Clients', client);

    booking['Client Name'] = client.Name;
    booking['Contact Information'] = client.Phone;
    booking['Event Type'] = String(payload.eventType || '');
    booking['Event Name'] = String(payload.eventName || '');
    booking['Event Date'] = String(payload.eventDate || '');
    booking['Event Time'] = normalizeEventTime_(payload.eventTime);
    booking.Location = String(payload.location || '');
    booking['Booking Status'] = String(payload.bookingStatus || booking['Booking Status'] || 'Pending');
    booking.Notes = String(payload.notes || '');
    booking['Updated Date'] = now;
    upsertRow_('Bookings', booking);
    setBookingEventTimeText_(bookingId, booking['Event Time']);
    booking['Event Time'] = getPersistedBookingEventTime_(bookingId) || normalizeEventTime_(payload.eventTime || payload.event_time || '');
    recalculateCrewPayForBooking_(bookingId);

    audit_(auth.user, 'UPDATE', 'Clients', bookingId, {
      clientId:clientId,
      clientName:client.Name
    });

    // Return only HTML-service-safe primitives. Returning Date objects here can make
    // google.script.run fail AFTER the spreadsheet update has already succeeded.
    return {
      ok:true,
      message:'Client/show updated successfully.',
      bookingId:bookingId,
      clientId:clientId
    };
  } finally {
    lock.releaseLock();
  }
}

function cancelClientShow(token, bookingId) {
  const auth = requireAuth_(token);
  if (!auth.ok) return auth;

  const id = String(bookingId || '').trim();
  if (!id) return {ok:false, message:'Client/show ID is required.'};

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    const booking = findRowById_('Bookings', id);
    if (!booking) return {ok:false, message:'Client/show not found.'};

    const previous = String(booking['Booking Status'] || 'Pending');
    booking['Booking Status'] = 'Cancelled';
    booking['Updated Date'] = new Date();
    upsertRow_('Bookings', booking);

    recalculateCrewPayForBooking_(id);

    audit_(auth.user, 'CANCEL', 'Clients', id, {
      from: previous,
      to: 'Cancelled'
    });

    return sanitizeForClient_({
      ok:true,
      message:'Client/show cancelled successfully.',
      booking:booking
    });
  } finally {
    lock.releaseLock();
  }
}

function confirmPendingClient(token, bookingId) {
  const result = updateClientShowStatus(token, bookingId, 'Booked');
  if (!result || !result.ok) return result;

  // Return only serializable values to google.script.run.
  // The booking itself can contain Date objects, but the frontend only needs success/message here.
  return {
    ok: true,
    message: result.message || 'Client is now booked.'
  };
}

function deleteClientShowPermanently(token, bookingId) {
  const auth = requireAuth_(token);
  if (!auth.ok) return auth;

  const id = String(bookingId || '').trim();
  if (!id) return {ok:false, message:'Client/show ID is required.'};

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    const booking = findRowById_('Bookings', id);
    if (!booking) return {ok:false, message:'Client/show not found.'};

    const bookingIdValue = String(booking.ID || booking.BookingID || id);
    const clientId = String(booking.ClientID || '').trim();

    // Remove all booking-linked child records so a permanent delete
    // really disappears from operational views.
    [
      'BookingItems',
      'BookingUsage',
      'BookingCrew',
      'Payments',
      'PaymentSchedules',
      'BookingExpenses',
      'Expenses',
      'StockOut'
    ].forEach(function(sheetName){
      deleteRowsByField_(sheetName, 'BookingID', bookingIdValue);
    });

    // Remove linked custom calendar records.
    [
      'CalendarEvents'
    ].forEach(function(sheetName){
      deleteRowsByField_(sheetName, 'RelatedID', bookingIdValue);
    });

    deleteRowById_('Bookings', bookingIdValue);

    // Only remove the client master record when no other booking still references it.
    if (clientId) {
      const stillUsed = getOrEmpty_('Bookings').some(function(b){
        return String(b.ClientID || '').trim() === clientId;
      });
      if (!stillUsed) {
        deleteRowById_('Clients', clientId);
      }
    }

    audit_(auth.user, 'DELETE', 'Clients', bookingIdValue, {
      clientId:clientId,
      reason:'Permanent client/show deletion'
    });

    return {
      ok:true,
      message:'Client/show permanently deleted.'
    };
  } finally {
    lock.releaseLock();
  }
}

// Backward-compatible server function name used by older frontend code.
function deleteClientShow(token, bookingId) {
  return deleteClientShowPermanently(token, bookingId);
}


/* =========================
   PRODUCTS / INVENTORY API
   ========================= */

function getProducts(token) {
  const auth = requireAuth_(token);
  if (!auth.ok) return auth;
  return { ok: true, products: getOrEmpty_('Products') };
}

function saveProduct(token, payload) {
  const auth = requireAuth_(token);
  if (!auth.ok) return auth;
  payload = payload || {};

  const name = String(payload.name || '').trim();
  const source = String(payload.source || 'Local').trim() === 'International' ? 'International' : 'Local';
  const category = String(payload.category || '').trim();
  const unit = String(payload.unit || '').trim();
  const purchaseCost = number_(payload.purchaseCost ?? payload.productCost);
  const sellingCost = number_(payload.sellingCost);
  const unitCost = payload.unitCost === '' || payload.unitCost === null || payload.unitCost === undefined
    ? purchaseCost
    : number_(payload.unitCost);
  const sellingPrice = number_(payload.sellingPrice);
  const minimumStock = Math.max(0, Math.floor(number_(payload.minimumStock)));

  if (!name) return { ok:false, message:'Product name is required.' };
  if (!category) return { ok:false, message:'Product category is required.' };
  if (!unit) return { ok:false, message:'Unit is required.' };
  if (purchaseCost < 0 || sellingCost < 0 || unitCost < 0 || sellingPrice < 0) return { ok:false, message:'Product cost, selling cost, unit cost, and selling price cannot be negative.' };

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const rows = getOrEmpty_('Products');
    const duplicate = rows.find(r =>
      String(r.Name || '').trim().toLowerCase() === name.toLowerCase() &&
      String(r.Category || '').trim().toLowerCase() === category.toLowerCase() &&
      String(r.ID || '') !== String(payload.id || '')
    );
    if (duplicate) return { ok:false, message:'A product with the same name and category already exists.' };

    const now = new Date();
    const id = payload.id ? String(payload.id) : makeId_('PROD');
    const existing = payload.id ? findRowById_('Products', id) : null;

    const record = {
      ID:id, ProductID:id, Name:name, Source:source, Category:category, Unit:unit,
      'Purchase Cost':purchaseCost, 'Selling Cost':sellingCost, 'Unit Cost':unitCost, 'Selling Price':sellingPrice, 'Minimum Stock':minimumStock, Status:payload.status === 'Inactive' ? 'Inactive' : 'Active',
      Description:String(payload.description || ''),
      'Created Date':existing ? existing['Created Date'] : now, 'Updated Date':now
    };

    upsertRow_('Products', record);
    ensureInventoryRecord_(record);

    // Initial Stock is recorded as a real Stock In transaction only when
    // the product is first created. Editing a product must NOT add stock again.
    const initialStock = Math.max(0, number_(payload.initialStock));
    if (!existing && initialStock > 0) {
      const stockInId = makeId_('SIN');
      const batchNumber = 'INITIAL-' + id;
      appendRow_('StockIn', {
        ID:stockInId,
        StockInID:stockInId,
        ProductID:id,
        'Product Name':name,
        Quantity:initialStock,
        'Purchase Cost':purchaseCost,
        'Shipping Cost':0,
        'Other Charges':0,
        Source:'Initial Stock',
        'Invoice Number':'',
        'Batch Number':batchNumber,
        'Storage Location':'Warehouse',
        'Transaction Date':now,
        'Created Date':now,
        'Updated Date':now
      });

      appendRow_('InventoryBatch', {
        ID:makeId_('BATCH'),
        BatchID:batchNumber,
        ProductID:id,
        'Product Name':name,
        'Batch Number':batchNumber,
        'Purchase Date':now,
        Source:'Initial Stock',
        Quantity:initialStock,
        'Remaining Quantity':initialStock,
        'Unit Cost':unitCost,
        'Landed Cost':purchaseCost,
        'Storage Location':'Warehouse',
        'Created Date':now,
        'Updated Date':now
      });
    }

    // Recalculate Current/Available Stock from Stock In - Stock Out.
    syncInventoryRecords_();

    audit_(auth.user, existing ? 'UPDATE' : 'CREATE', 'Products', id, {
      product:record,
      initialStock:existing ? 0 : initialStock,
      unitCost:unitCost
    });

    // Return only browser-safe values. The record contains Date objects, and returning
    // those raw Date objects through google.script.run can make the client receive
    // the failure handler even though the Google Sheet transaction already succeeded.
    const safeProduct = Object.assign({}, record, {
      'Created Date': record['Created Date'] instanceof Date ? record['Created Date'].toISOString() : String(record['Created Date'] || ''),
      'Updated Date': record['Updated Date'] instanceof Date ? record['Updated Date'].toISOString() : String(record['Updated Date'] || '')
    });

    return {
      ok:true,
      message:existing ? 'Product updated successfully.' : 'Product created successfully.',
      product:safeProduct,
      initialStock:existing ? 0 : initialStock
    };
  } finally {
    lock.releaseLock();
  }
}

function deleteProduct(token, productId) {
  const auth = requireAuth_(token);
  if (!auth.ok) return auth;

  const id = String(productId || '').trim();
  if (!id) return { ok:false, message:'Product ID is required.' };

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const row = findRowById_('Products', id);
    if (!row) return { ok:false, message:'Product not found.' };

    const inv = getOrEmpty_('Inventory').find(r => String(r.ProductID) === id);
    if (inv && number_(inv['Current Stock']) > 0) {
      return { ok:false, message:'Cannot deactivate a product with remaining stock. Use it after stock is cleared.' };
    }

    row.Status = 'Inactive';
    row['Updated Date'] = new Date();
    upsertRow_('Products', row);
    audit_(auth.user, 'DEACTIVATE', 'Products', id, row);
    return { ok:true, message:'Product marked inactive.' };
  } finally {
    lock.releaseLock();
  }
}

function getInventory(token) {
  const auth = requireAuth_(token);
  if (!auth.ok) return auth;
  syncInventoryRecords_();
  return { ok:true, inventory:getOrEmpty_('Inventory') };
}

function saveStockIn(token, payload) {
  const auth = requireAuth_(token);
  if (!auth.ok) return auth;
  payload = payload || {};

  const productId = String(payload.productId || '').trim();
  const quantity = number_(payload.quantity);
  const unitCost = number_(payload.purchaseCost);
  const shippingCost = number_(payload.shippingCost);
  const otherCharges = number_(payload.otherCharges);

  if (!productId) return { ok:false, message:'Select a product.' };
  if (quantity <= 0) return { ok:false, message:'Quantity must be greater than zero.' };
  if (unitCost < 0) return { ok:false, message:'Unit cost cannot be negative.' };

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const product = findRowById_('Products', productId);
    if (!product) return { ok:false, message:'Product not found.' };

    const now = new Date();
    const transactionDate = payload.transactionDate ? new Date(payload.transactionDate) : now;
    const landedUnitCost = unitCost;
    const batchNumber = String(payload.batchNumber || ('BATCH-' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss')));

    const id = makeId_('SIN');
    appendRow_('StockIn', {
      ID:id, StockInID:id, ProductID:productId, 'Product Name':product.Name,
      Quantity:quantity, 'Purchase Cost':unitCost, 'Shipping Cost':shippingCost,
      'Other Charges':otherCharges, Source:String(payload.source || ''),
      'Invoice Number':String(payload.invoiceNumber || ''), 'Batch Number':batchNumber,
      'Storage Location':String(payload.storageLocation || 'Warehouse'),
      'Transaction Date':transactionDate, 'Created Date':now, 'Updated Date':now
    });

    appendRow_('InventoryBatch', {
      ID:makeId_('BATCH'), BatchID:batchNumber, ProductID:productId, 'Product Name':product.Name,
      'Batch Number':batchNumber, 'Purchase Date':transactionDate, Source:String(payload.source || ''),
      Quantity:quantity, 'Remaining Quantity':quantity, 'Unit Cost':unitCost, 'Landed Cost':landedUnitCost,
      'Storage Location':String(payload.storageLocation || 'Warehouse'), 'Created Date':now, 'Updated Date':now
    });

    ensureInventoryRecord_(product);
    syncInventoryRecords_();
    audit_(auth.user, 'STOCK_IN', 'Inventory', productId, { quantity, landedUnitCost, batchNumber });
    return { ok:true, message:'Stock-in recorded successfully.' };
  } finally {
    lock.releaseLock();
  }
}

function saveStockOut(token, payload) {
  const auth = requireAuth_(token);
  if (!auth.ok) return auth;
  payload = payload || {};

  const productId = String(payload.productId || '').trim();
  const quantity = number_(payload.quantity);
  if (!productId) return { ok:false, message:'Select a product.' };
  if (quantity <= 0) return { ok:false, message:'Quantity must be greater than zero.' };

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    syncInventoryRecords_();
    const inv = getOrEmpty_('Inventory').find(r => String(r.ProductID) === productId);
    if (!inv) return { ok:false, message:'Inventory record not found.' };

    const available = number_(inv['Available Stock']);
    if (quantity > available) return { ok:false, message:'Insufficient available stock. Available: ' + available + '.' };

    const product = findRowById_('Products', productId);
    const now = new Date();
    const id = makeId_('SOUT');

    appendRow_('StockOut', {
      ID:id, StockOutID:id, ProductID:productId,
      'Product Name':product ? product.Name : inv['Product Name'],
      Quantity:quantity, 'Unit Cost':number_(inv['Unit Cost'] ?? inv['Purchase Cost']),
      BookingID:String(payload.bookingId || ''), Reason:String(payload.reason || ''),
      'Batch Number':String(payload.batchNumber || ''),
      'Storage Location':String(payload.storageLocation || ''),
      'Transaction Date':payload.transactionDate ? new Date(payload.transactionDate) : now,
      'Created Date':now, 'Updated Date':now
    });

    syncInventoryRecords_();
    audit_(auth.user, 'STOCK_OUT', 'Inventory', productId, { quantity, bookingId:String(payload.bookingId || '') });
    return { ok:true, message:'Stock-out recorded successfully.' };
  } finally {
    lock.releaseLock();
  }
}

function getPackages(token) {
  const auth = requireAuth_(token);
  if (!auth.ok) return auth;
  return { ok:true, packages:getOrEmpty_('Packages'), items:getOrEmpty_('PackageItems') };
}

function savePackage(token, payload) {
  const auth = requireAuth_(token);
  if (!auth.ok) return auth;
  payload = payload || {};

  const name = String(payload.name || '').trim();
  const items = Array.isArray(payload.items) ? payload.items : [];
  if (!name) return { ok:false, message:'Package name is required.' };
  if (!items.length) return { ok:false, message:'Add at least one product.' };

  const products = getOrEmpty_('Products');
  let estimatedCost = 0;
  const normalized = [];

  for (const item of items) {
    const p = products.find(x => String(x.ID) === String(item.productId));
    if (!p) return { ok:false, message:'A selected product no longer exists.' };

    const qty = number_(item.quantity);
    if (qty <= 0) return { ok:false, message:'Package quantities must be greater than zero.' };

    const unitCost = number_(p['Landed Cost']) || number_(p['Unit Cost']);
    const unitPrice = number_(item.unitPrice) || number_(p['Selling Price']);
    estimatedCost += qty * unitCost;

    normalized.push({ product:p, quantity:qty, unitCost, unitPrice, subtotal:qty * unitPrice });
  }

  const sellingPrice = number_(payload.sellingPrice);
  const profit = sellingPrice - estimatedCost;
  const margin = sellingPrice > 0 ? (profit / sellingPrice) * 100 : 0;
  const now = new Date();
  const id = payload.id ? String(payload.id) : makeId_('PACK');
  const existing = payload.id ? findRowById_('Packages', id) : null;

  const packageRecord = {
    ID:id, PackageID:id, 'Package Name':name, Description:String(payload.description || ''),
    'Selling Price':sellingPrice, 'Estimated Cost':estimatedCost, Profit:profit,
    'Profit Margin':margin, Status:payload.status === 'Inactive' ? 'Inactive' : 'Active',
    'Created Date':existing ? existing['Created Date'] : now, 'Updated Date':now
  };

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    upsertRow_('Packages', packageRecord);
    if (payload.id) deleteRowsByField_('PackageItems', 'PackageID', id);

    normalized.forEach(x => {
      const itemId = makeId_('PITEM');
      appendRow_('PackageItems', {
        ID:itemId, PackageItemID:itemId, PackageID:id, ProductID:x.product.ID,
        'Product Name':x.product.Name, Quantity:x.quantity, 'Unit Cost':x.unitCost,
        'Unit Price':x.unitPrice, Subtotal:x.subtotal, 'Created Date':now, 'Updated Date':now
      });
    });

    audit_(auth.user, existing ? 'UPDATE' : 'CREATE', 'Packages', id, packageRecord);
    return { ok:true, message:existing ? 'Package updated successfully.' : 'Package created successfully.' };
  } finally {
    lock.releaseLock();
  }
}

function getProductCategories(token) {
  const auth = requireAuth_(token);
  if (!auth.ok) return auth;

  let rows = getOrEmpty_('ProductCategories');
  if (!rows.length) {
    rows = [
      { ID:'LOCAL-FIREWORKS', CategoryID:'LOCAL-FIREWORKS', 'Category Type':'Local', 'Category Name':'Fireworks', Status:'Active' },
      { ID:'LOCAL-FIRECRACKERS', CategoryID:'LOCAL-FIRECRACKERS', 'Category Type':'Local', 'Category Name':'Firecrackers', Status:'Active' },
      { ID:'IMPORTED-FIREWORKS', CategoryID:'IMPORTED-FIREWORKS', 'Category Type':'Imported', 'Category Name':'Fireworks', Status:'Active' },
      { ID:'IMPORTED-FIRECRACKERS', CategoryID:'IMPORTED-FIRECRACKERS', 'Category Type':'Imported', 'Category Name':'Firecrackers', Status:'Active' }
    ];
  }
  return { ok:true, categories:rows };
}

function ensureInventoryRecord_(product) {
  const existing = getOrEmpty_('Inventory').find(r => String(r.ProductID) === String(product.ID));
  if (existing) {
    existing['Product Name'] = product.Name;
    existing['Minimum Stock'] = number_(product['Minimum Stock']);
    existing['Purchase Cost'] = number_(product['Purchase Cost']);
    existing['Unit Cost'] = number_(product['Unit Cost'] ?? product['Purchase Cost']);
    existing['Selling Price'] = number_(product['Selling Price']);
    existing['Updated Date'] = new Date();
    upsertRow_('Inventory', existing);
    return;
  }

  const now = new Date();
  appendRow_('Inventory', {
    ID:makeId_('INV'), InventoryID:makeId_('INVREF'), ProductID:product.ID,
    'Product Name':product.Name, 'Current Stock':0, 'Reserved Stock':0, 'Available Stock':0,
    'Minimum Stock':number_(product['Minimum Stock']), 'Purchase Cost':number_(product['Purchase Cost']),
    'Unit Cost':number_(product['Unit Cost'] ?? product['Purchase Cost']),
    'Selling Price':number_(product['Selling Price']), 'Inventory Value':0, 'Updated Date':now
  });
}

function syncInventoryRecords_() {
  const products = getOrEmpty_('Products');
  products.forEach(ensureInventoryRecord_);

  const stockIns = getOrEmpty_('StockIn');
  const stockOuts = getOrEmpty_('StockOut');
  const inventory = getOrEmpty_('Inventory');

  inventory.forEach(inv => {
    const pid = String(inv.ProductID);
    const ins = stockIns.filter(x => String(x.ProductID) === pid).reduce((s,x)=>s+number_(x.Quantity),0);
    const outs = stockOuts.filter(x => String(x.ProductID) === pid).reduce((s,x)=>s+number_(x.Quantity),0);
    const product = products.find(p => String(p.ID) === pid);

    inv['Current Stock'] = Math.max(0, ins - outs);
    inv['Reserved Stock'] = Math.min(inv['Current Stock'], number_(inv['Reserved Stock']));
    inv['Available Stock'] = Math.max(0, inv['Current Stock'] - inv['Reserved Stock']);

    if (product) {
      inv['Product Name'] = product.Name;
      inv['Minimum Stock'] = number_(product['Minimum Stock']);
      inv['Purchase Cost'] = number_(product['Purchase Cost']);
      inv['Unit Cost'] = number_(product['Unit Cost'] ?? product['Purchase Cost']);
      inv['Selling Price'] = number_(product['Selling Price']);
    }
    inv['Inventory Value'] = inv['Current Stock'] * number_(inv['Unit Cost'] ?? inv['Purchase Cost']);
    inv['Updated Date'] = new Date();
    upsertRow_('Inventory', inv);
  });
}

function normalizeEventTime_(value) {
  if (value === null || value === undefined || value === '') return '';

  // Canonical format for Event Time is plain 12-hour text, e.g. 11:30 PM.
  // Keep the existing Event Time column; do not return HH:mm here.
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'h:mm a');
  }

  let s = String(value).trim().replace(/^'/, '').trim();
  if (!s) return '';

  // Already canonical / legacy 12-hour text. Accept seconds too.
  let m = s.match(/^(\d{1,2}):(\d{2})(?::\d{2})?\s*(AM|PM)$/i);
  if (m) {
    let h = Number(m[1]);
    const min = m[2];
    const ap = m[3].toUpperCase();
    if (h < 1 || h > 12) return '';
    return h + ':' + min + ' ' + ap;
  }

  // HTML time input / legacy 24-hour value.
  m = s.match(/^(\d{1,2}):(\d{2})(?::\d{2})?$/);
  if (m) {
    const h24 = Number(m[1]);
    const min = Number(m[2]);
    if (h24 < 0 || h24 > 23 || min < 0 || min > 59) return '';
    const ap = h24 >= 12 ? 'PM' : 'AM';
    const h12 = h24 % 12 || 12;
    return h12 + ':' + String(min).padStart(2, '0') + ' ' + ap;
  }

  // Google Sheets may expose a time-only cell as a fraction of a day.
  if (/^0(?:\.\d+)?$/.test(s) || /^1(?:\.0+)?$/.test(s)) {
    const fraction = Number(s);
    if (isFinite(fraction) && fraction >= 0 && fraction < 1) {
      const totalMinutes = Math.round(fraction * 24 * 60) % (24 * 60);
      const h24 = Math.floor(totalMinutes / 60);
      const min = totalMinutes % 60;
      const ap = h24 >= 12 ? 'PM' : 'AM';
      const h12 = h24 % 12 || 12;
      return h12 + ':' + String(min).padStart(2, '0') + ' ' + ap;
    }
  }

  // Legacy Sheets Date/ISO values, including 1899-12-30 time serials.
  if (/^\d{4}-\d{2}-\d{2}T/.test(s)) {
    const d = new Date(s);
    if (!isNaN(d.getTime())) {
      return Utilities.formatDate(d, Session.getScriptTimeZone(), 'h:mm a');
    }
  }

  // Last chance for a parseable date/time value. Only return a clock time.
  const d = new Date(s);
  if (!isNaN(d.getTime())) {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'h:mm a');
  }

  return '';
}

function normalizeBookingHeaderKey_(value) {
  return String(value || '').toLowerCase().replace(/[^a-z0-9]/g, '');
}

function getPersistedBookingEventTime_(bookingId) {
  const sheet = getSpreadsheet_().getSheetByName('Bookings');
  if (!sheet || sheet.getLastRow() < 2) return '';

  const headers = getHeaders_(sheet);
  const idCol = headers.findIndex(h => normalizeBookingHeaderKey_(h) === 'id') + 1;
  if (!idCol) return '';

  // Accept all known/legacy spellings, including headers with spaces,
  // underscores, different casing, or small naming variations.
  const timeColumns = [];
  headers.forEach(function(header, index){
    const key = normalizeBookingHeaderKey_(header);
    if (key === 'eventtime' || key === 'time' || key === 'eventclocktime' || key === 'eventtimevalue') {
      timeColumns.push(index + 1);
    }
  });

  // Prefer the canonical column first if it exists.
  timeColumns.sort(function(a,b){
    const ka = normalizeBookingHeaderKey_(headers[a - 1]);
    const kb = normalizeBookingHeaderKey_(headers[b - 1]);
    return (ka === 'eventtime' ? -2 : ka === 'time' ? -1 : 0) -
           (kb === 'eventtime' ? -2 : kb === 'time' ? -1 : 0);
  });

  const rowCount = sheet.getLastRow() - 1;
  const ids = sheet.getRange(2, idCol, rowCount, 1).getDisplayValues().flat();
  let idx = ids.findIndex(v => String(v).trim() === String(bookingId).trim());

  // Legacy safety: some records may have only BookingID populated.
  if (idx < 0) {
    const bookingIdCol = headers.findIndex(h => normalizeBookingHeaderKey_(h) === 'bookingid') + 1;
    if (bookingIdCol) {
      const bookingIds = sheet.getRange(2, bookingIdCol, rowCount, 1).getDisplayValues().flat();
      idx = bookingIds.findIndex(v => String(v).trim() === String(bookingId).trim());
    }
  }

  if (idx < 0) return '';
  const rowNumber = idx + 2;

  for (const col of timeColumns) {
    const cell = sheet.getRange(rowNumber, col);
    const display = String(cell.getDisplayValue() || '').trim();
    const raw = cell.getValue();
    const normalized = normalizeEventTime_(display || raw);
    if (normalized) return normalized;
  }

  // Final fallback: inspect every cell in the booking row and only accept values
  // that actually parse as a clock time. This catches unusual legacy headers.
  const rowValues = sheet.getRange(rowNumber, 1, 1, headers.length).getValues()[0];
  const rowDisplays = sheet.getRange(rowNumber, 1, 1, headers.length).getDisplayValues()[0];
  for (let i = 0; i < headers.length; i++) {
    const headerKey = normalizeBookingHeaderKey_(headers[i]);
    if (headerKey === 'id' || headerKey === 'bookingid' || headerKey === 'eventdate') continue;
    const candidate = normalizeEventTime_(String(rowDisplays[i] || '').trim() || rowValues[i]);
    if (candidate) return candidate;
  }

  return '';
}

function getBookingEventTime(token, bookingId) {
  const auth = requireAuth_(token);
  if (!auth.ok) return auth;

  const id = String(bookingId || '').trim();
  if (!id) return {ok:false, message:'Booking / show ID is required.'};

  const display = getPersistedBookingEventTime_(id);
  if (display) setBookingEventTimeText_(id, display);

  return {ok:true, eventTime:display};
}

function setBookingEventTimeText_(bookingId, eventTime) {
  const sheet = getSpreadsheet_().getSheetByName('Bookings');
  if (!sheet || sheet.getLastRow() < 2) return;

  const headers = getHeaders_(sheet);
  const idCol = headers.findIndex(h => normalizeBookingHeaderKey_(h) === 'id') + 1;
  if (!idCol) return;

  const timeColIndex = headers.findIndex(h => normalizeBookingHeaderKey_(h) === 'eventtime');
  const bookingIdCol = headers.findIndex(h => normalizeBookingHeaderKey_(h) === 'bookingid');
  const targetCol = timeColIndex >= 0 ? timeColIndex + 1 : (headers.findIndex(h => normalizeBookingHeaderKey_(h) === 'time') + 1);
  if (!targetCol) return;

  const rowCount = sheet.getLastRow() - 1;
  const ids = sheet.getRange(2, idCol, rowCount, 1).getDisplayValues().flat();
  let idx = ids.findIndex(function(v){
    return String(v).trim() === String(bookingId).trim();
  });

  if (idx < 0 && bookingIdCol >= 0) {
    const bookingIds = sheet.getRange(2, bookingIdCol + 1, rowCount, 1).getDisplayValues().flat();
    idx = bookingIds.findIndex(function(v){
      return String(v).trim() === String(bookingId).trim();
    });
  }
  if (idx < 0) return;

  const rowNumber = idx + 2;
  const normalized = normalizeEventTime_(eventTime);
  const cell = sheet.getRange(rowNumber, targetCol);
  cell.setNumberFormat('@');
  cell.setValue(normalized ? normalized : '');
  cell.setNumberFormat('@');
}

function getBookingEventTimeDiagnostics(token, bookingId) {
  const auth = requireAuth_(token);
  if (!auth.ok) return auth;
  const id = String(bookingId || '').trim();
  if (!id) return {ok:false, message:'Booking / show ID is required.'};

  const sheet = getSpreadsheet_().getSheetByName('Bookings');
  if (!sheet || sheet.getLastRow() < 2) return {ok:true, bookingId:id, eventTime:'', candidates:[]};

  const headers = getHeaders_(sheet);
  const normalizedHeaders = headers.map(normalizeBookingHeaderKey_);
  const idCol = normalizedHeaders.indexOf('id');
  const bookingIdCol = normalizedHeaders.indexOf('bookingid');
  const rowCount = sheet.getLastRow() - 1;
  let idx = -1;

  if (idCol >= 0) {
    const vals = sheet.getRange(2, idCol + 1, rowCount, 1).getDisplayValues().flat();
    idx = vals.findIndex(v => String(v).trim() === id);
  }
  if (idx < 0 && bookingIdCol >= 0) {
    const vals = sheet.getRange(2, bookingIdCol + 1, rowCount, 1).getDisplayValues().flat();
    idx = vals.findIndex(v => String(v).trim() === id);
  }
  if (idx < 0) return {ok:true, bookingId:id, eventTime:'', candidates:[]};

  const rowNumber = idx + 2;
  const values = sheet.getRange(rowNumber, 1, 1, headers.length).getValues()[0];
  const displays = sheet.getRange(rowNumber, 1, 1, headers.length).getDisplayValues()[0];
  const candidates = [];

  headers.forEach(function(header, i){
    const value = String(displays[i] || values[i] || '').trim();
    if (value) candidates.push({header:String(header), value:value});
  });

  return {
    ok:true,
    bookingId:id,
    eventTime:getPersistedBookingEventTime_(id),
    candidates:candidates
  };
}

/**
 * google.script.run cannot transport raw Date objects (or objects containing
 * them) back to the browser - the request silently fails and the client's
 * failure handler runs instead of the success handler, even though the
 * underlying Sheet write already succeeded. This walks a response payload
 * and converts every Date into an ISO string so it is always safe to return.
 */
function sanitizeForClient_(value) {
  if (value instanceof Date) {
    return isNaN(value.getTime()) ? '' : value.toISOString();
  }
  if (Array.isArray(value)) {
    return value.map(sanitizeForClient_);
  }
  if (value && typeof value === 'object') {
    const out = {};
    Object.keys(value).forEach(function(key){
      out[key] = sanitizeForClient_(value[key]);
    });
    return out;
  }
  return value;
}

function findRowById_(sheetName, id) {
  return getOrEmpty_(sheetName).find(r => String(r.ID) === String(id)) || null;
}


function upsertRow_(sheetName, record) {
  const sheet = getSpreadsheet_().getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet not found: ' + sheetName);
  const headers = getHeaders_(sheet);
  const idCol = headers.indexOf('ID') + 1;
  if (!idCol) throw new Error('ID column missing from ' + sheetName);

  const last = sheet.getLastRow();
  if (last > 1) {
    const ids = sheet.getRange(2, idCol, last - 1, 1).getValues().flat();
    const idx = ids.findIndex(x => String(x) === String(record.ID));
    if (idx >= 0) {
      sheet.getRange(idx + 2, 1, 1, headers.length)
        .setValues([headers.map(h => record[h] !== undefined ? record[h] : '')]);
      return;
    }
  }
  sheet.appendRow(headers.map(h => record[h] !== undefined ? record[h] : ''));
}

function deleteRowsByField_(sheetName, field, value) {
  const sheet = getSpreadsheet_().getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return;
  const headers = getHeaders_(sheet);
  const col = headers.indexOf(field) + 1;
  if (!col) return;
  for (let row = sheet.getLastRow(); row >= 2; row--) {
    if (String(sheet.getRange(row, col).getValue()) === String(value)) sheet.deleteRow(row);
  }
}

function makeId_(prefix) {
  return prefix + '-' + Utilities.getUuid().replace(/-/g,'').slice(0,12).toUpperCase();
}


/* =========================
   PAYMENTS API (Booking-linked & global)
   ========================= */

function getAllPayments(token) {
  const auth = requireAuth_(token);
  if (!auth.ok) return auth;
  return {ok:true, payments:getOrEmpty_('Payments')};
}

function recalculateBookingPaymentTotals_(bookingId) {
  const id = String(bookingId || '').trim();
  if (!id) return null;

  const booking = findRowById_('Bookings', id);
  if (!booking) return null;

  const paid = getOrEmpty_('Payments')
    .filter(function(p){ return String(p.BookingID || '').trim() === id; })
    .reduce(function(sum, p){ return sum + number_(p.Amount); }, 0);

  const finalAmount = number_(booking['Final Amount'] ?? booking['Total Amount Due'] ?? booking['Total Amount'] ?? 0);
  const balance = Math.max(0, finalAmount - paid);

  booking['Paid Amount'] = paid;
  booking['Down Payment'] = paid;
  booking.DownPayment = paid;
  booking.Balance = balance;
  booking['Payment Status'] = balance <= 0 ? 'Paid' : paid > 0 ? 'Partial' : 'Unpaid';
  booking['Updated Date'] = new Date();
  upsertRow_('Bookings', booking);

  return {paid:paid, balance:balance, finalAmount:finalAmount};
}

function saveBookingPayment(token, payload) {
  const auth = requireAuth_(token);
  if (!auth.ok) return auth;
  payload = payload || {};

  const bookingId = String(payload.bookingId || '').trim();
  const amount = number_(payload.amount);

  if (!bookingId) return {ok:false, message:'Booking / show is required.'};
  if (amount <= 0) return {ok:false, message:'Payment amount must be greater than zero.'};

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const booking = findRowById_('Bookings', bookingId);
    if (!booking) return {ok:false, message:'Client/show not found.'};

    const id = payload.id ? String(payload.id) : makeId_('PAY');
    const existing = payload.id ? findRowById_('Payments', id) : null;

    // When editing, exclude the current payment's own amount before validating the balance.
    const finalAmount = number_(booking['Final Amount'] ?? booking['Total Amount Due'] ?? booking['Total Amount'] ?? 0);
    const paidExcludingThis = getOrEmpty_('Payments')
      .filter(function(p){ return String(p.BookingID || '').trim() === bookingId && String(p.ID || '') !== id; })
      .reduce(function(sum, p){ return sum + number_(p.Amount); }, 0);
    const availableBalance = Math.max(0, finalAmount - paidExcludingThis);

    if (finalAmount > 0 && amount > availableBalance + 0.01) {
      return {ok:false, message:'Payment amount cannot exceed the remaining balance of ' + availableBalance.toFixed(2) + '.'};
    }

    const now = new Date();
    const record = {
      ID:id,
      PaymentID:id,
      BookingID:bookingId,
      'Payment Date':payload.date ? new Date(payload.date) : now,
      Milestone:String(payload.milestone || 'Payment'),
      Amount:amount,
      Method:String(payload.method || ''),
      Reference:String(payload.reference || ''),
      Notes:String(payload.notes || ''),
      Status:'Completed',
      'Created Date':existing ? existing['Created Date'] : now,
      'Updated Date':now
    };

    upsertRow_('Payments', record);
    const totals = recalculateBookingPaymentTotals_(bookingId);
    audit_(auth.user, existing ? 'UPDATE' : 'CREATE', 'Payments', id, record);

    return sanitizeForClient_({
      ok:true,
      message:existing ? 'Payment updated successfully.' : 'Payment recorded successfully.',
      payment:record,
      totals:totals
    });
  } finally {
    lock.releaseLock();
  }
}

function deleteBookingPayment(token, paymentId) {
  const auth = requireAuth_(token);
  if (!auth.ok) return auth;
  const id = String(paymentId || '').trim();
  if (!id) return {ok:false, message:'Payment ID is required.'};

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const payment = findRowById_('Payments', id);
    if (!payment) return {ok:false, message:'Payment not found.'};

    const bookingId = String(payment.BookingID || '').trim();
    deleteRowById_('Payments', id);
    const totals = bookingId ? recalculateBookingPaymentTotals_(bookingId) : null;

    audit_(auth.user, 'DELETE', 'Payments', id, payment);
    return {ok:true, message:'Payment deleted successfully.', totals:totals};
  } finally {
    lock.releaseLock();
  }
}


/* =========================
   EXPENSES API (Booking-linked & global)
   ========================= */

function getAllExpenses(token) {
  const auth = requireAuth_(token);
  if (!auth.ok) return auth;
  return {ok:true, expenses:getOrEmpty_('Expenses')};
}

function saveShowExpense(token, payload) {
  const auth = requireAuth_(token);
  if (!auth.ok) return auth;
  payload = payload || {};

  const bookingId = String(payload.bookingId || '').trim();
  const amount = number_(payload.amount);

  if (!bookingId) return {ok:false, message:'Booking / show is required.'};
  if (amount <= 0) return {ok:false, message:'Expense amount must be greater than zero.'};

  const booking = findRowById_('Bookings', bookingId);
  if (!booking) return {ok:false, message:'Client/show not found.'};

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const id = payload.id ? String(payload.id) : makeId_('EXP');
    const existing = payload.id ? findRowById_('Expenses', id) : null;
    const now = new Date();

    const record = {
      ID:id,
      ExpenseID:id,
      Date:payload.date ? new Date(payload.date) : now,
      Amount:amount,
      Category:String(payload.category || 'Other'),
      SupplierID:existing ? (existing.SupplierID || '') : '',
      BookingID:bookingId,
      Description:String(payload.description || ''),
      'Approval Status':existing ? (existing['Approval Status'] || 'Approved') : 'Approved',
      'Created Date':existing ? existing['Created Date'] : now,
      'Updated Date':now
    };

    upsertRow_('Expenses', record);
    audit_(auth.user, existing ? 'UPDATE' : 'CREATE', 'Expenses', id, record);

    return sanitizeForClient_({
      ok:true,
      message:existing ? 'Expense updated successfully.' : 'Expense recorded successfully.',
      expense:record
    });
  } finally {
    lock.releaseLock();
  }
}

function deleteShowExpense(token, expenseId) {
  const auth = requireAuth_(token);
  if (!auth.ok) return auth;
  const id = String(expenseId || '').trim();
  if (!id) return {ok:false, message:'Expense ID is required.'};

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const expense = findRowById_('Expenses', id);
    if (!expense) return {ok:false, message:'Expense not found.'};
    deleteRowById_('Expenses', id);
    audit_(auth.user, 'DELETE', 'Expenses', id, expense);
    return {ok:true, message:'Expense deleted successfully.'};
  } finally {
    lock.releaseLock();
  }
}


/* =========================
   SETTINGS API
   ========================= */

function getSettings(token) {
  const auth = requireAuth_(token);
  if (!auth.ok) return auth;

  const rows = getOrEmpty_('Settings');
  const map = {};
  rows.forEach(function(r){
    const key = String(r.Key || '').trim();
    if (key) map[key] = String(r.Value || '');
  });

  return {
    ok:true,
    settings:{
      companyName:map.company_name || APP.NAME,
      companyAddress:map.company_address || '',
      companyContact:map.company_contact || '',
      version:map.version || APP.VERSION
    }
  };
}

function saveSettings(token, payload) {
  const auth = requireAuth_(token);
  if (!auth.ok) return auth;
  payload = payload || {};

  const companyName = String(payload.companyName || '').trim();
  if (!companyName) return {ok:false, message:'Company name is required.'};

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    saveSettingRaw_('company_name', companyName);
    saveSettingRaw_('company_address', String(payload.companyAddress || '').trim());
    saveSettingRaw_('company_contact', String(payload.companyContact || '').trim());

    audit_(auth.user, 'UPDATE', 'Settings', 'company_name', {companyName:companyName});
    return {ok:true, message:'Settings saved successfully.'};
  } finally {
    lock.releaseLock();
  }
}


/* =========================
   NOTIFICATIONS API
   ========================= */

function getNotifications(token) {
  const auth = requireAuth_(token);
  if (!auth.ok) return auth;

  const notifications = [];
  const today = startOfDay_(new Date());
  const in7Days = new Date(today.getTime() + 7 * 24 * 60 * 60 * 1000);

  const bookings = getOrEmpty_('Bookings').filter(b =>
    String(b['Booking Status'] || '').toLowerCase() !== 'cancelled'
  );

  bookings.forEach(function(b){
    const d = toDate_(b['Event Date']);
    if (d && d >= today && d <= in7Days) {
      notifications.push({
        type:'event',
        title:'Upcoming Event: ' + String(b['Event Name'] || b['Event Type'] || 'Show'),
        subtitle:String(b['Client Name'] || '') + ' · ' + formatDateKey_(d),
        time:d,
        bookingId:String(b.ID || b.BookingID || '')
      });
    }

    const balance = number_(b.Balance);
    if (balance > 0) {
      notifications.push({
        type:'payment',
        title:'Outstanding Balance: ' + String(b['Client Name'] || ''),
        subtitle:money_(balance) + ' remaining',
        time:d || new Date(),
        bookingId:String(b.ID || b.BookingID || '')
      });
    }
  });

  const lowStock = getOrEmpty_('Inventory').filter(i =>
    number_(i['Available Stock']) <= number_(i['Minimum Stock'])
  );
  lowStock.forEach(function(i){
    notifications.push({
      type:'stock',
      title:'Low Stock: ' + String(i['Product Name'] || ''),
      subtitle:'Available ' + number_(i['Available Stock']) + ' / Minimum ' + number_(i['Minimum Stock']),
      time:new Date(),
      bookingId:''
    });
  });

  notifications.sort(function(a,b){ return (toDate_(b.time) || 0) - (toDate_(a.time) || 0); });

  return sanitizeForClient_({ok:true, notifications:notifications.slice(0, 30)});
}


function countRows_(sheetName) {
  return readRows_(sheetName).length;
}

function appendRow_(sheetName, obj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('No active spreadsheet found.');

  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error('Sheet not found: ' + sheetName);

  // Use the actual sheet headers first. This supports every initialized ERP module,
  // not only the small APP.SHEETS foundation list.
  let headers = getHeaders_(sh);

  // Fallback for a sheet that has no header row yet.
  if (!headers.length) {
    headers = Object.keys(obj || {});
    if (!headers.length) throw new Error('No columns available for ' + sheetName + '.');
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  }

  const values = headers.map(h => obj[h] === undefined ? '' : obj[h]);
  sh.appendRow(values);

  return Object.fromEntries(headers.map((h, i) => [h, values[i]]));
}

/**
 * Appends several rows to a sheet with a single setValues() call instead of
 * calling appendRow_() once per record. Each individual appendRow_() call is
 * its own round trip to the Sheets service, so writing N records one at a
 * time costs N round trips; this costs one regardless of N.
 */
function appendRows_(sheetName, records) {
  const list = Array.isArray(records) ? records.filter(Boolean) : [];
  if (!list.length) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('No active spreadsheet found.');

  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error('Sheet not found: ' + sheetName);

  let headers = getHeaders_(sh);
  if (!headers.length) {
    headers = Object.keys(list[0] || {});
    if (!headers.length) throw new Error('No columns available for ' + sheetName + '.');
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  }

  const values = list.map(obj => headers.map(h => obj[h] === undefined ? '' : obj[h]));
  const startRow = sh.getLastRow() + 1;
  sh.getRange(startRow, 1, values.length, headers.length).setValues(values);
}

function readRows_(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) return [];

  const lastCol = sh.getLastColumn();
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];

  return values.map(row =>
    Object.fromEntries(
      headers.map((h, i) => [h, safeValue_(row[i])]).filter(([h]) => h !== '')
    )
  );
}

function updateById_(sheetName, id, data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(sheetName);
  const headers = APP.SHEETS[sheetName];

  if (!sh || sh.getLastRow() < 2) throw new Error('Record not found.');

  const values = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  const idIndex = headers.indexOf('ID');
  const rowIndex = values.findIndex(r => String(r[idIndex]) === String(id));

  if (rowIndex < 0) throw new Error('Record not found.');

  const current = values[rowIndex];

  headers.forEach((header, col) => {
    if (Object.prototype.hasOwnProperty.call(data, header)) {
      current[col] = data[header];
    }
  });

  sh.getRange(rowIndex + 2, 1, 1, headers.length)
    .setValues([current.slice(0, headers.length)]);
}

/* =========================
   DASHBOARD / DATA HELPERS
   ========================= */

function getOrEmpty_(sheetName) {
  try {
    return readRows_(sheetName);
  } catch (e) {
    return [];
  }
}

function toDate_(value) {
  if (!value) return null;
  if (value instanceof Date && !isNaN(value.getTime())) return value;

  const d = new Date(value);
  return isNaN(d.getTime()) ? null : d;
}

function startOfDay_(d) {
  const x = new Date(d);
  x.setHours(0, 0, 0, 0);
  return x;
}

function formatDateKey_(value) {
  const d = toDate_(value);
  if (!d) return '';
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function number_(value) {
  const n = parseFloat(String(value ?? '').replace(/,/g, ''));
  return isNaN(n) ? 0 : n;
}

function money_(value) {
  return '₱' + number_(value).toLocaleString('en-PH', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}

/* =========================
   PRIVATE HELPERS
   ========================= */

function ensureColumnIfMissing_(sheet, header) {
  if (!sheet || !header) return;
  const headers = getHeaders_(sheet);
  if (headers.some(h => String(h).trim().toLowerCase() === String(header).trim().toLowerCase())) return;
  sheet.getRange(1, sheet.getLastColumn() + 1).setValue(header);
}

function ensureSheetIfMissing_(ss, name, headers) {
  if (!ss.getSheetByName(name)) {
    ensureSheet_(ss, name, headers);
  }
}

function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);

  const desired = Array.isArray(headers) ? headers.filter(Boolean).map(String) : [];
  if (!desired.length) return sh;

  const lastCol = Math.max(sh.getLastColumn(), 1);
  const existing = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v || '').trim());

  // Brand-new/blank sheet.
  const hasAnyHeader = existing.some(Boolean);
  if (!hasAnyHeader) {
    sh.getRange(1, 1, 1, desired.length).setValues([desired]);
    sh.setFrozenRows(1);
    return sh;
  }

  // Preserve existing columns/data. Only append/fill missing ERP columns.
  const currentSet = new Set(existing.filter(Boolean));
  const missing = desired.filter(h => !currentSet.has(h));

  if (missing.length) {
    const startCol = existing.length + 1;
    sh.getRange(1, startCol, 1, missing.length).setValues([missing]);
  }

  sh.setFrozenRows(1);
  return sh;
}

function saveSettingRaw_(key, value) {
  const rows = readRows_('Settings');
  const row = rows.find(r => String(r.Key) === String(key));

  if (row) {
    updateById_('Settings', row.ID, {
      Value: String(value),
      'Updated Date': now_()
    });
  } else {
    appendRow_('Settings', {
      ID: uid_('SET'),
      Key: key,
      Value: String(value),
      'Created Date': now_(),
      'Updated Date': now_()
    });
  }
}

function requireAuth_(token) {
  const session = getSession_(token);
  return session
    ? { ok: true, user: session.user }
    : { ok: false, message: 'Session expired. Please log in again.' };
}

function createSession_(user, remember) {
  const token = Utilities.getUuid() + Utilities.getUuid();

  // Sessions are intentionally persistent. Do not expire them automatically.
  // The existing remember-me argument is kept for compatibility but no longer
  // controls a timeout.
  PropertiesService.getScriptProperties().setProperty(
    'session_' + token,
    JSON.stringify({ user: safeUser_(user) })
  );

  return token;
}

function getSession_(token) {
  if (!token) return null;

  const raw = PropertiesService.getScriptProperties().getProperty('session_' + token);
  if (!raw) return null;

  try {
    return JSON.parse(raw);
  } catch (e) {
    return null;
  }
}

function safeUser_(user) {
  return {
    ID: user.ID,
    UserID: user.UserID,
    'Full Name': user['Full Name'],
    Email: user.Email,
    Role: user.Role,
    Status: user.Status
  };
}


function audit_(user, action, module, recordId, details) {
  // Ensure the audit sheet exists even when setupSystem() was not re-run.
  const ss = getSpreadsheet_();
  ensureSheet_(ss, 'AuditLog', [
    'ID','User','Action','Module','Record ID','Timestamp','Details'
  ]);

  appendRow_('AuditLog', {
    ID: uid_('AUD'),
    User: user && (user.Email || user['Full Name']) || '',
    Action: String(action || ''),
    Module: String(module || ''),
    'Record ID': String(recordId || ''),
    Timestamp: new Date(),
    Details: typeof details === 'string' ? details : JSON.stringify(details || {})
  });

  // Keep the existing activity feed populated as well.
  try {
    logActivity_(user, action, module, details);
  } catch (e) {
    // Do not fail the main transaction because activity logging failed.
    console.warn('Activity log warning: ' + e.message);
  }
}

function logActivity_(user, action, module, details) {
  if (!user) return;

  appendRow_('ActivityLogs', {
    ID: uid_('ACT'),
    User: user.Email || user['Full Name'] || '',
    Action: action,
    Module: module,
    Details: String(details || ''),
    Timestamp: now_()
  });
}

function uid_(prefix) {
  return prefix + '-' +
    Utilities.getUuid().split('-')[0].toUpperCase() + '-' +
    Date.now().toString(36).toUpperCase();
}

function now_() {
  return new Date().toISOString();
}

function safeValue_(value) {
  if (value instanceof Date) return value.toISOString();
  if (value === null || value === undefined) return '';
  if (typeof value === 'object') {
    try {
      return JSON.stringify(value);
    } catch (e) {
      return String(value);
    }
  }
  return value;
}

function hashPassword_(password, salt) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(password) + String(salt)
  );

  return bytes
    .map(b => ('0' + (b & 0xff).toString(16)).slice(-2))
    .join('');
}
