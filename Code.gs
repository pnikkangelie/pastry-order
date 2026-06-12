// ── Sweet Crumb Bakery — Google Apps Script Backend ───────
var SHEET_ORDERS = 'Orders'
var SHEET_MENU   = 'Menu'

// ── GET ────────────────────────────────────────────────────
function doGet(e) {
  var action = e.parameter && e.parameter.action
  if (action === 'menu') return getMenuData()
  return getOrdersData()
}

function getOrdersData() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ORDERS)
    if (!sheet || sheet.getLastRow() <= 1) return jsonOut({ orders: [] })
    var rows = sheet.getDataRange().getValues()
    var orders = []
    for (var i = 1; i < rows.length; i++) {
      var r = rows[i]
      if (!r[0]) continue
      orders.push({
        id:        String(r[0]),
        timestamp: r[1] instanceof Date ? r[1].getTime() : (parseInt(r[1]) || 0),
        status:    String(r[2]),
        type:      String(r[3]),
        date:      sheetDateStr(r[4]),
        time:      String(r[5]),
        address:   String(r[6] || ''),
        customer:  { name: String(r[7]), phone: String(r[8]), email: String(r[9] || '') },
        items:     safeJSON(r[10], []),
        notes:         String(r[11] || ''),
        total:         parseFloat(r[12]) || 0,
        paymentMethod: String(r[13] || 'cash'),
        paymentRef:    String(r[14] || '')
      })
    }
    return jsonOut({ orders: orders })
  } catch (err) {
    return jsonOut({ orders: [], error: err.toString() })
  }
}

function getMenuData() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_MENU)
    if (!sheet || sheet.getLastRow() <= 1) return jsonOut({ items: [] })
    var rows = sheet.getDataRange().getValues()
    var items = []
    for (var i = 1; i < rows.length; i++) {
      var r = rows[i]
      if (!r[0]) continue
      items.push({
        id:    parseInt(r[0]) || i,
        cat:   String(r[1]),
        name:  String(r[2]),
        price: parseFloat(r[3]) || 0,
        desc:  String(r[4] || ''),
        img:   String(r[5] || '')
      })
    }
    return jsonOut({ items: items })
  } catch (err) {
    return jsonOut({ items: [], error: err.toString() })
  }
}

// ── POST ───────────────────────────────────────────────────
function doPost(e) {
  var lock = LockService.getScriptLock()
  lock.tryLock(10000)
  try {
    var data = JSON.parse(e.postData.contents)
    var ss   = SpreadsheetApp.getActiveSpreadsheet()

    // ── Orders ──
    if (data.action === 'addOrder') {
      var sheet = ss.getSheetByName(SHEET_ORDERS) || ss.insertSheet(SHEET_ORDERS)
      if (sheet.getLastRow() === 0) {
        sheet.appendRow(['ID','Submitted','Status','Type','Date','Time','Address','Name','Phone','Email','Items','Notes','Total','Payment','PayRef'])
        sheet.getRange(1,1,1,15).setFontWeight('bold').setBackground('#7b4a2c').setFontColor('#ffffff')
        sheet.setFrozenRows(1)
        sheet.setColumnWidth(1, 110)
        sheet.setColumnWidth(11, 300)
      }
      var o = data.order
      if (!o || !o.customer || !o.items || !o.items.length || o.total == null) {
        return jsonOut({ success: false, error: 'Missing required order fields' })
      }

      // Sanitize all string inputs
      o.customer.name  = sanitize(o.customer.name, 100)
      o.customer.phone = sanitize(o.customer.phone, 20)
      o.customer.email = sanitize(o.customer.email, 150)
      o.notes          = sanitize(o.notes, 500)
      o.address        = sanitize(o.address, 200)
      o.paymentRef     = sanitize(o.paymentRef, 100)
      o.id             = sanitize(o.id, 32)

      // Validate required fields
      if (!o.customer.name || !o.customer.phone || !o.date || !o.id) {
        return jsonOut({ success: false, error: 'Missing required fields' })
      }
      if (!/^[\d\s\+\-\(\)]{7,20}$/.test(o.customer.phone)) {
        return jsonOut({ success: false, error: 'Invalid phone number' })
      }
      if (o.customer.email && !/^[^\s@]{1,64}@[^\s@]{1,255}$/.test(o.customer.email)) {
        return jsonOut({ success: false, error: 'Invalid email address' })
      }

      // Rate limit by phone number
      if (!rateLimit(o.customer.phone)) {
        return jsonOut({ success: false, error: 'Too many orders from this number. Please wait before ordering again.' })
      }

      // Sanitize items
      o.items = (o.items || []).filter(function(i) { return i && i.name && i.qty > 0 })
      o.items = o.items.map(function(i) {
        return { name: sanitize(i.name, 100), price: Math.abs(parseFloat(i.price) || 0), qty: Math.min(Math.max(parseInt(i.qty) || 0, 1), 99) }
      })
      if (!o.items.length) return jsonOut({ success: false, error: 'No valid items in order' })

      var validStatuses = ['pending', 'awaiting_confirmation', 'ready', 'complete', 'cancelled']
      if (validStatuses.indexOf(o.status) < 0) o.status = 'pending'
      if (o.type !== 'pickup' && o.type !== 'delivery') o.type = 'pickup'
      sheet.appendRow([o.id, new Date(o.timestamp), o.status, o.type, o.date, o.time,
                       o.address||'', o.customer.name, o.customer.phone, o.customer.email||'',
                       JSON.stringify(o.items), o.notes||'', o.total,
                       o.paymentMethod||'cash', o.paymentRef||''])
      // Email notifications — owner alert + customer confirmation
      try {
        var props = PropertiesService.getScriptProperties()
        var ownerEmail = props.getProperty('OWNER_EMAIL')
        var itemLines = o.items.map(function(i) {
          return '  ' + i.qty + '× ' + i.name + ' — ₱' + (i.price * i.qty).toFixed(2)
        }).join('\n')
        if (ownerEmail) {
          MailApp.sendEmail(ownerEmail,
            '[New Order] ' + o.id + ' — ' + o.customer.name,
            'New order received!\n\n' +
            'Order ID: ' + o.id + '\n' +
            'Status: ' + o.status + '\n' +
            'Customer: ' + o.customer.name + ' | ' + o.customer.phone +
            (o.customer.email ? ' | ' + o.customer.email : '') + '\n' +
            'Type: ' + o.type + '\n' +
            'Date: ' + o.date + ' at ' + o.time +
            (o.address ? '\nAddress: ' + o.address : '') +
            '\n\nItems:\n' + itemLines + '\n' +
            'Total: ₱' + parseFloat(o.total).toFixed(2) + '\n' +
            'Payment: ' + (o.paymentMethod || 'cash') +
            (o.paymentRef ? ' — Ref: ' + o.paymentRef : '') +
            (o.notes ? '\nNotes: ' + o.notes : '')
          )
        }
        if (o.customer.email) {
          MailApp.sendEmail(o.customer.email,
            'Your order has been received — ' + o.id,
            'Hi ' + o.customer.name + ',\n\n' +
            'Thank you for your order! We\'ve received it and will confirm shortly.\n\n' +
            'Order ID: ' + o.id + '\n' +
            'Items:\n' + itemLines + '\n' +
            'Total: ₱' + parseFloat(o.total).toFixed(2) + '\n' +
            'Type: ' + o.type + ' on ' + o.date + ' at ' + o.time +
            (o.address ? '\nAddress: ' + o.address : '') +
            '\n\nWe\'ll be in touch to confirm your order.\n\nThank you!'
          )
        }
      } catch(mailErr) { /* Email is non-critical — don’t fail the order */ }
      return jsonOut({ success: true })
    }

    if (data.action === 'updateStatus') {
      var sheet = ss.getSheetByName(SHEET_ORDERS)
      if (sheet) {
        var vals = sheet.getDataRange().getValues()
        for (var i = 1; i < vals.length; i++) {
          if (String(vals[i][0]) === String(data.id)) { sheet.getRange(i+1,3).setValue(data.status); break }
        }
      }
      return jsonOut({ success: true })
    }

    if (data.action === 'deleteOrder') {
      var sheet = ss.getSheetByName(SHEET_ORDERS)
      if (sheet) {
        var vals = sheet.getDataRange().getValues()
        for (var i = 1; i < vals.length; i++) {
          if (String(vals[i][0]) === String(data.id)) { sheet.deleteRow(i+1); break }
        }
      }
      return jsonOut({ success: true })
    }

    // ── Menu ──
    if (data.action === 'addMenuItem') {
      var msheet = ss.getSheetByName(SHEET_MENU) || ss.insertSheet(SHEET_MENU)
      if (msheet.getLastRow() === 0) {
        msheet.appendRow(['ID','Category','Name','Price','Description','Image'])
        msheet.getRange(1,1,1,6).setFontWeight('bold').setBackground('#7b4a2c').setFontColor('#ffffff')
        msheet.setFrozenRows(1)
      }
      var newId = new Date().getTime()
      var it = data.item
      msheet.appendRow([newId, it.cat, it.name, it.price, it.desc||'', it.img||''])
      return jsonOut({ success: true, id: newId })
    }

    if (data.action === 'updateMenuItem') {
      var msheet = ss.getSheetByName(SHEET_MENU)
      if (msheet) {
        var vals = msheet.getDataRange().getValues()
        var it = data.item
        for (var i = 1; i < vals.length; i++) {
          if (String(vals[i][0]) === String(it.id)) {
            msheet.getRange(i+1,2,1,5).setValues([[it.cat, it.name, it.price, it.desc||'', it.img||'']])
            break
          }
        }
      }
      return jsonOut({ success: true })
    }

    if (data.action === 'deleteMenuItem') {
      var msheet = ss.getSheetByName(SHEET_MENU)
      if (msheet) {
        var vals = msheet.getDataRange().getValues()
        for (var i = 1; i < vals.length; i++) {
          if (String(vals[i][0]) === String(data.id)) { msheet.deleteRow(i+1); break }
        }
      }
      return jsonOut({ success: true })
    }

    return jsonOut({ success: false, error: 'Unknown action' })
  } catch (err) {
    return jsonOut({ success: false, error: err.toString() })
  } finally {
    lock.releaseLock()
  }
}

// ── Helpers ────────────────────────────────────────────────
function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON)
}
function sheetDateStr(val) {
  if (val instanceof Date) {
    return val.getFullYear() + '-' + String(val.getMonth()+1).padStart(2,'0') + '-' + String(val.getDate()).padStart(2,'0')
  }
  return String(val || '')
}
function safeJSON(val, fallback) {
  try { return JSON.parse(val) } catch(e) { return fallback }
}

// Strip HTML tags and trim to max length
function sanitize(val, maxLen) {
  if (val == null) return ''
  return String(val).replace(/<[^>]*>/g, '').replace(/[^\x20-\x7E -￿]/g, '').trim().slice(0, maxLen || 500)
}

// Rate limit: max 5 orders per phone number per hour
function rateLimit(phone) {
  var props = PropertiesService.getScriptProperties()
  var key   = 'rl_' + String(phone).replace(/\D/g, '').slice(-10)
  var now   = Date.now()
  try {
    var data = JSON.parse(props.getProperty(key) || '[]')
    data = data.filter(function(t) { return now - t < 3600000 })
    if (data.length >= 5) return false
    data.push(now)
    props.setProperty(key, JSON.stringify(data))
    return true
  } catch(e) { return true }
}
