/* ALTA — ROAS/CAC ყურადღების პანელი. გამოთვლების ბირთვი.
   DOM-ზე არ არის დამოკიდებული: იმავე კოდს იყენებს გვერდიც და ტესტიც. */
(function (root) {
  'use strict';

  var SALES_SHEETS = ['Smartphones', 'Laptops', 'TVs', 'Accessories'];
  var SPEND_SHEET = 'Ad_Spend';
  var PAID_CHANNELS = ['Google Ads', 'Meta Ads'];
  var ALL_CHANNELS = ['Google Ads', 'Meta Ads', 'Organic'];

  var SALES_COLS = ['date', 'order_id', 'customer_id', 'is_new_customer', 'sku', 'product_name',
    'units', 'net_revenue_gel', 'cost_of_goods_gel', 'channel', 'campaign_id'];
  var SPEND_COLS = ['date', 'campaign_id', 'sku', 'channel', 'spend_gel'];

  // campaign_id და product_name ცარიელი შეიძლება იყოს, დანარჩენი — არა
  var SALES_REQUIRED = ['date', 'order_id', 'customer_id', 'is_new_customer', 'sku',
    'units', 'net_revenue_gel', 'cost_of_goods_gel', 'channel'];
  var SPEND_REQUIRED = ['date', 'campaign_id', 'sku', 'channel', 'spend_gel'];

  var EPOCH = Date.UTC(1899, 11, 30);

  function serialToISO(n) {
    return new Date(EPOCH + n * 86400000).toISOString().slice(0, 10);
  }
  function isoToSerial(s) {
    return Math.round((Date.parse(s + 'T00:00:00Z') - EPOCH) / 86400000);
  }

  /* თარიღის უჯრის წაკითხვა. Excel-ის სერიული ნომერი ან yyyy-mm-dd ტექსტი.
     არარსებული კალენდარული დღე (მაგ. 2026-02-30) შეცდომაა, არა მომდევნო თვის რიცხვი. */
  function readDate(v) {
    if (typeof v === 'number') {
      if (!isFinite(v) || !Number.isInteger(v) || v < 1 || v > 80000) return { ok: false, why: 'range' };
      return { ok: true, serial: v };
    }
    if (v instanceof Date && !isNaN(v.getTime())) {
      return { ok: true, serial: Math.round((Date.UTC(v.getFullYear(), v.getMonth(), v.getDate()) - EPOCH) / 86400000) };
    }
    if (typeof v === 'string') {
      var s = v.trim();
      if (!s) return { ok: false, why: 'empty' };
      var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s);
      if (!m) return { ok: false, why: 'format' };
      var y = +m[1], mo = +m[2], d = +m[3];
      var dt = new Date(Date.UTC(y, mo - 1, d));
      if (dt.getUTCFullYear() !== y || dt.getUTCMonth() !== mo - 1 || dt.getUTCDate() !== d) {
        return { ok: false, why: 'nonexistent' };
      }
      return { ok: true, serial: Math.round((dt.getTime() - EPOCH) / 86400000) };
    }
    return { ok: false, why: 'empty' };
  }

  function readBool(v) {
    if (v === 1 || v === 0) return v === 1;
    if (v === true || v === false) return v;
    if (typeof v === 'string') {
      var s = v.trim().toUpperCase();
      if (s === '1' || s === 'TRUE') return true;
      if (s === '0' || s === 'FALSE') return false;
    }
    return null;
  }

  function readNum(v) {
    if (typeof v === 'number') return isFinite(v) ? v : null;
    if (typeof v === 'string' && v.trim() !== '') {
      var n = Number(v.trim());
      return isFinite(n) ? n : null;
    }
    return null;
  }

  function txt(v) {
    if (v === null || v === undefined) return '';
    return String(v).trim();
  }

  function isBlank(v) {
    return v === null || v === undefined || (typeof v === 'string' && v.trim() === '');
  }

  /* ---------- წაკითხვა და ვალიდაცია ---------- */

  /* sheets: { სახელი: [[უჯრა,...], ...] } — პირველი მწკრივი სათაურებია.
     აბრუნებს { errors: [...], data: {...} }. თუ errors ცარიელი არაა, data null-ია. */
  function load(sheets) {
    var errors = [];
    var err = function (sheet, row, what, value) {
      errors.push({ sheet: sheet, row: row, what: what, value: value === undefined ? '' : String(value) });
    };

    // --- სტრუქტურა ---
    var needed = SALES_SHEETS.concat([SPEND_SHEET]);
    for (var i = 0; i < needed.length; i++) {
      if (!sheets[needed[i]]) err(needed[i], null, 'ფურცელი აკლია ფაილს', '');
    }
    if (errors.length) return { errors: errors, data: null };

    function checkHeader(name, cols) {
      var rows = sheets[name];
      if (!rows || rows.length === 0) { err(name, null, 'ფურცელი ცარიელია', ''); return null; }
      var head = (rows[0] || []).map(txt);
      var idx = {};
      for (var c = 0; c < cols.length; c++) {
        var at = head.indexOf(cols[c]);
        if (at === -1) err(name, 1, 'სვეტი აკლია ან სახელი შეცვლილია', cols[c]);
        idx[cols[c]] = at;
      }
      if (rows.length < 2) err(name, null, 'ფურცელს მონაცემები არ აქვს, მხოლოდ სათაური', '');
      return idx;
    }

    var salesIdx = {}, k;
    for (k = 0; k < SALES_SHEETS.length; k++) salesIdx[SALES_SHEETS[k]] = checkHeader(SALES_SHEETS[k], SALES_COLS);
    var spendIdx = checkHeader(SPEND_SHEET, SPEND_COLS);
    if (errors.length) return { errors: errors, data: null };

    // --- გაყიდვების სტრიქონები ---
    var orders = [];
    var seenOrderId = {};
    for (k = 0; k < SALES_SHEETS.length; k++) {
      var sheet = SALES_SHEETS[k];
      var rows = sheets[sheet];
      var idx = salesIdx[sheet];
      for (var r = 1; r < rows.length; r++) {
        var raw = rows[r] || [];
        var excelRow = r + 1;
        var get = function (col) { var a = idx[col]; return a === -1 ? '' : raw[a]; };

        // მთლიანად ცარიელი მწკრივი გამოტოვებულია, არა შეცდომა
        var anything = false;
        for (var q = 0; q < SALES_COLS.length; q++) if (!isBlank(get(SALES_COLS[q]))) { anything = true; break; }
        if (!anything) continue;

        var bad = false;
        for (var q2 = 0; q2 < SALES_REQUIRED.length; q2++) {
          if (isBlank(get(SALES_REQUIRED[q2]))) {
            err(sheet, excelRow, 'სავალდებულო ველი ცარიელია', SALES_REQUIRED[q2]);
            bad = true;
          }
        }

        var d = readDate(get('date'));
        if (!d.ok) {
          if (d.why === 'nonexistent') err(sheet, excelRow, 'თარიღი არ არსებობს', txt(get('date')));
          else if (d.why !== 'empty') err(sheet, excelRow, 'თარიღი არასწორია', txt(get('date')));
          bad = true;
        }

        var orderId = txt(get('order_id'));
        if (orderId) {
          if (seenOrderId[orderId]) {
            err(sheet, excelRow, 'order_id მეორდება (მწკრივი ' + seenOrderId[orderId].row + ', ' + seenOrderId[orderId].sheet + ')', orderId);
            bad = true;
          } else {
            seenOrderId[orderId] = { sheet: sheet, row: excelRow };
          }
        }

        var isNew = readBool(get('is_new_customer'));
        if (isNew === null && !isBlank(get('is_new_customer'))) {
          err(sheet, excelRow, 'is_new_customer არც 1-ია და არც 0', txt(get('is_new_customer')));
          bad = true;
        }

        var units = readNum(get('units'));
        if (units === null) {
          if (!isBlank(get('units'))) { err(sheet, excelRow, 'units რიცხვი არ არის', txt(get('units'))); bad = true; }
        } else if (units <= 0 || !Number.isInteger(units)) {
          err(sheet, excelRow, 'units დადებითი მთელი უნდა იყოს', units);
          bad = true;
        }

        var rev = readNum(get('net_revenue_gel'));
        if (rev === null) {
          if (!isBlank(get('net_revenue_gel'))) { err(sheet, excelRow, 'net_revenue_gel რიცხვი არ არის', txt(get('net_revenue_gel'))); bad = true; }
        } else if (rev < 0) { err(sheet, excelRow, 'net_revenue_gel უარყოფითია', rev); bad = true; }

        var cogs = readNum(get('cost_of_goods_gel'));
        if (cogs === null) {
          if (!isBlank(get('cost_of_goods_gel'))) { err(sheet, excelRow, 'cost_of_goods_gel რიცხვი არ არის', txt(get('cost_of_goods_gel'))); bad = true; }
        } else if (cogs < 0) { err(sheet, excelRow, 'cost_of_goods_gel უარყოფითია', cogs); bad = true; }

        var channel = txt(get('channel'));
        if (channel && ALL_CHANNELS.indexOf(channel) === -1) {
          err(sheet, excelRow, 'channel დაუშვებელი მნიშვნელობაა', channel);
          bad = true;
        }

        var campaign = txt(get('campaign_id'));
        if (PAID_CHANNELS.indexOf(channel) !== -1 && !campaign) {
          err(sheet, excelRow, 'ფასიან შეკვეთას campaign_id არ აქვს', orderId);
          bad = true;
        }
        if (channel === 'Organic' && campaign) {
          err(sheet, excelRow, 'Organic შეკვეთას campaign_id აქვს', campaign);
          bad = true;
        }

        if (!bad) {
          orders.push({
            sheet: sheet, row: excelRow, category: sheet,
            date: d.serial, order_id: orderId, customer_id: txt(get('customer_id')),
            is_new: isNew === true, sku: txt(get('sku')), product_name: txt(get('product_name')),
            units: units, revenue: rev, cogs: cogs, channel: channel, campaign_id: campaign,
            paid: PAID_CHANNELS.indexOf(channel) !== -1
          });
        }
      }
    }

    // --- Ad_Spend ---
    var spend = [];
    var seenSpendKey = {};
    var srows = sheets[SPEND_SHEET];
    for (var s = 1; s < srows.length; s++) {
      var sraw = srows[s] || [];
      var sRow = s + 1;
      var sget = function (col) { var a = spendIdx[col]; return a === -1 ? '' : sraw[a]; };

      var some = false;
      for (var q3 = 0; q3 < SPEND_COLS.length; q3++) if (!isBlank(sget(SPEND_COLS[q3]))) { some = true; break; }
      if (!some) continue;

      var sbad = false;
      for (var q4 = 0; q4 < SPEND_REQUIRED.length; q4++) {
        if (isBlank(sget(SPEND_REQUIRED[q4]))) {
          err(SPEND_SHEET, sRow, 'სავალდებულო ველი ცარიელია', SPEND_REQUIRED[q4]);
          sbad = true;
        }
      }

      var sd = readDate(sget('date'));
      if (!sd.ok) {
        if (sd.why === 'nonexistent') err(SPEND_SHEET, sRow, 'თარიღი არ არსებობს', txt(sget('date')));
        else if (sd.why !== 'empty') err(SPEND_SHEET, sRow, 'თარიღი არასწორია', txt(sget('date')));
        sbad = true;
      }

      var sAmt = readNum(sget('spend_gel'));
      if (sAmt === null) {
        if (!isBlank(sget('spend_gel'))) { err(SPEND_SHEET, sRow, 'spend_gel რიცხვი არ არის', txt(sget('spend_gel'))); sbad = true; }
      } else if (sAmt < 0) {
        err(SPEND_SHEET, sRow, 'სარეკლამო ხარჯი უარყოფითია', sAmt);
        sbad = true;
      }

      var sCh = txt(sget('channel'));
      if (sCh && PAID_CHANNELS.indexOf(sCh) === -1) {
        err(SPEND_SHEET, sRow, 'Ad_Spend-ში channel მხოლოდ Google Ads ან Meta Ads შეიძლება იყოს', sCh);
        sbad = true;
      }

      var sSku = txt(sget('sku')), sCamp = txt(sget('campaign_id'));
      if (sd.ok) {
        var key = sd.serial + '|' + sSku + '|' + sCh + '|' + sCamp;
        if (seenSpendKey[key]) {
          err(SPEND_SHEET, sRow, 'იგივე ჩანაწერი მეორდება (მწკრივი ' + seenSpendKey[key] + ')', serialToISO(sd.serial) + ' ' + sCamp);
          sbad = true;
        } else {
          seenSpendKey[key] = sRow;
        }
      }

      if (!sbad) {
        spend.push({ row: sRow, date: sd.serial, campaign_id: sCamp, sku: sSku, channel: sCh, spend: sAmt });
      }
    }

    // --- კამპანიის კონფლიქტი ორივე ფურცელს შორის ---
    var camp = {};
    function noteCampaign(id, sku, channel, where, row) {
      if (!id) return;
      if (!camp[id]) camp[id] = { sku: sku, channel: channel, where: where, row: row };
      else {
        if (camp[id].sku !== sku) {
          err(where, row, 'კამპანია ორ სხვადასხვა SKU-ზეა მიბმული (' + camp[id].sku + ' და ' + sku + ')', id);
        }
        if (camp[id].channel !== channel) {
          err(where, row, 'კამპანია ორ სხვადასხვა არხზეა მიბმული (' + camp[id].channel + ' და ' + channel + ')', id);
        }
      }
    }
    for (k = 0; k < orders.length; k++) noteCampaign(orders[k].campaign_id, orders[k].sku, orders[k].channel, orders[k].sheet, orders[k].row);
    for (k = 0; k < spend.length; k++) noteCampaign(spend[k].campaign_id, spend[k].sku, spend[k].channel, SPEND_SHEET, spend[k].row);

    if (errors.length) return { errors: errors, data: null };

    if (orders.length === 0 && spend.length === 0) {
      return { errors: [{ sheet: '', row: null, what: 'ფაილში მონაცემები არ აღმოჩნდა', value: '' }], data: null };
    }

    // --- დამხმარე რუკები ---
    var skuCategory = {}, skuName = {};
    for (k = 0; k < orders.length; k++) {
      skuCategory[orders[k].sku] = orders[k].category;
      if (orders[k].product_name) skuName[orders[k].sku] = orders[k].product_name;
    }

    var allDates = [];
    for (k = 0; k < orders.length; k++) allDates.push(orders[k].date);
    for (k = 0; k < spend.length; k++) allDates.push(spend[k].date);

    return {
      errors: [],
      data: {
        orders: orders,
        spend: spend,
        skuCategory: skuCategory,
        skuName: skuName,
        categories: SALES_SHEETS.slice(),
        skus: Object.keys(skuCategory).concat(spend.map(function (x) { return x.sku; }))
          .filter(function (v, i, a) { return v && a.indexOf(v) === i; }).sort(),
        channels: ALL_CHANNELS.slice(),
        minDate: Math.min.apply(null, allDates),
        maxDate: Math.max.apply(null, allDates)
      }
    };
  }

  /* ---------- ფილტრი ---------- */

  function makeFilter(data, f) {
    var catSet = f.categories && f.categories.length ? f.categories : null;
    var skuSet = f.skus && f.skus.length ? f.skus : null;
    var chSet = f.channels && f.channels.length ? f.channels : null;
    var from = f.from, to = f.to;

    var orders = data.orders.filter(function (o) {
      if (o.date < from || o.date > to) return false;
      if (catSet && catSet.indexOf(o.category) === -1) return false;
      if (skuSet && skuSet.indexOf(o.sku) === -1) return false;
      if (chSet && chSet.indexOf(o.channel) === -1) return false;
      return true;
    });

    var spend = data.spend.filter(function (s) {
      if (s.date < from || s.date > to) return false;
      if (skuSet && skuSet.indexOf(s.sku) === -1) return false;
      if (chSet && chSet.indexOf(s.channel) === -1) return false;
      if (catSet) {
        var cat = data.skuCategory[s.sku];
        // უცნობი კატეგორიის SKU კატეგორიის ფილტრს ვერ აკმაყოფილებს
        if (!cat || catSet.indexOf(cat) === -1) return false;
      }
      return true;
    });

    return { orders: orders, spend: spend };
  }

  /* ---------- აგრეგაცია ---------- */

  function sum(arr, fn) {
    var t = 0;
    for (var i = 0; i < arr.length; i++) t += fn(arr[i]);
    return t;
  }

  function distinctNewCustomers(orders) {
    var seen = {}, n = 0;
    for (var i = 0; i < orders.length; i++) {
      var o = orders[i];
      if (o.paid && o.is_new && !seen[o.customer_id]) { seen[o.customer_id] = 1; n++; }
    }
    return n;
  }

  /* ROAS: ხარჯი 0 => null (N/A). ხარჯი > 0 => შეფარდება, მათ შორის 0.00x. */
  function roasOf(adRevenue, spendTotal) {
    if (!(spendTotal > 0)) return null;
    return adRevenue / spendTotal;
  }
  /* CAC: ახალი მომხმარებელი 0 => null (N/A). */
  function cacOf(spendTotal, newCustomers) {
    if (!(newCustomers > 0)) return null;
    return spendTotal / newCustomers;
  }

  /* მარჟა: მხოლოდ თვითღირებულება გამოიქვითება. გაყიდვა არ არის => null (N/A). */
  function marginOf(salesRevenue, cogs) {
    if (!(salesRevenue > 0)) return null;
    return (salesRevenue - cogs) / salesRevenue;
  }
  /* უწაგებო ROAS = 1 / მარჟა. მარჟა <= 0 => Infinity: ნებისმიერი რეკლამა ზარალია. */
  function breakEvenOf(margin) {
    if (margin === null) return null;
    if (margin <= 0) return Infinity;
    return 1 / margin;
  }

  function block(orders, spend) {
    var paidOrders = orders.filter(function (o) { return o.paid; });
    var spendTotal = sum(spend, function (s) { return s.spend; });
    var adRevenue = sum(paidOrders, function (o) { return o.revenue; });
    var salesRevenue = sum(orders, function (o) { return o.revenue; });
    var cogs = sum(orders, function (o) { return o.cogs; });
    var newCustomers = distinctNewCustomers(orders);
    var margin = marginOf(salesRevenue, cogs);
    return {
      orders: orders.length,
      salesRevenue: salesRevenue,
      adRevenue: adRevenue,
      organicRevenue: salesRevenue - adRevenue,
      cogs: cogs,
      spend: spendTotal,
      units: sum(orders, function (o) { return o.units; }),
      roas: roasOf(adRevenue, spendTotal),
      newCustomers: newCustomers,
      cac: cacOf(spendTotal, newCustomers),
      profit: salesRevenue - cogs - spendTotal,
      margin: margin,
      breakEven: breakEvenOf(margin)
    };
  }

  function inRange(d, a, b) { return d >= a && d <= b; }

  /* rules: { minSpend: 300 }
     b — ბოლო კვირის ბლოკი. breakEven — მთელი გაფილტრული პერიოდიდან. */
  function flagged(b, breakEven, rules, applyMinSpend) {
    if (!(b.spend > 0)) return false;
    if (applyMinSpend && b.spend < rules.minSpend) return false;
    if (breakEven === null) return false;        // შერჩევაში გაყიდვა არ აქვს
    if (breakEven === Infinity) return true;     // მარჟა <= 0
    if (b.roas === null) return false;
    return b.roas < breakEven;
  }

  function analyze(data, filters, rules) {
    var f = makeFilter(data, filters);
    var orders = f.orders, spend = f.spend;

    var overall = block(orders, spend);

    // ღუზა: გაფილტრული მონაცემების ბოლო თარიღი, გაყიდვებისა და ხარჯის მაქსიმუმიდან
    var dates = orders.map(function (o) { return o.date; }).concat(spend.map(function (s) { return s.date; }));
    var anchor = dates.length ? Math.max.apply(null, dates) : null;

    var lastWeek = null, prevWeek = null;
    if (anchor !== null) {
      lastWeek = { from: anchor - 6, to: anchor };
      prevWeek = { from: anchor - 13, to: anchor - 7 };
    }

    function windowed(win) {
      if (!win) return { orders: [], spend: [] };
      return {
        orders: orders.filter(function (o) { return inRange(o.date, win.from, win.to); }),
        spend: spend.filter(function (s) { return inRange(s.date, win.from, win.to); })
      };
    }
    var lw = windowed(lastWeek), pw = windowed(prevWeek);

    var prevWeekHasData = prevWeek !== null &&
      (data.minDate <= prevWeek.to) && (pw.orders.length > 0 || pw.spend.length > 0);

    // --- პროდუქტები: SKU-ები გაყიდვებიდან და ხარჯიდან, ერთად ---
    var skuSet = {};
    orders.forEach(function (o) { skuSet[o.sku] = 1; });
    spend.forEach(function (s) { skuSet[s.sku] = 1; });
    var skus = Object.keys(skuSet).sort();

    var products = skus.map(function (sku) {
      function forSku(src) {
        return {
          orders: src.orders.filter(function (o) { return o.sku === sku; }),
          spend: src.spend.filter(function (s) { return s.sku === sku; })
        };
      }
      var a = forSku(lw), b = forSku(pw), all = forSku({ orders: orders, spend: spend });
      var last = block(a.orders, a.spend);
      var prev = block(b.orders, b.spend);
      var period = block(all.orders, all.spend);
      var change = (last.roas !== null && prev.roas !== null) ? last.roas - prev.roas : null;
      return {
        sku: sku,
        name: data.skuName[sku] || '',
        category: data.skuCategory[sku] || '',
        last: last, prev: prev, period: period,
        margin: period.margin,
        breakEven: period.breakEven,
        change: change,
        flagged: flagged(last, period.breakEven, rules, true),
        reason: flagReason(last, period.breakEven, rules, true)
      };
    });

    // მონიშნულები ზემოთ, შემდეგ ROAS-ის ზრდადობით (N/A ბოლოში)
    products.sort(function (x, y) {
      if (x.flagged !== y.flagged) return x.flagged ? -1 : 1;
      var xr = x.last.roas === null ? Infinity : x.last.roas;
      var yr = y.last.roas === null ? Infinity : y.last.roas;
      if (xr !== yr) return xr - yr;
      return x.sku < y.sku ? -1 : 1;
    });

    // --- არხები ---
    var channels = ALL_CHANNELS.map(function (ch) {
      function forCh(src) {
        return {
          orders: src.orders.filter(function (o) { return o.channel === ch; }),
          spend: src.spend.filter(function (s) { return s.channel === ch; })
        };
      }
      var a = forCh(lw), b = forCh(pw), all = forCh({ orders: orders, spend: spend });
      var last = block(a.orders, a.spend);
      var prev = block(b.orders, b.spend);
      var period = block(all.orders, all.spend);
      var paid = PAID_CHANNELS.indexOf(ch) !== -1;
      return {
        channel: ch,
        paid: paid,
        last: last, prev: prev, period: period,
        margin: period.margin,
        breakEven: paid ? period.breakEven : null,
        change: (last.roas !== null && prev.roas !== null) ? last.roas - prev.roas : null,
        flagged: paid ? flagged(last, period.breakEven, rules, false) : false,
        reason: paid ? flagReason(last, period.breakEven, rules, false) : 'Organic-ზე ROAS და უწაგებო ROAS არ გამოიყენება'
      };
    }).filter(function (c) {
      return c.period.orders > 0 || c.period.spend > 0 || c.paid;
    });

    // --- დღიური რიგი ---
    var series = [];
    if (orders.length || spend.length) {
      var lo = Math.min.apply(null, dates), hi = anchor;
      var byDaySpend = {}, byDayAdRev = {}, byDayRev = {};
      spend.forEach(function (s) { byDaySpend[s.date] = (byDaySpend[s.date] || 0) + s.spend; });
      orders.forEach(function (o) {
        byDayRev[o.date] = (byDayRev[o.date] || 0) + o.revenue;
        if (o.paid) byDayAdRev[o.date] = (byDayAdRev[o.date] || 0) + o.revenue;
      });
      for (var d = lo; d <= hi; d++) {
        var sp = byDaySpend[d] || 0, ar = byDayAdRev[d] || 0;
        series.push({
          date: d, iso: serialToISO(d),
          spend: sp, adRevenue: ar, salesRevenue: byDayRev[d] || 0,
          roas: roasOf(ar, sp),
          inLastWeek: lastWeek ? inRange(d, lastWeek.from, lastWeek.to) : false
        });
      }
    }

    return {
      filters: filters,
      rules: rules,
      overall: overall,
      lastWeek: lastWeek,
      prevWeek: prevWeek,
      prevWeekHasData: prevWeekHasData,
      lastWeekBlock: block(lw.orders, lw.spend),
      prevWeekBlock: block(pw.orders, pw.spend),
      products: products,
      channels: channels,
      series: series,
      flaggedProducts: products.filter(function (p) { return p.flagged; }),
      flaggedChannels: channels.filter(function (c) { return c.flagged; })
    };
  }

  function flagReason(b, breakEven, rules, applyMinSpend) {
    if (!(b.spend > 0)) return 'ხარჯი არ ყოფილა';
    if (applyMinSpend && b.spend < rules.minSpend) return 'ხარჯი ' + rules.minSpend + ' GEL-ზე ნაკლებია';
    if (breakEven === null) return 'შერჩევაში გაყიდვა არ აქვს — მარჟა N/A';
    if (breakEven === Infinity) return 'მარჟა ნული ან უარყოფითია — ნებისმიერი რეკლამა ზარალია';
    if (b.roas === null) return 'ROAS არ გამოითვლება';
    if (b.roas < breakEven) return 'ROAS ' + fmtRoas(b.roas) + ', უწაგებო ' + fmtRoas(breakEven);
    return 'მომგებიანია — ROAS ' + fmtRoas(b.roas) + ', უწაგებო ' + fmtRoas(breakEven);
  }

  /* ---------- გაფრთხილებები (ჩატვირთვას არ აჩერებს) ---------- */

  function warnings(data) {
    var out = [];
    var ordersByCampaign = {};
    data.orders.forEach(function (o) { if (o.campaign_id) ordersByCampaign[o.campaign_id] = 1; });
    var noSale = {};
    data.spend.forEach(function (s) {
      if (!ordersByCampaign[s.campaign_id]) {
        noSale[s.campaign_id] = (noSale[s.campaign_id] || 0) + s.spend;
      }
    });
    Object.keys(noSale).forEach(function (c) {
      out.push('კამპანია ' + c + ' — ' + fmtGel(noSale[c]) + ' GEL დახარჯულია, შეკვეთის გარეშე. ხარჯი ჯამებში შედის.');
    });
    var zero = data.spend.filter(function (s) { return s.spend === 0; });
    zero.forEach(function (s) {
      out.push('ნულოვანი ხარჯი — ' + s.campaign_id + ', ' + serialToISO(s.date) + '. ROAS ამ ჩანაწერზე N/A-ა, არა 0.00x.');
    });
    var span = data.maxDate - data.minDate + 1;
    if (span < 14) {
      out.push('ფაილი ' + span + ' დღეს მოიცავს. 14 დღეზე ნაკლებია, ამიტომ წინა კვირის შედარება არასრულია.');
    }
    return out;
  }

  /* ---------- ფორმატირება ---------- */

  function fmtGel(n) {
    if (n === null || n === undefined || !isFinite(n)) return '—';
    var neg = n < 0;
    var s = Math.round(Math.abs(n)).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ' ');
    return (neg ? '−' : '') + s;
  }
  function fmtRoas(r) {
    if (r === null || r === undefined) return 'N/A';
    if (r === Infinity) return '∞';
    return r.toFixed(2) + 'x';
  }
  function fmtMargin(m) {
    if (m === null || m === undefined) return 'N/A';
    return (m * 100).toFixed(1) + '%';
  }
  function fmtCac(c) {
    if (c === null || c === undefined) return 'N/A';
    return c.toFixed(2);
  }
  function fmtChange(c) {
    if (c === null || c === undefined) return '—';
    return (c >= 0 ? '+' : '−') + Math.abs(c).toFixed(2) + 'x';
  }

  var API = {
    SALES_SHEETS: SALES_SHEETS,
    SPEND_SHEET: SPEND_SHEET,
    PAID_CHANNELS: PAID_CHANNELS,
    ALL_CHANNELS: ALL_CHANNELS,
    DEFAULT_RULES: { minSpend: 300 },
    load: load,
    analyze: analyze,
    warnings: warnings,
    serialToISO: serialToISO,
    isoToSerial: isoToSerial,
    fmtGel: fmtGel,
    fmtRoas: fmtRoas,
    fmtCac: fmtCac,
    fmtChange: fmtChange,
    fmtMargin: fmtMargin
  };

  if (typeof module !== 'undefined' && module.exports) module.exports = API;
  else root.AltaCore = API;
})(typeof self !== 'undefined' ? self : this);
