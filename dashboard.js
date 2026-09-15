const sourceName = document.querySelector('#source-name');
const headlineInsight = document.querySelector('#headline-insight');
const headlineDetail = document.querySelector('#headline-detail');
const totalValue = document.querySelector('#total-value');
const customerCount = document.querySelector('#customer-count');
const concentrationRisk = document.querySelector('#concentration-risk');
const lifecycleConfidence = document.querySelector('#lifecycle-confidence');
const totalNote = document.querySelector('#total-note');
const customerNote = document.querySelector('#customer-note');
const concentrationNote = document.querySelector('#concentration-note');
const lifecycleNote = document.querySelector('#lifecycle-note');
const categorySelect = document.querySelector('#category-select');
const categoryNote = document.querySelector('#category-note');
const selectedCategoryTitle = document.querySelector('#selected-category-title');
const selectedCategoryTotal = document.querySelector('#selected-category-total');
const customerRecords = document.querySelector('#customer-records');
const categoryAiTitle = document.querySelector('#category-ai-title');
const categoryAiDetail = document.querySelector('#category-ai-detail');
const categoryList = document.querySelector('#category-list');
const riskList = document.querySelector('#risk-list');
const evidenceGrid = document.querySelector('#evidence-grid');
const recordCount = document.querySelector('#record-count');

const currency = new Intl.NumberFormat('en-IN', { maximumFractionDigits: 0 });
const titleCase = (value) => String(value).replace(/[_-]+/g, ' ').replace(/\b\w/g, (letter) => letter.toUpperCase());

function getFile() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open('northstar-workspace', 1);
    request.onsuccess = () => {
      const transaction = request.result.transaction('files', 'readonly');
      const getRequest = transaction.objectStore('files').get('latest');
      getRequest.onsuccess = () => resolve(getRequest.result);
      getRequest.onerror = () => reject(getRequest.error);
    };
    request.onerror = () => reject(request.error);
  });
}

function parseNumber(value) {
  if (typeof value === 'number') return value;
  const normalized = String(value ?? '').replace(/[^\d.-]/g, '');
  return normalized ? Number(normalized) : 0;
}

function formatDate(value) {
  if (!value) return '';
  if (typeof value === 'string' && /^\d{4}-\d{2}-\d{2}/.test(value)) return value.slice(0, 10);
  const date = value instanceof Date ? value : new Date(value);
  if (value instanceof Date && !Number.isNaN(date.getTime())) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
  }
  return Number.isNaN(date.getTime()) ? String(value) : date.toISOString().slice(0, 10);
}

function findHeader(headers, terms) {
  return headers.find((header) => terms.some((term) => header.toLowerCase().includes(term))) || null;
}

function parseTextRows(text) {
  const lines = text.split(/\r?\n/).map((line) => line.trim()).filter(Boolean);
  if (lines.length < 2) return [];

  const delimiter = lines.some((line) => line.includes('\t')) ? '\t' : lines.some((line) => line.includes('|')) ? '|' : ',';
  const headers = lines[0].split(delimiter).map((header) => header.trim()).filter(Boolean);
  const headerText = headers.join(' ').toLowerCase();
  const looksLikeTable = ['customer', 'value', 'category', 'product', 'billing'].some((term) => headerText.includes(term));
  if (!looksLikeTable || headers.length < 2) return [];

  return lines.slice(1).map((line) => {
    const values = line.split(delimiter).map((value) => value.trim());
    return headers.reduce((row, header, index) => {
      row[header] = values[index] || '';
      return row;
    }, {});
  }).filter((row) => Object.values(row).some(Boolean));
}

async function readFile(file) {
  const extension = file.name.split('.').pop().toLowerCase();
  const buffer = await file.arrayBuffer();

  if (extension === 'docx') {
    const result = await window.mammoth.extractRawText({ arrayBuffer: buffer });
    return { rows: parseTextRows(result.value), text: result.value };
  }

  if (['xlsx', 'xls', 'csv'].includes(extension)) {
    const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    return { rows: XLSX.utils.sheet_to_json(sheet, { defval: '' }), text: XLSX.utils.sheet_to_csv(sheet) };
  }

  const text = await file.text();
  return { rows: parseTextRows(text), text };
}

function buildDataModel(rows, text) {
  const headers = rows.length ? Object.keys(rows[0]) : [];
  const valueHeader = findHeader(headers, ['value', 'sales', 'revenue', 'amount', 'billing']);
  const customerHeader = findHeader(headers, ['customer name', 'customer', 'account', 'company']);
  const codeHeader = findHeader(headers, ['customer code', 'code', 'id']);
  const categoryHeader = findHeader(headers, ['category', 'motion', 'type', 'segment']);
  const productHeader = findHeader(headers, ['product', 'sku', 'offering', 'item']);
  const billingHeader = findHeader(headers, ['billing date', 'invoice date', 'sales date', 'date']);
  const effectiveFromHeader = findHeader(headers, ['effective from', 'valid from', 'start date']);
  const effectiveToHeader = findHeader(headers, ['effective to', 'valid to', 'end date', 'expiry']);
  const enhancementHeader = findHeader(headers, ['enhancement', 'launch date', 'release date', 'last update', 'updated']);
  const normalizedRows = rows.map((row) => ({
    value: parseNumber(valueHeader ? row[valueHeader] : 0),
    customer: String(customerHeader ? row[customerHeader] : codeHeader ? row[codeHeader] : 'Unidentified').trim() || 'Unidentified',
    code: String(codeHeader ? row[codeHeader] : '').trim(),
    product: String(productHeader ? row[productHeader] : '').trim(),
    billingDate: formatDate(billingHeader ? row[billingHeader] : ''),
    effectiveFrom: formatDate(effectiveFromHeader ? row[effectiveFromHeader] : ''),
    effectiveTo: formatDate(effectiveToHeader ? row[effectiveToHeader] : ''),
    enhancementDate: formatDate(enhancementHeader ? row[enhancementHeader] : ''),
    category: String(categoryHeader ? row[categoryHeader] : 'Unclassified').trim() || 'Unclassified',
  }));
  const total = normalizedRows.reduce((sum, row) => sum + row.value, 0);
  const groups = Object.values(normalizedRows.reduce((result, row) => {
    result[row.category] ||= { name: row.category, value: 0, count: 0 };
    result[row.category].value += row.value;
    result[row.category].count += 1;
    return result;
  }, {})).sort((a, b) => b.value - a.value);
  const topShare = total ? groups[0]?.value / total : 0;
  const datedRows = normalizedRows.filter((row) => row.billingDate);
  const referenceDate = datedRows.length ? datedRows.map((row) => new Date(row.billingDate)).sort((a, b) => b - a)[0] : new Date();
  normalizedRows.forEach((row) => {
    const invoice = new Date(row.billingDate || referenceDate);
    const end = row.effectiveTo ? new Date(row.effectiveTo) : null;
    const enhancement = row.enhancementDate ? new Date(row.enhancementDate) : null;
    row.effectiveStatus = end && end < invoice ? 'Expired' : row.effectiveFrom && new Date(row.effectiveFrom) > invoice ? 'Not active' : 'Active';
    row.ageYears = enhancement ? Math.max(0, (invoice - enhancement) / (365.25 * 24 * 60 * 60 * 1000)) : null;
  });
  const hasLifecycleData = Boolean(productHeader && billingHeader && (effectiveFromHeader || effectiveToHeader) && enhancementHeader);
  return { headers, rows: normalizedRows, groups, total, topShare, customerHeader, productHeader, billingHeader, effectiveFromHeader, effectiveToHeader, enhancementHeader, hasLifecycleData, referenceDate, text };
}

function renderSelectedCategory(model, category) {
  const selectedRows = category === 'all' ? model.rows : model.rows.filter((row) => row.category.toUpperCase() === category.toUpperCase());
  const selectedTotal = selectedRows.reduce((sum, row) => sum + row.value, 0);
  const accountTotal = new Set(selectedRows.map((row) => row.customer)).size;
  const displayName = category === 'all' ? 'All categories' : titleCase(category);
  const selectedGroups = Object.values(selectedRows.reduce((result, row) => {
    result[row.category] ||= { name: row.category, value: 0, count: 0 };
    result[row.category].value += row.value;
    result[row.category].count += 1;
    return result;
  }, {})).sort((a, b) => b.value - a.value);
  const scopedModel = { ...model, rows: selectedRows, groups: selectedGroups, total: selectedTotal, topShare: selectedTotal ? selectedGroups[0]?.value / selectedTotal : 0, portfolioShare: model.total ? selectedTotal / model.total : 0 };
  const categoryMeaning = category.toUpperCase() === 'CCCP'
    ? 'Current customer / current product. Protect this base with renewal health checks, product enhancements, and early signs of value erosion.'
    : category.toUpperCase() === 'NCNP'
      ? 'New customer / new product. Treat this as a growth signal and verify repeat orders, adoption, and conversion into a durable account.'
      : 'The complete portfolio view. Select CCCP or NCNP to separate retention exposure from new-growth opportunity.';

  selectedCategoryTitle.textContent = displayName;
  selectedCategoryTotal.textContent = `${currency.format(selectedTotal)} value`;
  categoryNote.textContent = `${accountTotal} account${accountTotal === 1 ? '' : 's'} · ${currency.format(selectedTotal)} value`;
  categoryAiTitle.textContent = category === 'all' ? 'AI intelligence' : `${displayName} intelligence`;
  categoryAiDetail.textContent = categoryMeaning;
  customerRecords.innerHTML = selectedRows.length ? selectedRows.slice(0, 10).map((row) => `<div class="customer-record"><div><strong>${row.customer}</strong><span>${row.product || row.code || 'Customer record'}${row.billingDate ? ` · Invoice ${row.billingDate}` : ''}</span><small class="effective-status effective-status--${row.effectiveStatus.toLowerCase().replace(' ', '-')}">${row.effectiveStatus}${row.effectiveTo ? ` · valid to ${row.effectiveTo}` : ''}</small></div><b>${currency.format(row.value)}</b></div>`).join('') : '<div class="empty-state">No records found for this category.</div>';
  totalValue.textContent = selectedTotal ? currency.format(selectedTotal) : 'Text brief';
  customerCount.textContent = selectedRows.length ? accountTotal : '—';
  totalNote.textContent = selectedRows.length ? `${selectedRows.length} records in ${displayName}` : 'narrative file read';
  customerNote.textContent = model.customerHeader ? `unique ${displayName} accounts` : 'customer field not detected';
  concentrationRisk.textContent = scopedModel.portfolioShare ? `${Math.round(scopedModel.portfolioShare * 100)}%` : '—';
  concentrationNote.textContent = category === 'all' ? 'largest category share' : `${displayName} share of portfolio`;
  const activeRows = selectedRows.filter((row) => row.effectiveStatus === 'Active').length;
  lifecycleConfidence.textContent = model.hasLifecycleData ? 'High' : 'Provisional';
  lifecycleNote.textContent = selectedRows.length ? `${activeRows} of ${selectedRows.length} records active` : 'dates needed for precision';
  headlineInsight.textContent = selectedRows.length ? `${displayName} is in focus.` : 'No records found for this category.';
  headlineDetail.textContent = selectedRows.length ? `${displayName} represents ${Math.round(scopedModel.portfolioShare * 100)}% of portfolio value across ${accountTotal} customer account${accountTotal === 1 ? '' : 's'}.` : 'Choose another category to continue the intelligence read.';
  renderCategorySignals(scopedModel);
  renderRisks(scopedModel);
}

function setupCategorySelector(model) {
  const categories = [...new Set(model.rows.map((row) => row.category))].sort();
  categorySelect.innerHTML = '<option value="all">All categories</option>' + categories.map((category) => `<option value="${category}">${titleCase(category)}</option>`).join('');
  categorySelect.addEventListener('change', () => renderSelectedCategory(model, categorySelect.value));
  renderSelectedCategory(model, 'all');
}

function renderCategorySignals(model) {
  if (!model.groups.length) {
    categoryList.innerHTML = '<div class="empty-state">No tabular records were found. The narrative content is available in the evidence panel.</div>';
    return;
  }
  categoryList.innerHTML = model.groups.slice(0, 6).map((group) => {
    const share = model.total ? Math.round((group.value / model.total) * 100) : 0;
    const label = group.name.toUpperCase() === 'CCCP' ? 'Current customer / current product' : group.name.toUpperCase() === 'NCNP' ? 'New customer / new product' : titleCase(group.name);
    return `<div class="category-row"><div class="category-row__top"><strong>${titleCase(group.name)}</strong><span>${currency.format(group.value)}</span></div><div class="category-row__bar"><i style="width:${Math.max(share, 3)}%"></i></div><div class="category-row__bottom"><span>${label}</span><b>${share}% of value</b></div></div>`;
  }).join('');
}

function renderRisks(model) {
  const risks = [];
  const concentration = model.portfolioShare ?? model.topShare;
  if (concentration >= .7) risks.push({ level: 'High', title: 'Revenue is concentrated', detail: `${titleCase(model.groups[0].name)} represents ${Math.round(concentration * 100)}% of the portfolio value. A slowdown here would move the whole portfolio.` });
  if (model.groups.some((group) => group.name.toUpperCase() === 'CCCP')) risks.push({ level: 'Watch', title: 'Existing-product renewal exposure', detail: 'CCCP revenue signals reliance on current customers and current products. Review enhancement cadence before the base starts to soften.' });
  if (model.groups.some((group) => group.name.toUpperCase() === 'NCNP')) risks.push({ level: 'Opportunity', title: 'New motion needs conversion proof', detail: 'NCNP is a growth signal. Track repeat billing and time-to-second-order to distinguish durable demand from one-off wins.' });
  const expiredRows = model.rows.filter((row) => row.effectiveStatus === 'Expired');
  const agingRows = model.rows.filter((row) => row.ageYears !== null && row.ageYears >= 3);
  if (expiredRows.length) risks.push({ level: 'High', title: 'Effective-dated products have expired', detail: `${expiredRows.length} invoice record${expiredRows.length === 1 ? '' : 's'} fall outside the product validity period. Confirm renewal, replacement, or revenue leakage.` });
  if (agingRows.length) risks.push({ level: 'Watch', title: 'Products show enhancement aging', detail: `${agingRows.length} product record${agingRows.length === 1 ? '' : 's'} has had no enhancement for three or more years at invoice date. Prioritize roadmap review.` });
  if (!model.hasLifecycleData) risks.push({ level: 'Data gap', title: 'Product lifecycle dates are missing', detail: 'Add product, billing date, launch/enhancement date, and owner to give every product an evidence-based risk rating.' });

  const customers = Object.values(model.rows.reduce((result, row) => {
    result[row.customer] ||= { name: row.customer, value: 0, rows: [], categories: new Set() };
    result[row.customer].value += row.value;
    result[row.customer].rows.push(row);
    result[row.customer].categories.add(row.category.toUpperCase());
    return result;
  }, {})).sort((a, b) => b.value - a.value);
  customers.forEach((customer) => {
    const expired = customer.rows.filter((row) => row.effectiveStatus === 'Expired').length;
    const aging = customer.rows.filter((row) => row.ageYears !== null && row.ageYears >= 3).length;
    const isCccp = customer.categories.has('CCCP');
    const isNcnp = customer.categories.has('NCNP');
    let level = 'Low';
    let detail = `${currency.format(customer.value)} value across ${customer.rows.length} record${customer.rows.length === 1 ? '' : 's'}; active product evidence is present.`;
    if (expired) {
      level = 'High';
      detail = `${currency.format(customer.value)} value includes ${expired} expired effective-dated product record${expired === 1 ? '' : 's'}. Confirm renewal or replacement immediately.`;
    } else if (aging || isCccp) {
      level = 'Watch';
      detail = `${currency.format(customer.value)} value is exposed to ${isCccp ? 'the existing-customer/current-product base' : 'product enhancement aging'}${aging ? `; ${aging} product record${aging === 1 ? '' : 's'} need roadmap review` : ''}.`;
    } else if (isNcnp) {
      level = 'Opportunity';
      const productName = customer.rows.find((row) => row.product)?.product || 'the first product';
      const priority = customer.value < 100000 ? 'This is a small foothold and needs an expansion plan.' : 'This is an early growth account and needs a conversion plan.';
      detail = `${currency.format(customer.value)} from ${productName}. ${priority} Target a second order within the next billing cycle, map two adjacent products or services for cross-sell, and schedule a customer review to identify capacity, quality, or turnaround needs TAPAMS can solve.`;
    }
    risks.push({ level, title: `${customer.name} · customer rating`, detail });
  });
  riskList.innerHTML = risks.map((risk) => `<div class="risk-row"><span class="risk-badge risk-badge--${risk.level.toLowerCase().replace(' ', '-')}">${risk.level}</span><div><strong>${risk.title}</strong><p>${risk.detail}</p></div></div>`).join('');
}

function renderEvidence(model) {
  const checks = [
    ['Customer identity', Boolean(model.customerHeader), model.customerHeader || 'Add customer name or code'],
    ['Category motion', Boolean(model.headers.find((header) => header.toLowerCase().includes('categor'))), 'CCCP / NCNP can be compared'],
    ['Product identity', Boolean(model.productHeader), model.productHeader || 'Needed for product-level sickness'],
    ['Billing timeline', Boolean(model.billingHeader), model.billingHeader || 'Needed for effective-dated lifecycle'],
    ['Effective dates', Boolean(model.effectiveFromHeader || model.effectiveToHeader), model.effectiveFromHeader || model.effectiveToHeader || 'Needed for active / expired status'],
    ['Enhancement history', Boolean(model.enhancementHeader), model.enhancementHeader || 'Needed for aging risk'],
  ];
  evidenceGrid.innerHTML = checks.map(([label, available, detail]) => `<div class="evidence-item"><span class="evidence-item__mark ${available ? 'evidence-item__mark--yes' : ''}">${available ? '✓' : '!'}</span><div><strong>${label}</strong><span>${detail}</span></div></div>`).join('');
}

async function init() {
  try {
    const file = await getFile();
    if (!file) throw new Error('No file found');
    sourceName.textContent = file.name;
    const parsed = await readFile(file);
    const model = buildDataModel(parsed.rows, parsed.text);
    const categoryName = model.groups[0]?.name || 'the uploaded material';
    totalValue.textContent = model.total ? currency.format(model.total) : 'Text brief';
    customerCount.textContent = model.rows.length ? new Set(model.rows.map((row) => row.customer)).size : '—';
    concentrationRisk.textContent = model.topShare ? `${Math.round(model.topShare * 100)}%` : '—';
    lifecycleConfidence.textContent = model.hasLifecycleData ? 'High' : 'Provisional';
    totalNote.textContent = model.rows.length ? `${model.rows.length} records read` : 'narrative file read';
    customerNote.textContent = model.customerHeader ? 'unique customers found' : 'customer field not detected';
    concentrationNote.textContent = model.groups.length ? `${titleCase(categoryName)} is largest` : 'no value column detected';
    lifecycleNote.textContent = model.hasLifecycleData ? 'dated product evidence found' : 'dates needed for precision';
    recordCount.textContent = `${model.rows.length || 'Narrative'} ${model.rows.length === 1 ? 'record' : 'records'}`;
    headlineInsight.textContent = model.groups.length ? `${titleCase(categoryName)} is the first place to look.` : 'The narrative file is ready for interpretation.';
    headlineDetail.textContent = model.groups.length ? `It carries ${Math.round(model.topShare * 100)}% of recorded value. The dashboard has separated retention exposure from growth opportunity.` : 'The next read will extract commitments, dates, owners, risks, and dependencies from the document.';
    renderCategorySignals(model);
    renderRisks(model);
    renderEvidence(model);
    setupCategorySelector(model);
  } catch (error) {
    sourceName.textContent = 'No source file found';
    headlineInsight.textContent = 'Upload a requirement file to begin.';
    headlineDetail.textContent = 'Return to the upload page and choose the material you want TAPAMS to read.';
    categoryList.innerHTML = '<div class="empty-state">No file is available in this browser session.</div>';
    riskList.innerHTML = '<div class="empty-state">Upload a file to generate risk signals.</div>';
  }
}

init();
